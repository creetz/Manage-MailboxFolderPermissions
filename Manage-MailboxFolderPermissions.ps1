<#
    .SYNOPSIS
    Set and Remove ExchangeMailboxFolderPermissions with advanced options 
   
   	Christian Reetz 
    (Updated by Christian Reetz to support Exchange Server 2013/2016)
	
	THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
	RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
	23.03.2016
	
    .DESCRIPTION

    Set and Remove ExchangeMailboxFolderPermissions with different options
	
	PARAMETER Mailbox
    Die Mailbox auf die Berechtigungsänderungen angewandt werden, bzw. für die ein Ordnerreport erzeugt werden soll. 
	
	PARAMETER User
	Der zu berechtigende Nutzer. Username, E-Mail-Adresse, Identity
	
	PARAMETER LogFile
	Pfad der Logdatei z.B. c:\temp\username.txt
	
	PARAMETER ReferenceUser
    Der Referenznutzer, welcher als Referenz für die Berechtigungsstruktur dient. Optional und relevant für –mode: addlikereferenceuser
		
	PARAMETER AccessRights
	Das Zugriffsrechtslevel, welches gesetzt werden soll für alle Ordner. Optional und relevant für –mode: addallfolders und changerights
    Folgende Level sind verfügbar:

    PublishingEditor	Read, Write, Delete, Create subfolders
    Editor			  Read, Write, Delete 
    FolderVisible	   Show only the folder without content or other rights
    Contributor		 Only create elements, no read, change or other rights
    Owner			   fullaccess incl. manage rigths
    Reviewer		    Read
    Author			  Read, Write, Delete (only own elements)
    PublishingAuthor	Read, Write, Delete (only own elements), Create subfolders
    NonEditAuthor       Read, Write (only own elements) 
    
	PARAMETER mode
	-mode: addlikereferenceuser
    Hierbei wird eine vorhandene Berechtigungsstruktur, anhand eines Referenzbenutzers ermittelt und dann auf den –User (siehe Unten) angewandt  
    -mode: addallfolders
    Hier werden Berechtigungen auf alle Ordner und Unterordner für den –User (siehe Unten)  angewandt.
    -mode: removeallfolders
    Hier werden Berechtigungen von allen Ordnern und Unterordnern für den –User  (siehe Unten) aufgehoben.
    -mode: report
    Hier wird ein Report erzeugt, welcher Alle Rechte auflistet [für das definierte Postfach –Mailbox]
    -mode: getdelegate
    Die definierten Stellvertretter werden angezeigt
    -mode: removedelegate
    Alle definierten Stellvertretter werden gelöscht
    -mode: changerights
    Die Berechtigungen eines vorhandenen Benutzers anpassen (z.B. von Editor zu Autor)
    -mode: reportsendpermissions
    Hier wird ein Report erzeugt, welcher Alle Senderechte auflistet [für das definierte Postfach –Mailbox]

    PARAMETER targetrootfolder
    Mit diesem Parameter kann ein Ordner definieren, welcher ausschließlich (aller enthaltener Unterordner) für sämtliche Berechtigungsänderungen 
    verwendet wird.
      
	EXAMPLES
    .\Manage-MailboxFolderPermissions.ps1 -Mailbox User1 -User User2 -Mode removeallfolders -LogFile c:\temp\remove-User2.txt
                                                   
    .\Manage-MailboxFolderPermissions.ps1 -Mailbox User1 -User User2 -Mode addlikereferenceuser -ReferenceUser User3
                                                   
    .\Manage-MailboxFolderPermissions.ps1 -Mailbox User1 -User User2 -Mode addallfolders -LogFile c:\temp\remove-User2.txt -accessrights editor
                                                   
    .\Manage-MailboxFolderPermissions.ps1 -Mailbox User1 -User * -Mode getdelegate
                                                   
    .\Manage-MailboxFolderPermissions.ps1 -Mailbox User1 -User * -Mode removedelegate
                                                   
    .\Manage-MailboxFolderPermissions.ps1 -Mailbox User1 -User User2 -Mode changerights -accessrights author
                                                   
    .\Manage-MailboxFolderPermissions.ps1 -Mailbox User1 -mode reportsendpermissions
                                                   
    .\Manage-MailboxFolderPermissions.ps1 -Mailbox User1 -mode report
    #>

Param (
	[Parameter(Mandatory=$true)] $Mailbox,
	[Parameter(Mandatory=$true)] $Mode,
    [Parameter(Mandatory=$false)] $User,
    [Parameter(Mandatory=$false)] $LogFile,
    [Parameter(Mandatory=$false)] $ReferenceUser,
    [Parameter(Mandatory=$false)] $AccessRights,
    [Parameter(Mandatory=$false)] $targetrootfolder
)

#Mailboxfolders which will be ignored...
$exclusions = @("/Sync Issues", 
                 "/Sync Issues/Conflicts", 
                 "/Sync Issues/Local Failures", 
                 "/Sync Issues/Server Failures", 
                 "/Synchronisierungsprobleme/Serverfehler",
                 "/Synchronisierungsprobleme/Lokale Fehler",
                 "/Synchronisierungsprobleme/Konflikte",
                 "/Synchronisierungsprobleme",
                 "/Kontakte/GAL Contacts",
                 "/Kontakte/Recipient Cache",
                 "/Kontakte/Skype for Business-Kontakte",
                 "/Kontakte/{",
                 "/Contacts/GAL Contacts",
                 "/Contacts/Recipient Cache",
                 "/Contacts/Skype for Business-Kontakte",
                 "/Contacts/{*", #doesn't work
                 "/Einstellungen für QuickSteps",
                 "/Conversation Action Settings",
                 "/Aufgezeichnete Unterhaltungen",
                 "/Journal", #allways englisch
                 "/Calendar Logging", #allways englisch
                 "/Junk-E-Mail", #allways englisch
                 "/Deletions", #allways englisch
                 "/Purges",  #allways englisch
                 "/Versions", #allways englisch
                 "/Recoverable Items", #allways englisch
                 "/Working Set", #allways englisch
                 "/Verschlüsselt" #Customized Folder
                 "/Ordner/Verschlüsselt" #Customized Folder
                 )

if (!($user))
{
    $user = "*"
}

if (($user -eq "*") -and ($Mode -ne "report"))
{
    $user = Read-Host "define targetuser (f.E. -user User1)?"
}

function seterrordatared{
    if ($host.PrivateData)
    {
        $host.PrivateData.errorforegroundcolor="Red"
    }
}
function seterrordatadarkcyan{
    if ($host.PrivateData)
    {
        $host.PrivateData.errorforegroundcolor="DarkCyan"
    }
}


#$mailboxfolders = @(Get-MailboxFolderStatistics $Mailbox | Select FolderId, FolderPath)
$mailbox = Get-Mailbox $mailbox 
#$mailbox = $mailbox | where-object {$_.alias -eq "$($mailbox.alias)"}
$mailboxalias = $mailbox.identity

if ($targetrootfolder)
{
    Write-Host -ForegroundColor White "Specified RootFolder: $targetrootfolder"
    $mailboxfolders = @(Get-MailboxFolderStatistics "$Mailboxalias" | Where {!($exclusions -icontains $_.FolderPath) -and ($_.FolderPath -like "/$targetrootfolder*")} | Select FolderId, FolderPath)
}
else
{
    $mailboxfolders = @(Get-MailboxFolderStatistics "$Mailboxalias" | Where {!($exclusions -icontains $_.FolderPath)} | Select FolderId, FolderPath)
}

$nl = $([System.Environment]::NewLine)

### Ask for logging ########################################################################################################################################

if (-not ($LogFile))
{
$loggingenable = Read-Host "enable logging for changes? (y/n)"
    if (($loggingenable -like "yes") -or ($loggingenable -like "y") -or ($loggingenable -like "j") -or ($loggingenable -eq "like") -or ($logging -eq "z"))
    {
        Write-Host $nl
        $logFile = Read-Host "specifie logging-path (f.E.: c:\temp\rightschanges_user1_20151101.log)"
    }
}

#### Prepare LogFile Headline ##############################################################################################################################

if ($mode -eq "addlikereferenceuser")
{
    if ($referenceuser)
    {}
    else
    {
        $referenceuser = Read-Host "Specified an reference user, to copy permissions"
    }
    
    seterrordatadarkcyan
    $log = "Set permissions like " + $referenceuser + " for user " + $user + " in mailbox " + $mailbox + $nl
    Write-Host "Set permissions like $referenceuser for user $user in mailbox $mailbox" 
}

if ($mode -eq "addallfolders")
{
    if ($accessrights)
    {}
    else
    {
        $accessrights = Read-Host "Specified the rights-level (f.e. PublishingEditor)"
    }
    
    $log = "Set permissions $accessrights on all folders for user $user in mailbox $mailbox" + $nl 
    Write-Host "Set permissions on all folder for user $user in mailbox $mailbox"
}

if ($mode -eq "removeallfolders")
{
    $log = "Remove Rights in Mailbox " + $mailbox + " for User " + $user + $nl 
    Write-Host -ForegroundColor Yellow $folderpath 
}

if ($mode -eq "report")
{
    Write-Host "Checking $mailbox for permissions"
}

if ($mode -eq "getdelegate")
{
    Get-Mailbox $mailbox | Get-CalendarProcessing | select ResourceDelegates | fl
}

if ($mode -eq "removedelegate")
{
    Get-Mailbox $mailbox | Set-CalendarProcessing -ResourceDelegates $null
}

if ($mode -eq "changerights")
{
        if ($accessrights)
        {}
        else
        {
            $accessrights = Read-Host "Specified the rights-level (f.e. PublishingEditor)"
        }
    seterrordatadarkcyan
    $log = "Change rights in Mailbox " + $mailbox + " for User " + $user + $nl 
    
}

#### Set foreach folder in mailbox  #########################################################################################################################

foreach ($mailboxfolder in $mailboxfolders)
{
    $folder = $mailboxfolder.FolderId
    $folderpath = $mailboxfolder.FolderPath
    $identity = "$mailboxalias"+":" + "$folder"

    #IF Referenceuser was specified
    if ($mode -eq "addlikereferenceuser")
    {
        Write-Host -ForegroundColor Yellow $folderpath 
        $log += $folderpath + $nl
      
            try
            {
                $referencefolderpermission = "";
                $accessrights = "";
                $referencefolderpermission = Get-MailboxFolderPermission -Identity $identity -User $referenceuser | select name, AccessRights
                $accessrights = $referencefolderpermission.AccessRights

                if ($referencefolderpermission)
                {
                    Add-MailboxFolderPermission -Identity $identity -User $User -AccessRights $AccessRights -Confirm:$false -ErrorAction STOP
                    Write-Host -ForegroundColor Green "done!"
                    $log += "Add-Rights: $AccessRights" + $nl
                }
            }
            catch
            {
                Write-Warning $_.Exception.Message
            }
      
    }

    #IF ChangeRights was specefied
    if ($mode -eq "changerights")
    {
       Write-Host -ForegroundColor Yellow $folderpath
       $log += $folderpath + $nl
       
            try
            {
            $folderwithpermission = "";
            $folderwithpermission = Get-MailboxFolderPermission -Identity $identity -User $user | select name, AccessRights
            
                 if ($folderwithpermission)
                 {
                 Remove-MailboxFolderPermission -Identity $identity -User $User -Confirm:$false -ErrorAction STOP
                 Add-MailboxFolderPermission -Identity $identity -User $User -Confirm:$false -AccessRights $AccessRights -ErrorAction STOP
                 Write-Host -ForegroundColor Green "done!"
                 $log += "Change-Rights: $AccessRights from $($folderwithpermission.AccessRights)!" + $nl
                 }
            }
            catch
            {
                 Write-Warning $_.Exception.Message
            } 
    }

    #IF Remove was specefied

    if ($mode -eq "removeallfolders")
    {
        $log += $folderpath + $nl
        Write-Host -ForegroundColor Yellow $folderpath
            
            if (Get-MailboxFolderPermission -Identity $identity -User $user -ErrorAction SilentlyContinue)
            {
                try
                {
                    Remove-MailboxFolderPermission -Identity $identity -User $User -Confirm:$false -ErrorAction STOP
                    if ($mode -eq "removeallfolders")
                    {
                        Write-Host -ForegroundColor Green "Removed!"
                    }
                    $log += "Removed!" + $nl
                }
                catch
                {
                    Write-Warning $_.Exception.Message
                }
            }
    }

    #IF Alluserfolders was specified
    if ($mode -eq "addallfolders")
    {
        
        Write-Host -ForegroundColor Yellow $folderpath 
        $log += $folderpath + $nl
      if (Get-MailboxFolderPermission -Identity $identity -User $user -ErrorAction SilentlyContinue)
            {
                try
                {
                    Remove-MailboxFolderPermission -Identity $identity -User $User -Confirm:$false -ErrorAction STOP
                    if ($mode -eq "removeallfolders")
                    {
                        Write-Host -ForegroundColor Green "Removed!"
                    }
                    $log += "Removed!" + $nl
                }
                catch
                {
                    Write-Warning $_.Exception.Message
                }
            }


            try
            {
                Add-MailboxFolderPermission -Identity $identity -User $User -Confirm:$false -AccessRights $AccessRights -ErrorAction STOP
                Write-Host -ForegroundColor Green "done!"
                $log += "Add-Rights!" + $nl
            }
            catch
            {
                Write-Warning $_.Exception.Message
            }
    }

    
    #IF Report was specefied

    if ($mode -eq "report")
    {
            if (($user -eq "*") -or ($user -eq "$false"))
            {
                $permissions = Get-MailboxFolderPermission -Identity $identity | select name, user, AccessRights
            }
            if ($user -ne "*" -and ($user))
            {
                seterrordatadarkcyan
                $permissions = Get-MailboxFolderPermission -Identity $identity -User $user | select name, user, AccessRights 
            }
            $useraccess = "$($permissions.user)"
            $accesslevel = "$($permissions.AccessRights)" 
            
            Write-Host -ForegroundColor Yellow $folderpath 
            $permissions | ft user, accessrights
            #Write-Host -ForegroundColor Green $useraccess
            #Write-Host -ForegroundColor Cyan $accesslevel
            
            $log += $folderpath + $nl
            $log += $useraccess + $nl
            $log += $accesslevel + $nl 

    }
}

#### Außerhalb der Schleife

    #IF reportsendpermissions was specified
    if ($mode -eq "reportsendpermissions")
    {
        #$log += "$($mailbox.name)" + $nl
      if ($LogFile){
      Write-Host "Logfile-Option not available!"
      $LogFile = $false
      }


            try
            {
                Write-Host ""
                $sendas = $mailbox | Get-ADPermission | where {$_.ExtendedRights -like "*send-as*"}
                $SendOnBehalf = $mailbox | select-object GrantSendOnBehalfTo
                Write-Host -ForegroundColor Yellow "SendAs"
                $sendas
                Write-Host ""
                $SendOnBehalf
                
                #$log += "SendAs" + $nl
                #$log += "$($sendas.user.rawidentity)" + $nl
                #$log += "SendOnBehalfTo" + $nl
                #$log += "$($SendOnBehalf.GrantSendOnBehalfTo.Rdn.escapedname)" + $nl
            }
            catch
            {
                Write-Warning $_.Exception.Message
            }
    }


#### Log-File generation ##############################################################################################################################

if ($LogFile){
    $log > $LogFile
}

seterrordatared