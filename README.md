![example](https://github.com/creetz/Manage-MailboxFolderPermissions/blob/master/pic1.png)

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

    PublishingEditor Read, Write, Delete, Create subfolders
    Editor			     Read, Write, Delete 
    FolderVisible	   Show only the folder without content or other rights
    Contributor		   Only create elements, no read, change or other rights
    Owner			       fullaccess incl. manage rigths
    Reviewer		     Read
    Author			     Read, Write, Delete (only own elements)
    PublishingAuthor Read, Write, Delete (only own elements), Create subfolders
    NonEditAuthor    Read, Write (only own elements) 
    
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
