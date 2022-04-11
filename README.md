# MailStore-FolderRights_Exchange_Sync_PS
Synchronize folder permissions between Exchange and Mailstore server via Powershell
Das Skript funktioniert nur auf Exchange Servern oder einem Server auf dem die Exchange-PowerShell-Snapins installiert sind.
Für Exchange-Server vor EX2013 muss das Snapin angepasst werden:
  - Exchange 2007
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin
  - Exchange 2010
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
   - Ab Exchange 2013
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Für Mailstore wird ein Administrator oder ein User benötig der auf die API zugreifen darf. Außerdem muss natürlich auf dem Mailstore-Server die API aktiviert werden.
Vorbereitungen des Mailstore-Servers: https://help.mailstore.com/de/server/Administration_API_-_Using_the_API
Vorbereitungen für die PowerShell: https://help.mailstore.com/en/server/PowerShell_API_Wrapper_Tutorial

Um die Funktionen zu prüfen kann der Schalter "$OnlyDisplayOut=$true" genutzt werden, dabei werden keine Befehle an die API gesendet und es erfolgt nur eine Ausgabe in der Konsole.
