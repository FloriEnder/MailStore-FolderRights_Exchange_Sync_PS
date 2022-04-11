#Conf etc:
$OUSuchPfade = "ou=OU,dc=CONTONSO,dc=LOCAL"
$Domain = "CONTONSO"
$MailstoreAdminUser = "admin"
$MailstoreUserPW = "MyPas.123!"
$MailstoreServer = "MailstoreServer" # or "192.168.X.X"
$MailstoreServerport = 8463
$OnlyDisplayOut = $true
Import-Module 'C:\Path\To\MS.PS.Lib.psd1'
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
#Ende Conf
$StartDate = (Get-Date)
$msapiclient = New-MSApiClient -Username $MailstoreAdminUser -Password $MailstoreUserPW -Server $MailstoreServer -Port $MailstoreServerport -IgnoreInvalidSSLCerts
$MailstoreUsers = (Invoke-MSApiCall $msapiclient "GetUsers").result
$useradd = 0
$userdel = 0
foreach ($MailstoreUser in $MailstoreUsers){
    if ($MailstoreUser.userName -ne "a_ender"){
    Write-Host "Start user -> " $MailstoreUser.userName
        #Auslesen von Berechtigungen auf die User aktuell Berechtigungen hat -> MailstoreFolderRights
        $MailstoreFolderRights = @()
        $rigths =  ((Invoke-MSApiCall $msapiclient "GetUserInfo" @{userName = $MailstoreUser.userName}).result).privilegesOnFolders
        foreach ($right in $rigths){
            $MailstoreFolderRights += $right.folder
        }
        #Auslesen auf welche Postfächer User noch Berechtigungen hat -> UserWithRightsOnMailbox
        $UserWithRightsOnMailbox = @()
        $ADuser = $Domain+"\"+$MailstoreUser.userName
        $canonicialname = Get-Mailbox -OrganizationalUnit $OUSuchPfade | Get-MailboxPermission | where {$_.user.tostring() -eq $ADuser} | Select-Object -Property Identity
        foreach ($name in $canonicialname.Identity){
            $DistinguishedName = @()
            $DistinguishedName = (Get-ADObject -Filter * -Properties CanonicalName -SearchBase $OUSuchPfade | where {$_.CanonicalName.tostring() -eq $name} | Select $_.DistinguishedName)
            $UserWithRightsOnMailbox += Get-ADUser -Filter * -SearchBase $OUSuchPfade | where {$_.DistinguishedName.tostring() -eq $DistinguishedName} | Select SamAccountName
            $DistinguishedName = @()
        }
        #Prüfen auf welche Ordner User noch Berechtigungen hat
        foreach ($MailStoreFolderUser in $MailstoreFolderRights){
            $found = $false
            if (($UserWithRightsOnMailbox.count -eq 0) -and ($MailStoreFolderUser -eq $MailstoreUser.userName)){
                $found = $true
            }else{
                if ($MailStoreFolderUser -eq $MailstoreUser.userName){
                    $found = $true
                }else{
                    foreach ($EXpostfach in $UserWithRightsOnMailbox.SamAccountName){
                        if ($MailStoreFolderUser -eq $EXpostfach){
                            $found = $true
                            break
                        }
                    }
                }                
            }         
            if ($found -eq $false){
                #UserLöschen
                if ($OnlyDisplayOut){
                    $userdel += 1
                    write-Host "Delete user" $MailStoreFolderUser
                }else{
                    $userdel += 1
                    write-Host "Delete user" $MailStoreFolderUser
                    $returnDelete = Invoke-MSApiCall $msapiclient "SetUserPrivilegesOnFolder" @{userName = $MailstoreUser.userName;folder = $MailStoreFolderUser;privileges = "none"}
                    $returnDelete.statusCode 
                }                
            }
        }
        #Ordner mit leserechte zu User hinzufügen
        foreach ($EXpostfach in $UserWithRightsOnMailbox.SamAccountName){
            $found = $false    
            foreach ($UserWithRightsOnFolder in $MailstoreFolderRights){
                if ($UserWithRightsOnFolder -eq $EXpostfach){
                    $found = $true
                break
                }    
            }
            if ($found -eq $false){
            #User hinzufügen
            if ($OnlyDisplayOut){
                $useradd += 1
                write-Host "Add user" $EXpostfach
            }else{
                $useradd += 1
                write-Host "Add user" $EXpostfach
                $returnAdd = Invoke-MSApiCall $msapiclient "SetUserPrivilegesOnFolder" @{userName = $MailstoreUser.userName;folder = $EXpostfach;privileges = "read"}
                $returnADD.statusCode
            }            
            }
        }
        Write-Host "End user ->" $MailstoreUser.userName
    }
}
$Time = NEW-TIMESPAN –Start $StartDate –End (Get-Date)
Write-Host "User add -> " $useradd
Write-Host "User del -> " $userdel
Write-Host "Time needed" $Time.Minutes "min" $Time.Seconds "s"
