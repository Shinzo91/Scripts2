###   Create AD User    ###
try {
Write-Host "Creating Active Directory User..." -ForegroundColor Green
Import-Module activedirectory
new-ADUser $name -Enabled $true -AccountPassword $pw -Path "OU=$OUDepartment,DC=espc,DC=local" -UserPrincipalName $unlocal -Department $Office -Description $Description -DisplayName $name -OfficePhone $Phone -Title $Description -SamAccountName $un -GivenName $first -Surname $last -passwordneverexpires 0 -ChangePasswordAtLogon 1; 
Get-ADUser $un | Get-ADObject | Set-ADObject -Add @{"manager" = "CN=$MManager,OU=$OUDepartment,DC=espc,DC=local"}
set-aduser -Identity $un
Write-Host "SUCCESS" -ForegroundColor Green
}

catch {
Write-Host "ERROR: Failed to create Active Directory user, please ensure the new uses and the manager are in the same OU - Script 2" -ForegroundColor Red
break
}
###   Create Exchange on-prem mailbox    ###
try{
Write-Host "Creating Exchange Mailbox..." -ForegroundColor Green
$password = get-content \\espc-vmdata1\Library\IT\onrepm.txt | convertto-securestring
$ExchangeCreds = new-object -typename System.Management.Automation.PSCredential -argumentlist "espc\exadmin", $password
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://exchange.espc.local/PowerShell/ -Authentication Kerberos -Credential $ExchangeCreds
Import-PSSession $Session
Enable-RemoteMailbox -Identity $un -RemoteRoutingAddress $email
Remove-PSSession $Session
Write-Host "SUCCESS" -ForegroundColor Green
}
catch{
Write-Host "ERROR: Failed to create Exchange Mailbox, please ensure all services are running on the Exchange server - Script 2" -ForegroundColor Red
break
}

### Updating User Details and adding to AD groups ###
try {
Write-Host "Adding user to AD groups..." -ForegroundColor Green
switch ($Depart) {
    0 {
        Add-ADPrincipalGroupMembership -Identity:$un -MemberOf:"CN=SP - ESPC Staff,CN=Users,DC=espc,DC=local" -Server:"DC1.espc.local"
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"Altis Legal" 

    }
    1 {
        
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'Ad Reset' -Members $un
        add-ADGroupMember 'AltisRW' -Members $un
        add-ADGroupMember 'BS Users' -Members $un
        add-ADGroupMember 'CS Users' -Members $un
        add-ADGroupMember 'ESPC ReadyNas Access' -Members $un
        add-ADGroupMember 'Member Engagement Team' -Members $un
        add-ADGroupMember 'Member Firm Comms' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    2 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'Design' -Members $un
        add-ADGroupMember 'Design Studio Users' -Members $un
        add-ADGroupMember 'Design Team' -Members $un
        add-ADGroupMember 'ESPC ReadyNas Access' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    3 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'CS Users' -Members $un
        add-ADGroupMember 'Dunfermline Users' -Members $un
        add-ADGroupMember 'ESPC ReadyNas Access' -Members $un
        Set-ADUser -Identity:"$un" -City:"Fife" -Country:"GB" -PostalCode:"KY12 7EA" -Server:"DC1.espc.local" -State:"Dunfermline" -StreetAddress:"15 New Row"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    4 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'FAC Users' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    5 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'Finance Users' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 


    }
    6 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'CS Users' -Members $un
        add-ADGroupMember 'Lettings Users' -Members $un
        add-ADGroupMember 'Users' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    7 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'MP Users' -Members $un
        add-ADGroupMember 'SPC Scotland file access' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    8 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'DCA Users' -Members $un
        add-ADGroupMember 'ESPC ReadyNas Access' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    9 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'MM Users' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    10 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'SOL Users' -Members $un
        add-ADGroupMember 'Solutions Team' -Members $un
        add-ADGroupMember 'VPNUsers' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 

    }
    11 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'AD Reset' -Members $un
        add-ADGroupMember 'ESPC ReadyNas Access' -Members $un
        add-ADGroupMember 'IT Users' -Members $un
        add-ADGroupMember 'OnyxSupport' -Members $un
        add-ADGroupMember 'VPNUsers' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 
         

    }

    12 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'SOL Users' -Members $un
        add-ADGroupMember 'Solutions Team' -Members $un
        add-ADGroupMember 'VPNUsers' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3DF" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"90a George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd"
        Add-DistributionGroupMember -Identity "SolutionsTeam@espc.com" -Member "$email" -Confirm:$false
        Add-MailboxPermission -Identity "Uat@espc.com" -User "$email" -AccessRights "FullAccess" -InheritanceType "All"
        Set-ADUser -Identity "&un" -Replace @{extensionAttribute6 = "SOL"}

    }

    13 {
        add-ADGroupMember 'SP - ESPC Staff' -Members $un
        add-ADGroupMember 'CS Users' -Members $un
        Set-ADUser -Identity:"$un" -City:"Edinburgh" -Country:"GB" -PostalCode:"EH2 3ES" -Server:"DC1.espc.local" -State:"Midlothian" -StreetAddress:"107 George Street"
        Set-ADUser $un -Company:"ESPC (UK) Ltd" 
        

    }
  
}

add-ADGroupMember 'zScaler' -Members $un
set-aduser $un -Enabled $true
Write-Host "SUCCESS" -ForegroundColor Green
}

catch{
Write-Host "ERROR: Failed - Script 2" -ForegroundColor Red
break
}

###   Syncing Domain Controllers     ###
try {
Write-Host "Syncing Domain Controllers..." -ForegroundColor Green
$DomainControllers = Get-ADDomainController -Filter *
ForEach ($DC in $DomainControllers.Name) {
    Write-Host "Processing for "$DC -ForegroundColor Green
    If ($Mode -eq "ExtraSuper") {
        REPADMIN /kcc $DC REPADMIN /syncall /A /e /q $DC
    }
    Else {
        REPADMIN /syncall $DC "dc=espc,dc=local" /d /e /q
    }
}
Write-Host "SUCCESS" -ForegroundColor Green
}

catch{
Write-Host "ERROR: Failed to Sync Domain Controllers  - Script 2" -ForegroundColor Red

}


$confirmation = Read-Host "Please Confirm if you would like to continue to sync user to Office 365 (y/n)"
if ($confirmation -eq 'y') {

  Write-Host "CONFIRMED" -ForegroundColor Green
  Invoke-Expression -Command "C:\Scripts\NewUser\CreateUser03.ps1"

} else {

Write-Host "CANCELLED" -ForegroundColor Red
    break
}

