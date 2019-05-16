### Sync AD with 365/AzureAD   ###




try {    
    
    $UserCredential = Get-Credential
    Connect-MsolService -Credential $UserCredential    
        
}
catch {
        
    try {
        $UserCredential = Get-Credential
        Connect-MsolService -Credential $UserCredential
    }
    catch {
    
    Write-Host "ERROR: Failed to log into Office 365 Admin Account" -ForegroundColor Red
    
    }
}




try {
    Invoke-Command -ComputerName DC1.espc.local -ScriptBlock {Import-Module ADSync; Start-ADSyncSyncCycle -PolicyType Initial}
    Write-host "Syncing to Office 365 - This might take a while..." -ForegroundColor Green
    Start-Sleep -Seconds 300
    Set-MsolUser -UserPrincipalName $e365 -UsageLocation GB 
    Set-MsolUserLicense -UserPrincipalName $e365 -AddLicenses espcuk:O365_BUSINESS_PREMIUM
    Set-MsolUserPrincipalName -UserPrincipalName $e365 -NewUserPrincipalName $email
    Write-host "Still syncing Office 365 Account..." -ForegroundColor Green
    Start-Sleep -Seconds 60
}
catch {

    Write-Host "ERROR: Failed to sync to Office 365" -ForegroundColor Red
    break

}

### Add user to Shared Mailboxes ###


try{

Write-host "Connecting to Exchange Online..." -ForegroundColor Green
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking

}
catch{

    Write-Host "ERROR: Failed to connect to Exchange Online" -ForegroundColor Red

}
try {
    Write-host "Adding user to Shared Mailboxes..." -ForegroundColor Green
    switch ($Depart) {
        0 {
      

        }
        1 {
            Add-MailboxPermission -Identity "Support@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Support@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false

        }
        2 {
            Add-MailboxPermission -Identity "PropertyPress@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "PropertyPress@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false

        }
        3 {
            Add-MailboxPermission -Identity "Dunfermline@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Dunfermline@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false

        }
        4 {

        }
        5 {
            Add-MailboxPermission -Identity "Finance@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Finance@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
            Add-MailboxPermission -Identity "Finance@espclettings.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Finance@espclettings.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
            Add-MailboxPermission -Identity "Finance@altislegal.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Finance@altislegal.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false

        }
        6 {
            Add-MailboxPermission -Identity "Lettings@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Lettings@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
            Add-MailboxPermission -Identity "Landlord@espclettings.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Landlord@espclettings.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false

        }
        7 {
            Add-MailboxPermission -Identity "Marketing@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Marketing@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false

        }
        8 {
            Add-MailboxPermission -Identity "Mediasales@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Mediasales@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false

        }
        9 {


        }
        10 {
            Add-MailboxPermission -Identity "UAT@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "UAT@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
            
            Add-MailboxPermission -Identity "Solutions.Issues@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Soutions.Issues@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false 

        }
        11 {
            Add-MailboxPermission -Identity "IT@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "IT@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false

        }

        12 {
            Add-MailboxPermission -Identity "UAT@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "UAT@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
            

        }

        13 {
            Add-MailboxPermission -Identity "Support@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Support@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
            Add-MailboxPermission -Identity "My.ESPC@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "My.ESPC@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
            Add-MailboxPermission -Identity "property@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Property@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
            Add-MailboxPermission -Identity "Showroom.Admin@espc.com" -user "$email" -AccessRights "fullaccess" -InheritanceType "all"
            Add-RecipientPermission -Identity "Showroom.Admin@espc.com" -Trustee "$email" -AccessRights "sendas" -Confirm:$false
        }
  
    }

    Write-host "SUCCESS" -ForegroundColor Green
    Remove-PSSession $Session
}
catch {

    Write-Host "ERROR: Failed add user to shared mailbox" -ForegroundColor Red

}

$confirmation = Read-Host "Please Confirm if you would like to continue to request accounts for other applications (y/n)"
if ($confirmation -eq 'y') {

  Write-Host "CONFIRMED" -ForegroundColor Green
  Invoke-Expression -Command "C:\Scripts\NewUser\CreateUser04.ps1"

} else {

Write-Host "CANCELLED" -ForegroundColor Red
    break
}