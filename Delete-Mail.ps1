# THIS SCRIPT WILL RUN A COMPLICANCE SEARCH ON ALL MAILBOXES AND DELETE IDENTIFED MAIL

# https://docs.microsoft.com/en-gb/office365/securitycompliance/search-for-and-delete-messages-in-your-organization

# Get login credentials 
$UserCredential = Get-Credential 
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.compliance.protection.outlook.com/powershell-liveid -Credential $UserCredential -Authentication Basic -AllowRedirection 
Import-PSSession $Session -AllowClobber -DisableNameChecking 
$Host.UI.RawUI.WindowTitle = $UserCredential.UserName + " (Office 365 Security & Compliance Center)" 



# CHANGE THE BELOW VERIABLES TO MATCH SEARCH   

# name of the compliance search (can be anything)
$SearchName = ""
#Content Query 
$ContentQuery = 'sent>=07/03/2017 AND sent<=07/05/2017 AND subject:"open this attachment!!"'


###################################################################

# search for the email in all mailboxes 
New-ComplianceSearch -Name $SearchName -ExchangeLocation all -ContentMatchQuery $ContentQuery
Start-ComplianceSearch -Identity $SearchName

# wait a minute and view the results
Get-ComplianceSearch -Identity $SearchName
Get-ComplianceSearch -Identity $SearchName | Format-List

# delete the messages if the results look right
New-ComplianceSearchAction -SearchName $SearchName -Purge -PurgeType SoftDelete

# check when it is completed
Get-ComplianceSearchAction -Identity $SearchName

