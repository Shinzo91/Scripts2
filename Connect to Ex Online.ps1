Set-ExecutionPolicy RemoteSigned
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking




Get-Recipient | where {$_.EmailAddresses -match "fsenquiries@espc.com"} | fL Name, RecipientType,emailaddresses


F11E

76966