Get-AdmPwdPassword

get-service | where-object Status -eq 'Stopped' | Select-Object status,Name,Displayname | export-csv C:\Scripts\services2.csv


export-cvs C:\Scripts


$PSVersionTable 


Verb-Noun

Get-Verb

get-verb -verb s*

cls - clear host

get-service -Name M* -ComputerName DC1 

Start-Transcript -Path C:\Scripts\Transcript01.txt
Get-Command *counter*



## Displays when the PC was rebooted ##

Get-EventLog -log System -Newest 1000 | where-object eventid -eq '1074' | ft MachineName, UserName, TimeGenerated -AutoSize


## Networking with PowerShell ##

Get-Command *IP*

Get-NetIPAddress

Get-NetIPConfiguration

Get-DnsClientCache #shows the dns cache on a pc#


#Create a new network drive #
New-SmbMapping -LocalPath S: -RemotePath \\ \\

Test-NetConnection 8.8.8.8 -TraceRoute


Get-PSProvider


#Copy call files from one directory to another 

Copy-Item C:\Scripts -Destination 'C:\Scripts Two Test' -Recurse -Verbose

Get-ChildItem 'C:\Scripts Two Test' -Recurse


# Add and Printer (location name reqired)
Add-Printer -ConnectionName \\location\\printer 



# Active Directory # 

Get-ADUser -Identity snowlan -Properties *

Search-ADAccount -LockedOut | select name


#Show call acount is AD that have been locked out#

Search-ADAccount -AccountDisabled | select name


Get-ADComputer -Filter * | Out-File -FilePath C:\Scripts\ADComp.txt


Get-Command *group*

Get-ADGroup *


#Get all members of a AD group#
Get-ADGroupMember -Identity Solutions | select name

Get-ADUser -Property Name, Department -Filter



#Remoting with Powershell#


#Enable Remoting#
Enable-PSRemoting

#Set Permissions#
Get-PSSessionConfiguration
Set-PSSessionConfiguration


#Modify Windows Firewall#
Set-NetFirewallRule

Get-NetFirewallRule | where DisplayName -Like "*Windows Management Instrumentation*" | select DisplayName,name,enabled

Get-NetFirewallRule | where DisplayName -Like "*Windows Management Instrumentation*" | Set-NetFirewallRule -Enabled True -Verbose

Enter-PSSession -ComputerName 


Get-Variable


#set location users home#
sl $HOME

Write-Output "The home location for this PC is " $HOME 

gcm Get-PSSession


#Scripts#

Get-ExecutionPolicy



