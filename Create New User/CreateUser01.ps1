Set-ExecutionPolicy RemoteSigned




### Import CSV and add Variables ###
$CSVUN = read-host 'Enter Username of New User'

Try {

$ImportCSV = import-csv \\espc-vmdata1\Scratch\NewUser\$CSVUN.csv

}

Catch{

Write-Host "ERROR: The name you have chosen does not exist. Please check the name and location of the csv - \\espc-vmdata1\Scratch\NewUser - Script 1, Line[6-17]" -ForegroundColor Red
break

}

### Add Variables ###
ForEach ($item in $ImportCSV){

    # Assign each property value to a simpler named variable
    

    $first = $item.("FirstName") 
    $last = $item.(“LastName”)
    $Description = $item.(“Job Description")
    $mfirst = $item.(“Mfirst”)
    $mlast = $item.(“mlast”)
    $Office = $item.("Department")
    $Company = $item.("Company")
    $Phone = $item.("Phone")
    $PlainPassword = $item.("Password")
    $MBPort = $item.("MemberP")
    $o365Li = $item.("o365Li")
    $Mobile = $item.("Mobile")
    $Slack = $item.("Slack")
    $Cap = $item.("Cap")
    $MemberP = $item.("MemberP")
    $BDPAC = $item.("BDPAC")
    $GITHUB = $item.("GITHUB")
    $DPC = $item.("DPC")
    $Fdesk = $item.("Fdesk")
    $Phone = $item.("Phone")
    $Trello = $item.("Trello")
    $Jira = $item.("Jira")
    $Umbraco = $item.("Umbraco")
    $AltisAC = $item.("AltisAC")
    $Twin = $item.("Twin")
    $SPC = $item.("SPC")
    $TESTMB = $item.("TESTMB")
    $HMRC = $item.("HMRC")
    $Bank = $item.("Bank")
    $SME = $item.("SME")
    $Crystal = $item.("Crystal")
    $AccessP = $item.("AccessP")
    $Sage = $item.("Sage")
    $BaseC = $item.("BaseC")
    $MSDN = $item.("MSDN")
    $ROI = $item.("ROI")
    $GDFP = $item.("GDFP")
    $ADSRV = $item.("ADSRV")
    $HRTool = $item.("HRTool")
    $Depart = $item.("Depart")

    Write-Output $item 
}

        $IT = 'it@espc.com'
        $email = $first + '.' + $last + '@espc.com'
        $OUDepartment = $Office
        $Manager = $mfirst.substring(0, 1) + $mlast
        $MManager = $mfirst + ' ' + $mlast
        $un = $first.substring(0, 1) + $last
        $Domain = ',DC=espc,DC=local'
        $pw = $PlainPassword | ConvertTo-SecureString -AsPlainText -Force
        $Name = $first + ' ' + $last
        $OU = 'OU=$OUDepartment' + $Domain
        $365 = $first + '.' + $last + '@espcuk.onmicrosoft.com'
        $e365 = $un + '@espcuk.onmicrosoft.com'
        $unlocal = $un + '@espc.local'
        $memail = $mfirst + '.' + $mlast + '@espc.com'



$confirmation = Read-Host "Please Confirm the user detailes are correct (y/n)"
if ($confirmation -eq 'y') {

  Write-Host "CONFIRMED" -ForegroundColor Green
  Invoke-Expression -Command "C:\Scripts\NewUser\CreateUser02.ps1"

} else {

Write-Host "CANCELLED" -ForegroundColor Red
    break
}