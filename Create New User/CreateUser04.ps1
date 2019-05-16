$Applications = @("Slack","Freshdesk","Horizon","Member Portal","BDP","Github","Umbraco","Altis","MSDN")


for ($i=0; $i -lt $Applications.length; $i++){

$confirmation = Read-Host "Would you like to request a " $Applications[$i]  " account? (y/n)"
if ($confirmation -eq 'y') {
    
    $From = "$IT"
    $To = "support@espc.com"
    $Subject = "New " + $Applications[$i] +" Account"
    $Body = "Hi,

Can you please create a " + $Applications[$i] + " account for $first $last, their email address is $email.

Thanks

IT Department"
    $SMTPServer = "outlook.office365.com"
    $SMTPPort = "587"
    
    
    try{

    Send-MailMessage -From $From -to $To -Subject $Subject `
        -Body $Body -SmtpServer $SMTPServer -port $SMTPPort -UseSsl `
        -Credential $UserCredential

    Write-Host "Requesting"  $Applications[$i]  " Account" -ForegroundColor Green

    }

    catch {
    
    Write-Host "ERROR: Failed to send email request for " + $Applications[$i] + "account" -ForegroundColor Red
    
    }
    
}
else {
    
    Write-Host $Applications[$i] "account not required" -ForegroundColor Red
    
}


}

Write-Host ".................FINISHED.........................." -ForegroundColor Green