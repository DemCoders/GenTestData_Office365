Function getBody {
param(
[ValidateSet("e2e", "m2m", "e2m", "m2e")]
[Parameter(Mandatory=$true)][string]$type
)
If($type -like "*e2e*"){
$BcontentPath= $e2e
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.body
return $Bbody
}
elseif($type -like "*m2m*"){
$BcontentPath= $m2m
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.body
return $Bbody
}
elseif($type -like "*e2m*"){
$BcontentPath= $e2m
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.body
return $Bbody
}
elseif($type -like "*m2e*"){
$BcontentPath= $m2e
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.body
return $Bbody
}
else{  Throw "e2e", "m2m", "e2m", "m2e"}
}

Function getSubject {
param(
[ValidateSet("e2e", "m2m", "e2m", "m2e")]
[Parameter(Mandatory=$true)][string]$type
)
If($type -like "*e2e*"){
$BcontentPath= $e2e
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.subject
return $Bbody
}
elseif($type -like "*m2m*"){
$BcontentPath= $m2m
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.subject
return $Bbody
}
elseif($type -like "*e2m*"){
$BcontentPath= $e2m
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.subject
return $Bbody
}
elseif($type -like "*m2e*"){
$BcontentPath= $m2e
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.subject
return $Bbody
}
else{  Throw "e2e", "m2m", "e2m", "m2e"}
}

Function GetRandomDateBetween{
        <#
        .EXAMPLE
        Get-RandomDateBetween -StartDate (Get-Date) -EndDate (Get-Date).AddDays(-15)
        #>
        [Cmdletbinding()]
        param(
            [parameter(Mandatory=$True)][DateTime]$StartDate,
            [parameter(Mandatory=$True)][DateTime]$EndDate
            )

        process{
           return Get-Random -Minimum $StartDate.Ticks -Maximum $EndDate.Ticks | Get-Date -Format "ddd, dd MMM yyyy HH':'mm':'ss 'GMT'"
        }
    }

Function getTo{
param(
[ValidateSet("e2e", "m2m", "e2m", "m2e")]
[Parameter(Mandatory=$true)][string]$type
)

If($type -like "*e2e*"){
$content= $rootUsers | ? {$_.Roles -like "*co-worker*"}
$content = $content.WindowsEmailAddress -join ","
return $content
}
elseif($type -like "*m2m*"){
$Tcontent= $rootUsers | ? {$_.Roles -like "*manager*"}
$Tcontent = $Tcontent.WindowsEmailAddress -join ","
return $Tcontent
}
elseif($type -like "*e2m*"){
$Tcontent= $rootUsers | ? {$_.Roles -like "*manager*"}
$Tcontent = $Tcontent.WindowsEmailAddress -join ","
return $Tcontent
}
elseif($type -like "*m2e*"){
$Tcontent= $rootUsers | ? {$_.Roles -like "*co-worker*"}
$Tcontent = $Tcontent.WindowsEmailAddress -join ","
return $Tcontent
}
else{  Throw "Input type, coworker, SMBOwner, Investigator, or everyone"}
}

Function getFrom{
param(
[ValidateSet("e2e", "m2m", "e2m", "m2e")]
[Parameter(Mandatory=$true)][string]$type
)
If($type -like "e2e"){
$content= $rootUsers | ? {$_.Roles -like "*co-worker*"}
$content= get-random $content -count 1
return $content.windowsEmailAddress
}
elseif($type -like "m2m"){
$content= $rootUsers | ? {$_.Roles -like "*manager*"}
$content= get-random $content -count 1
return $content.windowsEmailAddress
}
elseif($type -like "e2m"){
$content= $rootUsers | ? {$_.Roles -like "*co-worker*"}
$content= get-random $content -count 1
return $content.windowsEmailAddress
}
elseif($type -like "m2e"){
$content= $rootUsers | ? {$_.Roles -like "*manager*"}
$content= get-random $content -count 1
return $content.windowsEmailAddress
}
else{  Throw "Input e2e", "m2m", "e2m", "m2e"}
}


Function getAttachment{

$attachmentPaths = (get-childitem $attachments).FullName
return get-random $attachmentPaths


}



#####

Function GenEML {
param(
[Parameter(Mandatory=$true)][int]$number
)
$i=0
do{
$rootPath = "D:\dev\scenarios4\"
$rootUsers= import-csv "D:\dev\scenario4\Scenario4_users.csv"
$e2e= "D:\dev\scenario4\E2E\E2E.csv"
$m2m= "D:\dev\scenario4\M2M\m2m.csv"
$e2m= "D:\dev\scenario4\E2M\E2M.csv"
$m2e= "D:\dev\scenario4\M2E\M2E.csv"

$outputDir= "D:\dev\scenario4\EMLs\E2E\"

#####


#$date= "Date: " + (GetRandomDateBetween -StartDate (Get-Date).AddDays(-1068) -EndDate (Get-Date).AddDays(-33))
$from= getFrom E2E
$to= getTo E2E
$subject = getSubject E2E
$body= getBody E2E

#$attachment= getAttachment
#$emailFileName = ($from  + "_" + (get-random)+".eml").tostring()
#$emailFileFullPath = $emlFilePath+$emailFileName

###  Submit message to file system
#$mailMessage = New-Object System.Net.Mail.MailMessage
#$mailMessage.From = New-Object System.Net.Mail.MailAddress("$from")
#$mailMessage.To.Add($to)
#$mailMessage.Subject = $subject
#$mailMessage.Body = $body
#$mailMessage.Attachments.Add($attachment)
#$smtpClient = New-Object System.Net.Mail.SmtpClient
#$smtpClient.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::SpecifiedPickupDirectory;
#$smtpClient.PickupDirectoryLocation = "C:\dev\EMLfiles\";
#$smtpClient.Send($mailMessage);
#$smtpClient.Dispose()
#$mailMessage.Dispose()


####  submit message for sending
$mailMessage = New-Object System.Net.Mail.MailMessage
$mailMessage.From = New-Object System.Net.Mail.MailAddress("$from")
$mailMessage.To.Add($to)
$mailMessage.Subject = $subject
$mailMessage.Body = $body
$mailMessage.date.$date
#$mailMessage.date = $date
#$mailMessage.Attachments.Add($attachment)

$smtpClient = New-Object System.Net.Mail.SmtpClient
#$SMTPServer = “a360labs-com.mail.protection.outlook.com”
#$SMTPClient = New-Object Net.Mail.SmtpClient($SMTPServer,25)


#$SMTPServer = "smtp.office365.com"
$smtpClient.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::SpecifiedPickupDirectory;
$smtpClient.PickupDirectoryLocation = "$outputDir";
$smtpClient.Send($mailMessage);

Write host " Now Processing $i mail messages `n FROM: $from `n TO: $to `n SUBJECT: $Subject `n BODY: $body "


$i++


}
Until(($i -eq $number) -match $true )

}


geneml