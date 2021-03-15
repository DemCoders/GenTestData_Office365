Function getBody {
param(
[ValidateSet("General", "BanterPurchases", "SideBusiness", "ThanksEmails")]
[Parameter(Mandatory=$true)][string]$type
)
If($type -like "*general*"){
$BcontentPath= "C:\dev\eml\emailDatabase.csv"
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.body
return $Bbody
}
elseif($type -like "*banterpurchases*"){
$BcontentPath= "C:\dev\scenarios\Scenario2\BanterPurchases\BanterPurchases.csv"
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.body
return $Bbody
}
elseif($type -like "*sidebusiness*"){
$BcontentPath= "C:\dev\scenarios\Scenario2\sidebusiness\SideBusiness.csv"
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.body
return $Bbody
}
elseif($type -like "*thanksemails*"){
$BcontentPath= "C:\dev\scenarios\Scenario2\thanksemails\ThanksEmails.csv"
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.body
return $Bbody
}
else{  Throw "General", "BanterPurchases", "SideBusiness", "ThanksEmails"}
}

Function getSubject {
param(
[ValidateSet("General", "BanterPurchases", "SideBusiness", "ThanksEmails")]
[Parameter(Mandatory=$true)][string]$type
)
If($type -like "*general*"){
$BcontentPath= "C:\dev\eml\emailDatabase.csv"
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.subject
return $Bbody
}
elseif($type -like "*banterpurchases*"){
$BcontentPath= "C:\dev\scenarios\Scenario2\BanterPurchases\BanterPurchases.csv"
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.subject
return $Bbody
}
elseif($type -like "*sidebusiness*"){
$BcontentPath= "C:\dev\scenarios\Scenario2\sidebusiness\SideBusiness.csv"
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.subjects
return $Bbody
}
elseif($type -like "*thanksemails*"){
$BcontentPath= "C:\dev\scenarios\Scenario2\thanksemails\ThanksEmails.csv"
$Bcontent= import-csv $BcontentPath
$Bbody= get-random $Bcontent.subjects
return $Bbody
}
else{  Throw "General", "BanterPurchases", "SideBusiness", "ThanksEmails"}
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
           return Get-Random -Minimum $StartDate.Ticks -Maximum $EndDate.Ticks | Get-Date -Format "ddd, dd MMM yyyy HH':'00':'00 'GMT'"
        }
    }

Function getTo{
param(
[ValidateSet("coworker", "SMBOwner", "Investigator", "everyone")]
[Parameter(Mandatory=$true)][string]$type
)

If($type -like "*coworker*"){
$content= $rootUsers | ? {$_.Roles -like "*co-worker*"}
$content = $content.WindowsEmailAddress -join ","
return $content
}
elseif($type -like "*investigator*"){
$Tcontent= $rootUsers | ? {$_.Roles -like "*investigator*"}
$Tcontent = $Tcontent.WindowsEmailAddress -join ","
return $Tcontent
}
elseif($type -like "*SMBOwner*"){
$Tcontent= $rootUsers | ? {$_.Roles -like "*smbowner*"}
$Tcontent = $Tcontent.WindowsEmailAddress -join ","
return $Tcontent
}
elseif($type -like "*everyone*"){
$count= get-random -Minimum 1 -Maximum 12
$Tcontent= get-random $rootUsers -count $count
#$Tcontent = $Tcontent.WindowsEmailAddress -join ","
return $Tcontent
}
else{  Throw "Input type, coworker, SMBOwner, Investigator, or everyone"}
}

Function getFrom{
param(
[ValidateSet("coworker", "SMBOwner", "Investigator", "everyone")]
[Parameter(Mandatory=$true)][string]$type
)

If($type -like "*coworker*"){
$content= $rootUsers | ? {$_.Roles -like "*co-worker*"}
$content= get-random $rootUsers -Minimum 1 
$content = $content.WindowsEmailAddress
return  $content
}
elseif($type -like "*investigator*"){
$content= $rootUsers | ? {$_.Roles -like "*investigator*"}
$content = $content.WindowsEmailAddress
return  $content
}
elseif($type -like "*smbowner*"){
$content= $rootUsers | ? {$_.Roles -like "*smbowner*"}
$content = $content.WindowsEmailAddress
return  $content
}
elseif($type -like "*everyone*"){
$content= get-random $rootUsers -count 1
#$content = $content.WindowsEmailAddress
return  $content
}
else{  Throw "Input type, coworker, SMBOwner, Investigator, everyone"}
}

Function getAttachment{

$attachmentPaths = (get-childitem $attachments).FullName
return get-random $attachmentPaths


}


#  please specify the path to the csv files that will contain the data.  Look inside the ReoAppt function and specify the connection URL and creds for accessing the mailboxes. 
#  Please also specify the path to the EWS DLL in the ReoAppt function
###########################################################################

$rootPath = "C:\dev\"
$rootUsers= get-content "C:\dev\newMailServerUsers_Top624.txt"
$attachments= "C:\dev\eml\emailAttachments\"
$general= "C:\dev\scenarios\Scenario2\General\"
$outputDir= "C:\dev\output\"

########################################################################


function ReoAppt {
param(
[Parameter(Mandatory=$true)][int]$number
)

        ## Ignore Certificate Prompts
add-type @"
using System.Net;
using System.Security.Cryptography.X509Certificates;
public class TrustAllCertsPolicy : ICertificatePolicy {
    public bool CheckValidationResult(
        ServicePoint srvPoint, X509Certificate certificate,
        WebRequest request, int certificateProblem) {
        return true;
    }
}
"@
[System.Net.ServicePointManager]::CertificatePolicy = New-Object TrustAllCertsPolicy 


$i=0

do { 
$mailbox= getFrom everyone
$mandatoryTo= getTo everyone
$optionalTo = getTo everyone
$Subject = getSubject general
$body= getBody general



[DateTime]$starttime= GetRandomDateBetween -StartDate (Get-Date).AddDays(-14) -EndDate (Get-Date)
$enddateNum= 30, 60, 120 | get-random -count 1 
[DateTime]$endtime= $starttime.AddMinutes($enddateNum)
$reminderTimeNum= -15, -30| get-random -count 1 
[DateTime]$reminderTime= $starttime.AddMinutes($reminderTimeNum)


######################################################
    ##  Please specify the path to the EWS DLL

$dllpath = "C:\dev\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath

######################################################

## Set Exchange Version
$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
#$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"

#################################################

##  please specify the creds and EWS service connection point (or mailbox server URL)

$uri=[system.URI] ""
$service.url = $uri
$userName=""
$password= ""

##################################################

$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$Mailbox);

$NewAppointment = new-object Microsoft.Exchange.WebServices.Data.Appointment($service)
$NewAppointment.Recurrence = New-Object Microsoft.Exchange.WebServices.Data.Recurrence+WeeklyPattern($starttime, 1, ($starttime.DayOfWeek));

$NewAppointment.Body = $body
#$NewAppointment.Start = $starttime
#$NewAppointment.End = $endtime
$NewAppointment.Subject = $Subject
$NewAppointment.Location = "Tennis club"
$NewAppointment.ReminderDueBy = $reminderTime
$NewAppointment.Recurrence.StartDate = $starttime
$NewAppointment.Recurrence.NumberOfOccurrences = 10



$mandatoryTo | % {

$NewAppointment.RequiredAttendees.Add($_) | Out-Null

        }

$optionalTo | % {

$NewAppointment.OptionalAttendees.Add($_) | Out-Null

        }
try{
$NewAppointment.Save()
}
Catch{$error}

write-host -ForegroundColor green "appt created for $mailbox on $starttime "

$i++


}
Until(($i -eq $number) -match $true )

}

ReoAppt