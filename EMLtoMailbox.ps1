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

Function GenEml {
param(
[Parameter(Mandatory=$true)][int]$number
)

## Ignore certificate prompts if using self-signed certificates

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

#################################

# start the count at 0

$i=0

# start the loop to create and upload message via EWS

do{

$rootPath = "C:\dev\"
$rootUsers= get-content "C:\dev\newMailServerUsers_Top624.txt"
$attachments= "C:\dev\eml\emailAttachments\"
$general= "C:\dev\scenarios\Scenario2\General\"
$outputDir= "C:\dev\output\"


$mailbox= getFrom everyone
$To= getTo everyone
$CCTo = getTo everyone
$BCCTo = getTo everyone
$Subject = getSubject general
$body= getBody general
$attach = getAttachment 
$SaveinFolder = "Inbox"
#$date = (GetRandomDateBetween -StartDate (Get-Date).AddDays(-2014) -EndDate (Get-Date).AddDays(-419))
$date = (GetRandomDateBetween -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date))

$dllpath = "C:\dev\Microsoft.Exchange.WebServices.dll"
Import-Module $dllpath
        

                                   
            ##assemble all recipients to one array because we will need to inject this message via EWS to all the would-be recipient mailboxes
            [array]$allMbx = [array]$to + [array]$CCTo + [array]$BCCTo

            write-host -ForegroundColor yellow " writing the email to the following mailboxes: "
            $allMbx

            ##  adding email to sent-items of the person who sent the email
             $mailbox 
                        ## Set Exchange Version
                        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
                        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
                        #$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
                        $uri=[system.URI] ""
                        $service.url = $uri
                        $userName=""
                        $password= ""
                        $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
                        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$mailbox);
                        $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mailbox)
                        $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                        $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
                        $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                        $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                        $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note" } | select ID,Displayname
                        $DI = $allFolders | ? {$_.DisplayName -eq "Sent Items"} 
                        $Folderid=$DI.ID
                                                                                                               
                        $Email = new-object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
                        $extSubmit = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0057,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime) 
                        $extDelivery = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3590,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime) 
                        $PR_Flags = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3591, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
                                             
                        $Email.SetExtendedProperty($extSubmit,$date)
                        $Email.SetExtendedProperty($extDelivery, $date)
                        $Email.SetExtendedProperty($PR_Flags,"1") 

                        $Email.From = $mailbox
                        $Email.Body = $body
                        $Email.ToRecipients.Add($to) | Out-Null
                        $Email.CcRecipients.Add($CCTo) | Out-Null
                        $Email.BccRecipients.Add($BCCTo) | Out-Null
                        $Email.Subject = $Subject
                        $eml = $Email.Subject
                        $Email.Attachments.AddFileAttachment($attach) | Out-Null
                        Write-host -ForegroundColor yellow "count: $i" 
                        try {

                                   
                                    $error.Clear()
                                    $Email.Save($Folderid)
                                    
                                    Write-host -ForegroundColor yellow " - Upload item to Sent Items folder of $mailbox, Subject: $eml Date: $date, using the endpoint $URI"
                            }
                                    Catch{$Error}
                                    #pause
                             

##upload the email to all the recipients

            foreach($mbx in $allMbx){
                        $mbx
                        ## Set Exchange Version
                        $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1
                        $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
                        #$uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
                        $uri=[system.URI] "https://ewslb.a360labs.com/ews/exchange.asmx"
                        $service.url = $uri
                        $userName="a360-mbximper@a360labs.com"
                        $password= "Letmein001--"
                        $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
                        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$mbx);
                        $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$mbx)
                        $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                        $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
                        $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                        $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                        $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note" } | select ID,Displayname
                        $DI = $allFolders | ? {$_.DisplayName -eq $SaveinFolder} 
                        $Folderid=$DI.ID
                        
                               if([string]::IsNullOrEmpty($Folderid.ChangeKey) -eq $true){

                                                    $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note" } | select ID,Displayname
                                                    $parentFolderID = $allFolders | ? {$_.DisplayName -eq "inbox"} 
                                                    $newFolder = new-object Microsoft.Exchange.WebServices.Data.Folder($service)
                                                    $newFolder.DisplayName = $SaveinFolder
                                                    $newFolder.FolderClass = "IPF.Note"
                                                    $newfolder.Save($parentFolderID.Id)
                                                    $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                                                    $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note" } | select ID,Displayname
                                                    $DI = $allFolders | ? {$_.DisplayName -eq $SaveinFolder} 
                                                    $Folderid=$DI.Id

                                                                                        }
                        $Email = new-object Microsoft.Exchange.WebServices.Data.EmailMessage($service)
                        $extSubmit = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0057,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime) 
                        $extDelivery = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3590,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime) 
                        $PR_Flags = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(3591, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer);  
                                             
                        $Email.SetExtendedProperty($extSubmit,$date)
                        $Email.SetExtendedProperty($extDelivery, $date)
                        $Email.SetExtendedProperty($PR_Flags,"1") 

                        $Email.From = $mailbox
                        $Email.Body = $body
                        $Email.ToRecipients.Add($to) | Out-Null
                        $Email.CcRecipients.Add($CCTo) | Out-Null
                        $Email.BccRecipients.Add($BCCTo) | Out-Null
                        $Email.Subject = $Subject
                        $Email.Attachments.AddFileAttachment($attach) | Out-Null
                        try {

                                   
                                    $error.Clear()
                                    $Email.Save($Folderid)
                                    
                                    Write-host -ForegroundColor yellow "count: $i - Upload item to $mbx, Subject: $eml Date: $date, using the endpoint $URI"
                            }
                                    Catch{$Error}
                                    #pause
                             }
                                        
            
             $i++
             #sleep -Milliseconds 10000
                                
    }
    Until(($i -eq $number) -match $true )
     }        
            
     GenEml                              

                                   