
$allmbx = get-mailbox -ResultSize unlimited | ? {$_.identity -notLike "*Discovery*" -and $_.identity -notlike "*Demetrios*" -and $_.alias -notlike "a360*"}

$emea = "OU=EMEAUsers,DC=EMEA,DC=authA360Labs,DC=com"
$na = "OU=NAEastUsers,DC=NAEAST,DC=authA360Labs,DC=com"
$paths += @($emea,$na)

$emeaDC = "emeadc1.emea.autha360labs.com"
$naDC = "nadc1.naeast.autha360labs.com"


$allmbx | % {

$error.Clear()
write-host "Processing......"
write-host -ForegroundColor yellow "disabling $_ local mailbox "
Disable-Mailbox -Identity $_.identity -Confirm:$false
$path = get-random $paths -Count 1
Set-ADUser $_.Alias -Enabled:$false
write-host -ForegroundColor yellow "re-enabling  $_ in $path  "

try{
    if($path -like "OU=EMEA*"){
                
                $linkedAccount = $_.Alias + "@emea.autha360labs.com"
                Connect-Mailbox -Identity $_.displayName -LinkedDomainController $emeaDC -LinkedMasterAccount $linkedAccount -Database $_.database -User $_.displayName
            
            }
            Else {

                $linkedAccount = $_.Alias + "@naeast.autha360labs.com"
                Connect-Mailbox -Identity $_.displayName -LinkedDomainController $naDC -LinkedMasterAccount $linkedAccount -Database $_.database -User $_.displayName

                }
        }
    Catch{throw "something went wrong" ; $error}

    
}



