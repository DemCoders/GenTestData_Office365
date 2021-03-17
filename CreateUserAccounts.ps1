$emea = "OU=EMEAUsers,DC=EMEA,DC=authA360Labs,DC=com"
$na = "OU=NAEastUsers,DC=NAEAST,DC=authA360Labs,DC=com"
#$mbxpwd = Get-Credential
$csv = import-csv C:\dev\EmeaAllUsers.csv
$path = $emea

$csv | % {

$displayName = $_.DisplayName
$name = $_.Name
$PrimaryEmailAddress = $_.PrimarySmtp
$upn = $_.upn
    
    New-ADUser -Name $displayName  -DisplayName $displayName  -EmailAddress $PrimaryEmailAddress `
    -AccountPassword $mbxpwd.password -ChangePasswordAtLogon:$false -PasswordNeverExpires:$true `
    -Path $path -Enabled:$True -UserPrincipalName $upn
    
}



