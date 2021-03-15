function Connect-Exchange
{ 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] $MailboxName
		#[Parameter(Position=1, Mandatory=$true)] [System.Management.Automation.PSCredential]$Credentials
    )  
 	Begin
		 {
		## Load Managed API dll  
		###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
		$EWSDLL = "C:\dev\Microsoft.Exchange.WebServices.dll"
        #$EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
		
        if (Test-Path $EWSDLL)
		    {
		    Import-Module $EWSDLL
		    }
		else
		    {
		    "$(get-date -format yyyyMMddHHmmss):"
		    "This script requires the EWS Managed API 1.2 or later."
		    "Please download and install the current version of the EWS Managed API from"
		    "http://go.microsoft.com/fwlink/?LinkId=255472"
		    ""
		    "Exiting Script."
		    $exception = New-Object System.Exception ("Managed Api missing")
			throw $exception
		    } 
  
		## Set Exchange Version  
		$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1  
		  
		## Create Exchange Service Object  
		$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)  
		  
		## Set Credentials to use two options are availible Option1 to use explict credentials or Option 2 use the Default (logged On) credentials  
		  
		#Credentials Option 1 using UPN for the windows Account  
		#$psCred = Get-Credential  
		#$creds = New-Object System.Net.NetworkCredential($Credentials.UserName.ToString(),$Credentials.GetNetworkCredential().password.ToString())  
        $userName=""
        $password= ""
        $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
        $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$MailboxName);
		#$service.Credentials = $creds      
		#Credentials Option 2  
		#service.UseDefaultCredentials = $true  
		 #$service.TraceEnabled = $true
		## Choose to ignore any SSL Warning issues caused by Self Signed Certificates  
		  
		## Code From http://poshcode.org/624
		## Create a compilation environment
		$Provider=New-Object Microsoft.CSharp.CSharpCodeProvider
		$Compiler=$Provider.CreateCompiler()
		$Params=New-Object System.CodeDom.Compiler.CompilerParameters
		$Params.GenerateExecutable=$False
		$Params.GenerateInMemory=$True
		$Params.IncludeDebugInformation=$False
		$Params.ReferencedAssemblies.Add("System.DLL") | Out-Null

$TASource=@'
  namespace Local.ToolkitExtensions.Net.CertificatePolicy{
    public class TrustAll : System.Net.ICertificatePolicy {
      public TrustAll() { 
      }
      public bool CheckValidationResult(System.Net.ServicePoint sp,
        System.Security.Cryptography.X509Certificates.X509Certificate cert, 
        System.Net.WebRequest req, int problem) {
        return true;
      }
    }
  }
'@ 
		$TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
		$TAAssembly=$TAResults.CompiledAssembly

		## We now create an instance of the TrustAll and attach it to the ServicePointManager
		$TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
		[System.Net.ServicePointManager]::CertificatePolicy=$TrustAll

		## end code from http://poshcode.org/624
		  
		## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use  
		  
		#CAS URL Option 1 Autodiscover  
		$service.AutodiscoverUrl($MailboxName,{$true})
        $scpCheck = $service.url.OriginalString.ToString()
            if([string]::IsNullOrEmpty($scpCheck) -eq $true){
                write-host "No local mailbox found, checking Office 365"
                $uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
                $service.url = $uri
                $userName="a360ingestor@a360labs.com"
                $password= "Letmein001--"
                $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $userName, $password
                $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$MailboxName);
                $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
                
                try{
                    $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                    }
                    Catch{ throw "The SMTP address has no mailbox associated with it."}
           
          Write-host ("Using EndPoint: " + $Service.url)

            }
              
		Write-host ("Using EndPoint: " + $Service.url)   
		   
		#CAS URL Option 2 Hardcoded  
		  
		#$uri=[system.URI] "https://casservername/ews/exchange.asmx"  
		#$service.Url = $uri    
		  
		## Optional section for Exchange Impersonation  
		  
		#$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName) 
		if(!$service.URL){
			throw "Error connecting to EWS"
		}
		else
		{		
			return $service
		}
	}
}


function Create-Contact 
{ 
    [CmdletBinding()] 
    param( 
    	[Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
 		[Parameter(Position=1, Mandatory=$true)] [string]$DisplayName,
		[Parameter(Position=2, Mandatory=$true)] [string]$FirstName,
		[Parameter(Position=3, Mandatory=$true)] [string]$LastName,
		[Parameter(Position=4, Mandatory=$true)] [string]$EmailAddress
		#[Parameter(Position=5, Mandatory=$false)] [string]$CompanyName,
		#[Parameter(Position=7, Mandatory=$false)] [string]$Department,
		#[Parameter(Position=8, Mandatory=$false)] [string]$Office,
		#[Parameter(Position=9, Mandatory=$false)] [string]$BusinssPhone,
		#[Parameter(Position=10, Mandatory=$false)] [string]$MobilePhone,
		#[Parameter(Position=11, Mandatory=$false)] [string]$HomePhone,
		#[Parameter(Position=12, Mandatory=$false)] [string]$IMAddress,
		#[Parameter(Position=13, Mandatory=$false)] [string]$Street,
		#[Parameter(Position=14, Mandatory=$false)] [string]$City,
		#[Parameter(Position=15, Mandatory=$false)] [string]$State,
		#[Parameter(Position=16, Mandatory=$false)] [string]$PostalCode,
		#[Parameter(Position=17, Mandatory=$false)] [string]$Country,
		#[Parameter(Position=18, Mandatory=$false)] [string]$JobTitle,
		#[Parameter(Position=19, Mandatory=$false)] [string]$Notes,
		#[Parameter(Position=20, Mandatory=$false)] [string]$Photo,
		#[Parameter(Position=21, Mandatory=$false)] [string]$FileAs,
		#[Parameter(Position=22, Mandatory=$false)] [string]$WebSite,
		#[Parameter(Position=23, Mandatory=$false)] [string]$Title,
		#[Parameter(Position=24, Mandatory=$false)] [string]$Folder,
		#[Parameter(Position=25, Mandatory=$false)] [string]$EmailAddressDisplayAs,
		#[Parameter(Position=26, Mandatory=$false)] [switch]$useImpersonation

		
    )  
 	Begin
	{
		#Connect
		$service = Connect-Exchange -MailboxName $MailboxName
		if($useImpersonation.IsPresent){
			$service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
		}
		$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Contacts,$MailboxName)   
		if($Folder){
			$Contacts = Get-ContactFolder -service $service -FolderPath $Folder -SmptAddress $MailboxName
		}
		else{
			$Contacts = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)
		}
		if($service.URL){
			$type = ("System.Collections.Generic.List"+'`'+"1") -as "Type"
			$type = $type.MakeGenericType("Microsoft.Exchange.WebServices.Data.FolderId" -as "Type")
			$ParentFolderIds = [Activator]::CreateInstance($type)
			$ParentFolderIds.Add($Contacts.Id)
			$Error.Clear();
			$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)  
			$ncCol = $service.ResolveName($EmailAddress,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::DirectoryThenContacts,$true,$cnpsPropset);
			$createContactOkay = $false
			if($Error.Count -eq 0){
				if ($ncCol.Count -eq 0) {
					$createContactOkay = $true;	
				}
				else{
					foreach($Result in $ncCol){
						#if($Result.Contact -eq $null){
						#	Write-host "Contact already exists " + $Result.Mailbox.Name
						#	throw ("Contact already exists")
						#}
						#else{
							#if((Validate-EmailAddres -EmailAddress $EmailAddress)){
								if($Result.Mailbox.MailboxType -eq [Microsoft.Exchange.WebServices.Data.MailboxType]::Mailbox){
									$UserDn = Get-UserDN -Credentials $Credentials -EmailAddress $Result.Mailbox.Address
									$cnpsPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
									$ncCola = $service.ResolveName($UserDn,$ParentFolderIds,[Microsoft.Exchange.WebServices.Data.ResolveNameSearchLocation]::ContactsOnly,$true,$cnpsPropset);
									if ($ncCola.Count -eq 0) {  
										$createContactOkay = $true;		
									}
									#else
									#{
									#	Write-Host -ForegroundColor  Red ("Number of existing Contacts Found " + $ncCola.Count)
									#	foreach($Result in $ncCola){
									#		Write-Host -ForegroundColor  Red ($ncCola.Mailbox.Name)
									#	}
									#	throw ("Contact already exists")
									#}
								#}
							#}
							#else{
							#	Write-Host -ForegroundColor Yellow ("Email Address is not valid for GAL match")
							#}
						}
					}
				}
				if($createContactOkay){
					$Contact = New-Object Microsoft.Exchange.WebServices.Data.Contact -ArgumentList $service 
					#Set the GivenName
					$Contact.GivenName = $FirstName
					#Set the LastName
					$Contact.Surname = $LastName
					#Set Subject  
					$Contact.Subject = $DisplayName
					$Contact.FileAs = $DisplayName
					
                    #if([string]::IsNullOrEmpty($Title) -eq $false ){
				    #	$PR_DISPLAY_NAME_PREFIX_W = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3A45,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String);  
					#	$Contact.SetExtendedProperty($PR_DISPLAY_NAME_PREFIX_W,$Title)						
					#                }
					
                    $Contact.CompanyName = $CompanyName
					$Contact.DisplayName = $DisplayName
					$Contact.Department = $Department
					$Contact.OfficeLocation = $Office
					$Contact.CompanyName = $CompanyName
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::BusinessPhone] = $BusinssPhone
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::MobilePhone] = $MobilePhone
					$Contact.PhoneNumbers[[Microsoft.Exchange.WebServices.Data.PhoneNumberKey]::HomePhone] = $HomePhone
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business] = New-Object  Microsoft.Exchange.WebServices.Data.PhysicalAddressEntry
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].Street = $Street
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].State = $State
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].City = $City
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].CountryOrRegion = $Country
					$Contact.PhysicalAddresses[[Microsoft.Exchange.WebServices.Data.PhysicalAddressKey]::Business].PostalCode = $PostalCode
					$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1] = $EmailAddress
					if([string]::IsNullOrEmpty($EmailAddressDisplayAs)-eq $false){
						$Contact.EmailAddresses[[Microsoft.Exchange.WebServices.Data.EmailAddressKey]::EmailAddress1].Name = $EmailAddressDisplayAs
					} 
					$Contact.ImAddresses[[Microsoft.Exchange.WebServices.Data.ImAddressKey]::ImAddress1] = $IMAddress 
					$Contact.FileAs = $FileAs
					$Contact.BusinessHomePage = $WebSite
					#Set any Notes  
					$Contact.Body = $Notes
					$Contact.JobTitle = $JobTitle
					if($Photo){
						$fileAttach = $Contact.Attachments.AddFileAttachment($Photo)
						$fileAttach.IsContactPhoto = $true
					}
			   		$Contact.Save($Contacts.Id)				
					Write-Host ("Contact Created for mailbox $MailboxName")
				}
			}
		}
	}
}

if ((Test-Path C:\dev\Contacts_Data.csv) -eq $false -or (Test-Path C:\dev\newMailServerUsers_Top624.txt) -eq $false){
throw "missing key input files for mailboxes and contacts.  Please locate files and place them in the c:\dev dir"
    }

Write-host -ForegroundColor yellow " Starting contact creation process"

$contacts_data = import-csv C:\dev\Contacts_Data.csv
$mailbox_data = get-content C:\dev\newMailServerUsers_Top624.txt

$contacts = $contacts_data | get-random -Count 50


Foreach ($mbx in $mailbox_data){
    $i = 0
    #connect-exchange $mbx.UserPrincipalName
    write-host  -ForegroundColor yellow  "connecting to" + $mbx


    ForEach ($cnt in $contacts){
        $i++
        Create-Contact -MailboxName $mbx -DisplayName $cnt.DisplayName -FirstName $cnt.FirstName -LastName $cnt.LastName -EmailAddress $cnt.EmailAddress 
        write-host "loop count: " $i
        

        }


}