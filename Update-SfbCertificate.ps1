<#  
.SYNOPSIS  
Quickly and easily copy and refresh an existing Lync or Skype for Business server certificate.
No longer will you forget or misspell a SAN!

.DESCRIPTION  
This script acts as a front-end for the "Request-CsCertificate" commandlet in Lync and Skype for Business.
It reads the existing certificate for a given "type" using "Get-CsCertificate" & uses the values from it to populate the new certificate request.
It's also able to take a certificate just by its thumbnail and use that as the source of the new request.
No longer will you forget or misspell a SAN!
You are also able to add extra SANs or change any of the existing values (e.g., OU, State, City)
Most of the command-line attributes of "Request-CsCertificate" have been catered for in this release & are just passed-through unaltered.
After the request is generated the script outputs a table comparing the values of the existing and new certificates, highlighting
in colour any values that differ between them. Expected changes (expiry, thumbprint & values changed by the user from the command-line)
will be highlighted as Warnings. Values not meant to have changed will be highlighted as Errors.
If you agree with the content of the new certificate, you'll find the appropriate "Assign-CsCertificate" commandlet on your Clipboard ready to 
paste and enter. Alternatively you can use the Deployment Wizard to activate it.

.NOTES  
    Version				: 1.3
	Date				: 12th May 2018
	Author    			: Greig Sheridan
	Credits & references at the bottom of the script
	
	Revision History 	:
				v1.3: 	12th May 2018
						Added an abort line that kills the script when running in the (unsupported) PowerShell ISE. (Screen-width and coloured output don't work)
	
				v1.2	24th December 2017
						Improved the way the "Subject" is parsed in ParseCertSubject by trimming leading spaces
						Added "E=" for those certs where an e-mail address has been provided. (Not applicable to new cert requests via the "Request-CsCertificate")
						Changed the cert comparison highlighting: no longer shows in 'warning' colour if the user provided a 'new' value but nothing changed in the resulting cert
						Incorporated my version of Pat's "Get-UpdateInfo". Credit: https://ucunleashed.com/3168
				
				v1.1:  19th February 2017 - the "Thank You Mike Shivtorov" bugfix & suggestions release
						Added 'XmppServer' certificate type, overlooked in the original release.
						Changed Output request file example text from ".pfx" to ".req".
						Added another example to better document how to generate Edge certificates.
						Improved offline request process: now finds and re-opens the offline request, then feeds it to the comparison engine for display.
						Added an extra Exception trap to the Request-CsCertificate handling: script now reports cleanly if you request an inappropriate Type.
						Sorted Key Usages before sending them to the Compare engine in an effort to reduce false positives.
						Fixed bug where KeySize was reported in red instead of yellow when the user provided a new value
				
				v1.0:  7th May 2016. Initial public release.
						Improved the error reporting when Request-CsCertificate fails
						
				v0.3: 28th Apr 2016 - Beta 3.
						Changed the way I read SANs for improved Server 2008 capability
				
				v0.2: 27th Apr 2016 - Beta Release 2.
				
				v0.1: 1st Apr 2016 - Limited Beta release.
				
			
.LINK  
    https://greiginsydney.com/Update-SfbCertificate.ps1

.EXAMPLE
	.\Update-SfbCertificate.ps1 
 
	Description
	-----------
    With no input parameters passed to it, the script will prompt for the "Type" (or "usages") of the running certificate you want to update.


.EXAMPLE
	.\Update-SfbCertificate.ps1 -type Default,WebServicesExternal,WebServicesInternal -CA "dc01.contoso.com\MyPKiCa" -template "SfB5YrExpiry"
	-FriendlyName "SfB AU Pool internal cert April 2016" -Organization "Contoso"
 
	Description
	-----------
	This reads the existing Default, WebServicesExternal & WebServicesInternal certificate(s) from Lync/SfB and sends a new request to the nominated CA for a replacement
	certificate of the same type. If there was an "Organization" value in the original cert it is replaced with "Contoso" in this request.
	
.EXAMPLE
	.\Update-SfbCertificate.ps1 -Thumbprint "7C742B833005EBCF59D29D06CCEB140BC1EAAABB" -type Default,WebServicesExternal,WebServicesInternal 
	-CA "dc01.contoso.com\MyPKiCa" -template "SfB5YrExpiry" -FriendlyName "SfB AU Pool internal cert April 2016" -AllPoolMemberServers
 
	Description
	-----------
	This reads the certificate with the nominated thumbprint and submits a request to Lync/Sfb for a new Default, WebServicesExternal & WebServicesInternal 
	certificate that matches the original. If this server is a member of a pool, all of the member server FQDNs will be added as SANs.
	
.EXAMPLE
	.\Update-SfbCertificate.ps1 -type AccessEdgeExternal,DataEdgeExternal,AudioVideoAuthentication,XmppServer -ComputerFqdn $null -PrivateKeyExportable $True
	-FriendlyName "SfB AU Edge Pool public cert Jan 2017" -Output "c:\EdgeCsrJan2017.req"
 
	Description
	-----------
	This generates a new Offline certificate request for an Edge server. It reads the certificate(s) with the nominated usages (Types) 
	and generates a new certificate request that matches the original.
	
	

.PARAMETER AllPoolMemberServers
		Switch. When present, if this server is a member of a Front-End pool, the FQDNs of all OTHER member servers will also be added as SANs.
		
.PARAMETER AllSipDomain		
		Switch. When present, all your SIP domains are automatically added to the certificates Subject Alternative Name field.
		If not present, only the primary SIP domain is added by default.

.PARAMETER CA
		String. Fully qualified domain name (FQDN) that points to the CA. For example: -CA "atl-ca-001.litwareinc.com\myca". To obtain a list 
		of known CAs, type "certutil" at the Windows PowerShell prompt. The Config property returned by Certutil indicates the location of a CA
		
.PARAMETER CaAccount
		String. Account name of the user requesting the new certificate, using the format domain_name\user_name. 
		For example: -CaAccount "litwareinc\kenmyer". If not specified, the Request-CsCertificate cmdlet will use the credentials 
		of the logged-on user when requesting the new certificate

.PARAMETER CaPassword
		String. Password for the user requesting the new certificate (as specified using the CaAccount parameter)
		
.PARAMETER ClientEKU 
		Boolean. Set this parameter to True if the certificate is to be used for client authentication. This type of authentication 
		is required if you want your users to be able to exchange instant messages with people who have accounts with AOL. The EKU 
		portion of the parameter name is short for extended key usage; the extended key usage field lists the valid uses for the certificate
		
.PARAMETER City
		String. City where the certificate will be deployed		
		
.PARAMETER ComputerFqdn
		String. If NOT provided, the Topology will be consulted to build the new cert using this machine's FQDN.
		Specify '-computerFQDN $null' to NOT have this parameter included in the request passed to Lync/SfB

.PARAMETER Confirm
		Switch. Prompts you for confirmation before executing the command. (Defaults to $True))
		
.PARAMETER Country
		String. Country/region where the certificate will be deployed
		
.PARAMETER DomainName
		String. Comma-separated list of fully-qualified domain names that should be added to the certificate’s Subject Alternative Name field.
		For example: -DomainName "atl-cs-001.litwareinc.com, atl-cs-002.litwareinc.com,atl-cs-003.litwareinc.com"
		
.PARAMETER FriendlyName
		String. User-assigned name that makes it easier to identify the certificate
		
.PARAMETER GlobalCatalog
		String. FQDN of a global catalog server in your domain. This parameter is not required if you are running the Request-CsCertificate cmdlet 
		on a computer with an account in your domain.
	
.PARAMETER GlobalSettingsDomainController 
		String. FQDN of a domain controller where global settings are stored. If global settings are stored in the System container in Active 
		Directory Domain Services then this parameter must point to the root domain controller. If global settings are stored in the Configuration 
		container then any domain controller can be used and this parameter can be omitted.
		
.PARAMETER KeySize
		Int. Indicates the size (in bits) of the private key used by the certificate. Larger key sizes are more secure, but require more processing 
		overhead in order to be decrypted. Valid key sizes are 1024; 2048; and 4096. For example: -KeySize 2048

.PARAMETER Organization
		String. Name of the organization requesting the new certificate. For example: -Organization "Litwareinc"
		
.PARAMETER OU
		String. Name of the department requesting the new certificate. For example: -OU "IT"

.PARAMETER Output
		String. Path to the certificate file. If you want to create an offline certificate request use the 
		Output parameter and specify a file path for the certificate request; for example: -Output C:\Certificates\NewCertificate.req
		
.PARAMETER PrivateKeyExportable 
		Boolean. Set this parameter to True if you want to make the certificate’s private key exportable.
		When a private key is exportable, the certificate can be copied and used on multiple computers

.PARAMETER Report
		String. Enables you to specify a file path for the log file created when the cmdlet runs. For example: -Report "C:\Logs\Certificates.html"

.PARAMETER State
		String. State where the certificate will be deployed. For example: -State WA
		
.PARAMETER Template
		String. Indicates the certificate template to be used when generating the new certificate; for example: -Template "WebServer". 
		The requested template must be installed on the CA. Note that the value entered must be the template name, not the template display name

.PARAMETER Thumbprint
		String. If provided, the new cert (of the type you nominate) will be based upon this certificate
		
.PARAMETER Type
		String. The type of certificate. Choose from AccessEdgeExternal, AudioVideoAuthentication, DataEdgeExternal, Default, 
		External, Internal, iPhoneAPNService, iPadAPNService, MPNService, PICWebService, ProvisionService, WebServicesExternal,
		WebServicesInternal, WsFedTokenTransfer, OAuthTokenIssuer
	
.PARAMETER SkipUpdateCheck
		Boolean. Skips the automatic check for an Update. Courtesy of Pat: http://www.ucunleashed.com/3168
		
#>

[CmdletBinding(SupportsShouldProcess = $False)]
Param(
	
	[Parameter(ParameterSetName = "Offline",Mandatory)]
	[Parameter(ParameterSetName = "Online",Mandatory)]
    [ValidateSet("AccessEdgeExternal", "AudioVideoAuthentication", "DataEdgeExternal", "Default", "External", "Internal",`
	"iPhoneAPNService", "iPadAPNService", "MPNService", "PICWebService", "ProvisionService", "WebServicesExternal",`
		"WebServicesInternal", "WsFedTokenTransfer", "OAuthTokenIssuer", "XmppServer")]
	[String[]] $Type, #Microsoft.Rtc.Management.Deployment.CertType[]

	[Parameter(ParameterSetName = "Online",Mandatory)]
	[string] $CA,
	
	[Parameter(ParameterSetName = "Online")]
	[string]$CaAccount,
	
	[Parameter(ParameterSetName = "Online")]
	[string]$CaPassword,	
		
	[Parameter(ParameterSetName = "Offline",Mandatory)]
	[string]$Output,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[switch]$AllSipDomain,

	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[switch]$AllPoolMemberServers,	
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
    [string]$City,

	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
    [bool]$ClientEKU = $False,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$ComputerFqdn= ([System.Net.Dns]::GetHostByName(($env:computername))).Hostname,	
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[switch]$Confirm = $True,

	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$Country,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$DomainName,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$FriendlyName = "",

	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$GlobalCatalog,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$GlobalSettingsDomainController,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[ValidateSet(1024, 2048, 4096)]
	[int]$KeySize = 0,
		
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[alias("Organisation")][string]$Organization,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$OU,

	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
    [bool]$PrivateKeyExportable = $True,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$Report,	

	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$State,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$Template,
	
	[Parameter(ParameterSetName = "Offline")]
	[Parameter(ParameterSetName = "Online")]
	[string]$Thumbprint,
	
	[switch] $SkipUpdateCheck
)


$ScriptVersion = "1.3"  #Used by Get-UpdateInfo
$Error.Clear()          #Clear PowerShell's error variable
$Global:Debug = $psboundparameters.debug.ispresent


#--------------------------------
# START FUNCTIONS ---------------
#--------------------------------

function CompareCertParameters
{
	param (
	[Parameter(Mandatory=$True)][string]$parameterName,
	[Parameter(Mandatory=$False)][string]$OldCertValue = "",
	[Parameter(Mandatory=$False)][string]$NewCertValue = "",
	[Parameter(Mandatory=$False)][string]$UserOverrideValue = "",
	[Parameter(Mandatory=$False)][bool]$WarnOnly
	)

	#If no highlighting, default to the user's normal colours:
	$OldBackground = $UserBackgroundColour 
	$OldForeground = $USerForegroundColour 
	$NewBackground = $UserBackgroundColour 
	$NewForeground = $USerForegroundColour 
	
	#Test if we need to change colour:
	if ($UserOverrideValue -ne "")
	{
		if ($UserOverrideValue -eq $NewCertValue)
		{
			if ($NewCertValue -eq $OldCertValue)
			{
				#The user has specified a value, however it's the same as existing, so don't highlight it.
			}
			else
			{
				#Warning - alert the user their change was successful
				$NewForeground = $UserColours.WarningForegroundColor  
				$NewBackground = $UserColours.WarningBackgroundColor  
			}
		}
		else
		{
			#Error
			if ($NewCertValue -ne "")
			{
				#Error
				$NewForeground = $UserColours.ErrorForegroundColor    
				$NewBackground = $UserColours.ErrorBackgroundColor 
			}
			else
			{
				#If the value is no longer present in the new cert, apply the highlight to the old cert:
				#Error
				$OldForeground = $UserColours.ErrorForegroundColor    
				$OldBackground = $UserColours.ErrorBackgroundColor 
			}   
		}
	}
	else
	{
		if ($OldCertValue -ne $NewCertValue)
		{
			if ($NewCertValue -ne "")
			{
				if ($WarnOnly) #For some values a difference is OK (expected). Warn rather than Err.
				{
					$NewForeground = $UserColours.WarningForegroundColor  
					$NewBackground = $UserColours.WarningBackgroundColor  
				}
				else
				{
					#Error
					$NewForeground = $UserColours.ErrorForegroundColor    
					$NewBackground = $UserColours.ErrorBackgroundColor 
				}
			}
			else
			{
				#If the value is no longer present in the new cert, apply the highlight to the old cert:
				#Error
				$OldForeground = $UserColours.ErrorForegroundColor    
				$OldBackground = $UserColours.ErrorBackgroundColor 
			}
		}
	}
	
	$OldCertValue =  truncate $OldCertValue ($global:ColumnWidth)
	$NewCertValue =  truncate $NewCertValue ($global:ColumnWidth)
	write-host ($parameterName).PadRight($global:HeaderWidth," ") -noNewLine 
	write-host " " -NoNewLine
	write-host ($OldCertValue).PadRight($global:ColumnWidth,' ') -noNewLine -foregroundcolor $OldForeground -backgroundcolor $OldBackground
	write-host " " -NoNewLine
	write-host ($NewCertValue).PadRight($global:ColumnWidth,' ') -foregroundcolor $NewForeground -backgroundcolor $NewBackground
}

function truncate
{
	param ([string]$value, [int]$MaxLength)
	
	if ($MaxLength -gt 0) { $MaxLength-- }
	if ($value.Length -gt $MaxLength)
	{
		$value = $value[0..($MaxLength - 3)] -join ""
		$value += "..."
	}
	return $value
}

function ParseCertSubject
{
	param (
	[Parameter(Mandatory=$True)][string]$Subject
	) 
	
	$ItemHash = @{ "Common Name" = ""; "Country" = ""; "State" = ""; "City" = ""; "Organization" = ""; "OU" = ""; "E-mail" = ""}
	$CertSubject = ($Subject).Split(",")
	foreach ($CertSubjectValue in $CertSubject)
	{
		$CertSubjectValue = $CertSubjectValue.Trim()
		if ($CertSubjectValue.StartsWith("CN=")) { $ItemHash.("Common Name")  =	$CertSubjectValue.Substring(3) }
		if ($CertSubjectValue.StartsWith("C="))  { $ItemHash.("Country") 	  = $CertSubjectValue.Substring(2) }
		if ($CertSubjectValue.StartsWith("S="))  { $ItemHash.("State")        = $CertSubjectValue.Substring(2) }
		if ($CertSubjectValue.StartsWith("L="))  { $ItemHash.("City")         =	$CertSubjectValue.Substring(2) }
		if ($CertSubjectValue.StartsWith("O="))  { $ItemHash.("Organization") = $CertSubjectValue.Substring(2) }
		if ($CertSubjectValue.StartsWith("OU=")) { $ItemHash.("OU")           =	$CertSubjectValue.Substring(3) }
		if ($CertSubjectValue.StartsWith("E="))  { $ItemHash.("E-mail")       = $CertSubjectValue.Substring(2) }
	}
	return $ItemHash
}

function DecodeSANs
{
	param (
	[Parameter(Mandatory=$True)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Cert
	) 
	
	$SANs = @()
	#$SANs += $singleCert.dnsnamelist.unicode
	
	#Server 2008(?) hides the SANs away here (if there are any at all):
	try
	{
		$S2008SANs = ($cert.Extensions | Where-Object {$_.Oid.FriendlyName -match "subject alternative name"}).Format(1)
	}
	catch
	{
		$S2008SANs = ""
	}
	$S2008SANs = [regex]::replace($S2008SANs, "`r`n", "") #Trim CRLFs
	$S2008SANsArray = $S2008SANs -split "DNS Name="
	$SANs += $S2008SANsArray
	$SANs  = $SANs | ? {$_} | select -uniq 	#De-dupe
	return $SANs
}


function Get-UpdateInfo
{
  <#
      .SYNOPSIS
      Queries an online XML source for version information to determine if a new version of the script is available.
	  *** This version customised by Greig Sheridan. @greiginsydney https://greiginsydney.com ***

      .DESCRIPTION
      Queries an online XML source for version information to determine if a new version of the script is available.

      .NOTES
      Version               : 1.2 - See changelog at https://ucunleashed.com/3168 for fixes & changes introduced with each version
      Wish list             : Better error trapping
      Rights Required       : N/A
      Sched Task Required   : No
      Lync/Skype4B Version  : N/A
      Author/Copyright      : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
      Email/Blog/Twitter    : pat@innervation.com  https://ucunleashed.com  @patrichard
      Donations             : https://www.paypal.me/PatRichard
      Dedicated Post        : https://ucunleashed.com/3168
      Disclaimer            : You running this script/function means you will not blame the author(s) if this breaks your stuff. This script/function 
                            is provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including, without limitation, 
                            any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use 
                            or performance of the sample scripts and documentation remains with you. In no event shall author(s) be held liable for 
                            any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss 
                            of business information, or other pecuniary loss) arising out of the use of or inability to use the script or 
                            documentation. Neither this script/function, nor any part of it other than those parts that are explicitly copied from 
                            others, may be republished without author(s) express written permission. Author(s) retain the right to alter this 
                            disclaimer at any time. For the most up to date version of the disclaimer, see https://ucunleashed.com/code-disclaimer.
      Acknowledgements      : Reading XML files 
                            http://stackoverflow.com/questions/18509358/how-to-read-xml-in-powershell
                            http://stackoverflow.com/questions/20433932/determine-xml-node-exists
      Assumptions           : ExecutionPolicy of AllSigned (recommended), RemoteSigned, or Unrestricted (not recommended)
      Limitations           : 
      Known issues          : 

      .EXAMPLE
      Get-UpdateInfo -Title "Update-SfbCertificate.ps1"

      Description
      -----------
      Runs function to check for updates to script called <Varies>.

      .INPUTS
      None. You cannot pipe objects to this script.
  #>
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
	[string] $title
	)
	try
	{
		[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
		if ($HasInternetAccess)
		{
			write-verbose "Performing update check"
			# ------------------ TLS 1.2 fixup from https://github.com/chocolatey/choco/wiki/Installation#installing-with-restricted-tls
			$securityProtocolSettingsOriginal = [System.Net.ServicePointManager]::SecurityProtocol
			try {
			  # Set TLS 1.2 (3072). Use integers because the enumeration values for TLS 1.2 won't exist in .NET 4.0, even though they are 
			  # addressable if .NET 4.5+ is installed (.NET 4.5 is an in-place upgrade).
			  [System.Net.ServicePointManager]::SecurityProtocol = 3072
			} catch {
			  Write-verbose 'Unable to set PowerShell to use TLS 1.2 due to old .NET Framework installed.'
			}
			# ------------------ end TLS 1.2 fixup
			[xml] $xml = (New-Object -TypeName System.Net.WebClient).DownloadString('https://greiginsydney.com/wp-content/version.xml')
			[System.Net.ServicePointManager]::SecurityProtocol = $securityProtocolSettingsOriginal #Reinstate original SecurityProtocol settings
			$article  = select-XML -xml $xml -xpath "//article[@title='$($title)']"
			[string] $Ga = $article.node.version.trim()
			if ($article.node.changeLog)
			{
				[string] $changelog = "This version includes: " + $article.node.changeLog.trim() + "`n`n"
			}
			if ($Ga -gt $ScriptVersion)
			{
				$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
				$updatePrompt = $wshell.Popup("Version $($ga) is available.`n`n$($changelog)Would you like to download it?",0,"New version available",68)
				if ($updatePrompt -eq 6)
				{
					Start-Process -FilePath $article.node.downloadUrl
					Write-Warning "Script is exiting. Please run the new version of the script after you've downloaded it."
					exit
				}
				else
				{
					write-verbose "Upgrade to version $($ga) was declined"
				}
			}
			elseif ($Ga -eq $ScriptVersion)
			{
				write-verbose "Script version $($Scriptversion) is the latest released version"
			}
			else
			{
				write-verbose "Script version $($Scriptversion) is newer than the latest released version $($ga)"
			}
		}
		else
		{
		}
	
	} # end function Get-UpdateInfo
	catch
	{
		write-verbose "Caught error in Get-UpdateInfo"
		if ($Global:Debug)
		{				
			$Global:error | fl * -f #This dumps to screen as white for the time being. I haven't been able to get it to dump in red
		}
	}
}


#--------------------------------
# END  FUNCTIONS ---------------
#--------------------------------


#--------------------------------
# THE FUN STARTS HERE -----------
#--------------------------------

#Requires -Modules Lync #Loads the Lync or SfB module if it's not already loaded. (NB: SfB servers can get away with 'requiring' the "Lync" module)

## #Requires -RunAsAdministrator #Can't use this here - it wasn't added until v4.
If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
 {    
  Echo "This script needs to be run As Admin"
  Break
 }
# Why force Admin? Strangely, the keysize isn't reported in the compare if you run as an ordinary user.
 
If ($PsVersionTable.PsVersion.Major -lt 3)
{
	write-warning "This version of PowerShell is not able to read some certificate values."
	write-warning "The script comparison cannot be guaranteed to be complete."
}

if ($skipupdatecheck)
{
	write-verbose "Skipping update check"
}
else
{
	write-progress -id 1 -Activity "Performing update check" -PercentComplete (50)
	Get-UpdateInfo -title "Update-SfbCertificate.ps1"
	write-progress -id 1 -Activity "Back from performing update check" -Completed
}


$UserColours = (Get-Host).PrivateData
$USerForegroundColour = (get-host).ui.rawui.ForegroundColor
$UserBackgroundColour = (get-host).ui.rawui.BackgroundColor
$UserScreenWidth = [int](get-host).UI.rawui.Windowsize.Width
if ($UserScreenWidth -eq 0)
{
	echo "Powershell ISE detected. It doesn't report the screen width & mishandles/errs when writing to the screen in colour"
	echo "Please re-run from a normal PS window, elevated"
	break
}
$global:HeaderWidth = ([Math]::Truncate($UserScreenWidth * 0.2) -2) #Subtracting 2 allows for the space that P$ automatically
$global:ColumnWidth = ([Math]::Truncate($UserScreenWidth * 0.4) -2) # puts between the columns when using the PadLeft/Right commands

$DomainName = [regex]::replace($DomainName, " ", "") #Strip any spaces (if the user separated the SANs)

$Thumbprints = @() #The list of 1 or more certificate thumbprints to be queried
$ServerCerts = @() #The list of 1 or more certificates to be queried (from the above list))

# Did the user provide a Thumbprint to copy, or are we determining the cert(s) to copy from the usage ($Type)?
if ($Thumbprint)
{
	$Thumbprints = [regex]::replace($Thumbprint, "[^A-Fa-f0-9]", "") #Trim spaces and the leading "?" that comes if you paste
}
else
{
	try
	{
		#Read all of the certificate(s) for the given cert types:
		$Thumbprints = (get-cscertificate -Type $Type -verbose:$false).Thumbprint
		$Thumbprints  = $Thumbprints  | select -uniq 	#De-dupe
		write-verbose "Type '$($type)' returned $($Thumbprints.count) unique certificate(s)"
	}
	catch
	{
		$_ | fl * -f
		$Thumbprints = $null
	}
}
#Read all the unique thumbprints:
foreach ($Thumb in $Thumbprints)
{
	$ServerCerts += (get-childitem cert:\localmachine\my | ? {$_.thumbprint -match $Thumb} )
}

if ($ServerCerts -ne $null)
{
	$CertRequestParams = @{ Type = $Type } 	# Declare the param's to be passed to the Request. Initialise with the Type
	$NewSANs = @()
	$NewSANs = $DomainName.Split(",")	# The ones the user nominated
										# I use this array later for the comparison test to detect what the user added
	if ($AllPoolMemberServers)
	{
		$MyPool = (Get-CsComputer -Local -verbose:$false).Pool 
		$MyPoolServers = (get-cspool $MyPool -verbose:$false).Computers  #or ".FQDN" here?
		foreach ($MemberServer in $MyPoolServers)
		{
			write-verbose "Adding Pool Servers   : $($MemberServer)"
			$NewSANs += $MemberServer
		}
	}
	# -------------------
	# By my reading, the "AllSipDomain" switch is broken, so if the user's specified it, let's read and add all the SIP domains to be sure:
	#--------------------
	if ($AllSipDomain)
	{
		$TopoSIPDomains = (Get-CsSipDomain -verbose:$false).Identity
		foreach ($TopoSIPDomain in $TopoSIPDomains)
		{
			$NewSANs += ("sip.$($TopoSIPDomain)")
			write-verbose "Adding all SIP domains: sip.$($TopoSIPDomain)"
		}
	}
	# -------------------
	# This next bit consolidates what might be multiple certs back into one for the request.
	# If the nominated "types" are spread across more than one cert, the Thumb will be updated for the comparison check
	# All of the SANs will be consolidated here too (just in case you're going from separate certs per role/type on your FE back to one)
	#--------------------
	$SfBThumb = ""	#This is the "before" thumbprint value we'll pass to the comparison test
	$DonorSANs = @()
	foreach ($singleCert in $ServerCerts)
	{
		if ($SfBThumb -eq "") { $SfBThumb = $singleCert.Thumbprint } #This initialises the value with $ServerCerts[0]
		if ($SfBThumb -eq $singleCert.Thumbprint)
		{
			$DonorSANs += DecodeSANs $singleCert
		}
		else
		{
			$SfBThumb = "<Multiple>"
			$DonorSANs += DecodeSANs $singleCert
		}
	}
	if ($SfBThumb -match "<Multiple>")
	{
		write-warning "Multiple existing certificates are being consolidated into one new certificate"
	}
	$TotalSANs = $DonorSANs + $NewSANs
	$TotalSANs  = $TotalSANs  | select -uniq 	#De-dupe
	$NewCertSanCsv = ""
	foreach ($OneSAN in $TotalSANs)
	{ 
		if ($OneSAN -ne "") 
		{
			$NewCertSanCsv += $OneSAN + "," #Convert the SANs to a comma-separated list
			write-verbose "Adding SAN            : $($OneSAN)"
		}	
	}
	if ($NewCertSanCsv -ne "") 
	{
		$NewCertSanCsv = $NewCertSanCsv.Substring(0,$NewCertSanCsv.Length-1) #Strip the trailing comma
		$CertRequestParams.DomainName = $NewCertSanCsv 
	}
	#--------------------
	# OK, that's the SANs sorted.
	# If there are multiple donor certs, we only read the remaining ("City" etc values) from the first one.
	#--------------------
	$SfbCertSubjectItem = @{}
	$SfbCertSubjectItem = ParseCertSubject $ServerCerts[0].Subject # Convert the string's values into a hash table
	if ($SfBCertSubjectItem.Get_Item("Country") -ne "")      { $CertRequestParams.Country      = $SfBCertSubjectItem.Get_Item("Country")}
	if ($SfBCertSubjectItem.Get_Item("State") -ne "")        { $CertRequestParams.State        = $SfBCertSubjectItem.Get_Item("State")}
	if ($SfBCertSubjectItem.Get_Item("City") -ne "")         { $CertRequestParams.City         = $SfBCertSubjectItem.Get_Item("City")}
	if ($SfBCertSubjectItem.Get_Item("Organization") -ne "") { $CertRequestParams.Organization = $SfBCertSubjectItem.Get_Item("Organization")}
	if ($SfBCertSubjectItem.Get_Item("OU") -ne "")           { $CertRequestParams.OU           = $SfBCertSubjectItem.Get_Item("OU")}
	#SfB won't accept a request with an e-mail value specified. If there is one in the original, it won't be copied across to the request:
	#if ($SfBCertSubjectItem.Get_Item("E-mail") -ne "")       { $CertRequestParams.E            = $SfBCertSubjectItem.Get_Item("E-mail")}
	write-host # A blank link here adds some spacing between the verbose text and the warnings that might follow.	
	if ($ServerCerts[0].PrivateKey.KeySize -ne $null)
	{
		$CertRequestParams.KeySize = $ServerCerts[0].PrivateKey.KeySize
		if ($ServerCerts[0].PrivateKey.KeySize -eq 1024)
		{
			# This gives the user an opportunity to abort and change to a more secure cert:
			write-warning "The existing certificate has a Keysize of only 1024 bits"
		}
	}
	if ($FriendlyName -eq "") { write-warning "You've not specified a Friendly Name. A default one will be applied if you proceed" }
	$SnChanged = 0 #This influences how we display the consolidated "CN=" line in the comparison
	#--------------------
	# Now stamp all of the values from the command-line into the parameters list, over-writing any already there:
	#--------------------
	if ($AllSipDomain -ne $false) { $CertRequestParams.AllSipDomain = $True}
	if ($Ca -ne "") 			{ $CertRequestParams.Ca = $Ca}
	if ($CaAccount -ne "") 		{ $CertRequestParams.CaAccount = $CaAccount}
	if ($CaPassword -ne "") 	{ $CertRequestParams.CaPassword = $CaPassword}
	if ($City -ne "")    		{ $CertRequestParams.City = $City; $SnChanged = 1}
	if ($ComputerFQDN -ne "") 	{ $CertRequestParams.ComputerFQDN = $ComputerFQDN}
	if ($Country -ne "")    	{ $CertRequestParams.Country = $Country; $SnChanged = 1}
	if ($FriendlyName -ne "") 	{ $CertRequestParams.FriendlyName = $FriendlyName}
	if ($GlobalCatalog -ne "") 	{ $CertRequestParams.GlobalCatalog = $GlobalCatalog}
	if ($GlobalSettingsDomainController -ne "") 	{ $CertRequestParams.GlobalSettingsDomainController = $GlobalSettingsDomainController}
	if ($KeySize -ne 0) 		{ $CertRequestParams.KeySize = $KeySize} 
	if ($Organization -ne "") 	{ $CertRequestParams.Organization = $Organization; $SnChanged = 1}
	if ($OU -ne "") 			{ $CertRequestParams.OU = $OU; $SnChanged = 1}
	if ($Output -ne "") 		{ $CertRequestParams.Output = $Output}
	if ($Report -ne "") 		{ $CertRequestParams.Report = $Report}
	if ($State -ne "") 			{ $CertRequestParams.State = $State; $SnChanged = 1}
	if ($Template -ne "") 		{ $CertRequestParams.Template = $Template}
	$CertRequestParams.PrivateKeyExportable = $PrivateKeyExportable
	$CertRequestParams.ClientEKU = $ClientEKU
	if ($Confirm -ne $False)
	{
		$CertRequestParams.Confirm = $True
		# Dumps the arguments to the screen ahead of the Request commandlet prompting you to continue
		$CertRequestParams
	}
	if ($Output -ne "")
	{
		# It's an offline request: snapshot the current REQUEST store so we can find out request after.
		$PendingRequests = @()
		$PendingRequests = Get-ChildItem cert:\localmachine\REQUEST
	}
	$CertCreated = $null
	try
	{
		$CertCreated = Request-csCertificate -new @CertRequestParams -verbose:$false -warningaction silentlycontinue
	}
	catch [System.ArgumentException]
	{
		write-output ""
		if ($_.FullyQualifiedErrorId -match "CentralMgmtStoreInaccessible")
		{
			write-warning "The ""Request-CsCertificate"" commandlet threw an error:"
			write-warning "Unable to read the Topology"
			write-warning "Is this an Edge (non domain-joined) machine?"
			write-warning "Re-run, adding ""-ComputerFqdn `$null"""
			exit
		}
		elseif ($_.Exception -match "Unable to contact certification authority")
		{
			write-warning "The ""Request-CsCertificate"" commandlet threw an error:"
			write-warning "Unable to contact the CA ""$($ca)"""
			exit
		}
		elseif ($_.Exception -match "and that you logged on with the correct credentials")
		{
			write-warning "The ""Request-CsCertificate"" commandlet threw an error:"
			write-warning "The CA reported a problem with your credentials"
			exit
		}
		else
		{
			write-warning "The ""Request-CsCertificate"" commandlet threw an error."
			write-warning "There was a problem with the arguments you passed:"
			$_ | fl * -f
		}
	}
	catch [Microsoft.Rtc.Management.Common.Certificates.CertificateException]
	{
		write-warning "The ""Request-CsCertificate"" commandlet threw an error."
		if ($_.Exception -match "The request was for a certificate template that is not supported")
		{
			write-warning "The CA was unable to issue the certificate with the template ""$($Template)"""
			exit
		}
		elseif ($_.Exception -match "Denied by Policy Module 0x80094800")
		{
			write-warning "The CA reported a problem:"
			write-warning "Denied by Policy Module 0x80094800"
			exit
		}
		elseif ($_.Exception -match "Denied by Policy Module")
		{
			write-warning "The CA reported a problem:"
			write-warning "Denied by Policy Module"
			exit
		}
		else
		{
			write-warning "The CA reported a problem:"
			$_ | fl * -f
			exit
		}
	}
	catch  [System.InvalidOperationException]
	{
		write-warning "The ""Request-CsCertificate"" commandlet threw an error."
		if ($_.Exception -match "The computer does not need a certificate")
		{
			$ReformatError = ($_.ErrorDetails).ToString().Split(".")
			write-warning (($ReformatError[0]) -replace "Command execution failed: ", "")
			write-warning (($ReformatError[1]).Trim())
			exit
		}
	}
	catch
	{
		$_ | fl * -f
	}
	
	if ($CertCreated -ne $null)
	{
		$IsOfflineCsr = 0
		if ($CertCreated.RequestStatus -eq "Offline")
		{
			#The output of the commandlet doesn't give me the Thumbprint of the new request, so I need to determine it myself.
			$CurrentRequests = Get-ChildItem cert:\localmachine\REQUEST
			if ($CurrentRequests.Count -eq 0)
			{
				write-warning ""
				write-warning "New cert request not found in the localmachine\REQUEST store. Aborting"
				write-warning ""
				exit
			}
			if ($PendingRequests.Count -eq 0)
			{
				#There were no requests previously pending - nothing to compare
				$NewCert = $CurrentRequests 
			}
			else
			{
				$NewCert = Compare-Object -ReferenceObject $PendingRequests -DifferenceObject $CurrentRequests -PassThru
			}
			$IsOfflineCsr = 1 #Changes the expected differences in Issuer and IssuerName to show yellow, not red.
		}
		else
		{
			$Newcert = (Get-ChildItem cert:\localmachine\My | ? {$_.Thumbprint -match $CertCreated.Thumbprint} ) 
		}
		#Read all the properties of BOTH certs & then de-dupe. This will trap any that are present on one but not the other:
		$properties  = ($ServerCerts[0] | Get-Member -MemberType Property | Select-Object  -ExpandProperty Name)
		$properties += ($NewCert | Get-Member -MemberType Property | Select-Object  -ExpandProperty Name)
		$properties  = $properties  | select -uniq	
		
		write-host ""
		write-host  "Attribute".PadRight($HeaderWidth, " ")"Certificate 1".PadRight($ColumnWidth, " ")"Certificate 2".PadRight($ColumnWidth, " ")
		write-host  ("---------").PadRight($HeaderWidth, " ")("-------------").PadRight($ColumnWidth, " ")("-------------").PadRight($ColumnWidth, " ")
		foreach ($property in $properties)
		{
			switch ($property)
			{
				{($_ -eq "Thumbprint") -or ($_ -eq "PrivateKey")}
				{
					#Skip Thumbprint - we'll manually write it as the last parameter. Private Key we do manually under "HasPrivateKey"
					Continue
				} 
				{($_ -eq "notbefore") -or ($_ -eq "notafter") -or ($_ -eq "SerialNumber")}
				{
					CompareCertParameters $property $ServerCerts[0]."$($property)" $NewCert."$($property)" "" 1 #Force changes to show as yellow - they're expected
				}
				"issuer"
				{
					CompareCertParameters $property $ServerCerts[0]."$($property)" $NewCert."$($property)" "" $IsOfflineCsr
				}
				"IssuerName"
				{
					CompareCertParameters $property ($ServerCerts[0]."$($property)").Name ($NewCert."$($property)").Name "" $IsOfflineCsr
				}
				"FriendlyName"
				{
					$Cert1FriendlyName = "<None>"
					$Cert2FriendlyName = "<None>"
					if ($ServerCerts[0].FriendlyName -ne "") { $Cert1FriendlyName = $ServerCerts[0].FriendlyName }
					if ($NewCert.FriendlyName -ne "") { $Cert2FriendlyName = $NewCert.FriendlyName }
					CompareCertParameters "Friendly Name" $Cert1FriendlyName $Cert2FriendlyName $FriendlyName
				}
				"HasPrivateKey"
				{
					CompareCertParameters $property $ServerCerts[0]."$($property)" $NewCert."$($property)" ""
					CompareCertParameters "Key Size" $ServerCerts[0].PrivateKey.KeySize $NewCert.PrivateKey.KeySize $NewCert.PrivateKey.KeySize
				}
				"subject"
				{
					CompareCertParameters "Subject" $ServerCerts[0]."$($property)" $NewCert."$($property)" "" $SnChanged
					$NewCertSubjectItem = @{}
					$NewCertSubjectItem = ParseCertSubject $NewCert.Subject # Convert the string's values into a hash table
					CompareCertParameters "Common Name" 	$SfBCertSubjectItem.Get_Item("Common Name") 		$NewCertSubjectItem.Get_Item("Common Name") ""
					CompareCertParameters "Country" 		$SfBCertSubjectItem.Get_Item("Country") 			$NewCertSubjectItem.Get_Item("Country") 	$Country
					CompareCertParameters "State" 			$SfBCertSubjectItem.Get_Item("State") 				$NewCertSubjectItem.Get_Item("State") 		$State
					CompareCertParameters "City" 			$SfBCertSubjectItem.Get_Item("City") 				$NewCertSubjectItem.Get_Item("City") 		$City
					CompareCertParameters "Organization" 	$SfBCertSubjectItem.Get_Item("Organization") 		$NewCertSubjectItem.Get_Item("Organization") $Organization
					CompareCertParameters "OU" 				$SfBCertSubjectItem.Get_Item("OU") 					$NewCertSubjectItem.Get_Item("OU") 			$OU
					CompareCertParameters "E-mail" 			$SfBCertSubjectItem.Get_Item("E-mail")				$NewCertSubjectItem.Get_Item("E-mail") 		"" # The cmdlet doesn't support providing an e-mail address
				}
				"SignatureAlgorithm"
				{
					CompareCertParameters "Sig Algorithm" ($ServerCerts[0]."$($property)").FriendlyName ($NewCert."$($property)").FriendlyName ""
				}
				"Extensions"
				{
					# Retrieve the usages -> to String -> Strip spaces -> to Array -> Sort -> Back to CSV string !
					$ServerCertsUsagesSorted = ""
					$NewCertUsagesSorted = ""
					if ($ServerCerts[0].extensions.KeyUsages -ne $null)
					{
						$ServerCertsUsagesSorted = ((($ServerCerts[0].extensions.KeyUsages).ToString() -replace " ","").Split(",") | sort) -join ", "
					}
					if ($NewCert.extensions.KeyUsages -ne $null)
					{
						$NewCertUsagesSorted     = ((($NewCert.extensions.KeyUsages       ).ToString() -replace " ","").Split(",") | sort) -join ", "
					}
					
					CompareCertParameters "Key Usages" $ServerCertsUsagesSorted $NewCertUsagesSorted "" $IsOfflineCsr
				}
				#This outputs too much / all information. Might be worthwhile enabling for a -verbose or -allParameters switch in v2?
				default 
				{
					#CompareCertParameters $property $ServerCerts[0]."$($property)" $NewCert."$($property)" "" #No override for these values
				}
			}
		}
		#Create a master SAN list (like we did with Properties above) & de-dupe:
		$AllSANs = $DonorSANs
		$NewcertSANs = DecodeSANs $Newcert
		$AllSANs += $NewcertSANs
		$AllSANs = $AllSANs  | select -uniq	
		foreach ($SAN in $AllSANs)
		{
			# Setting "UserAdded" influences the colour display in the comparison
			if ($NewSANs -contains $SAN)
			{
				$UserAdded = $SAN
			}
			else
			{
				$UserAdded = ""
			}
			#Now pass the 3 possible combinations of SAN:
			#	either it was there and it still is,
			#	it was there and it isn't now, 
			#	its been added new 
			if (($DonorSANs -contains $SAN) -and ($NewcertSANs -contains $SAN))
			{
				CompareCertParameters "SAN" $SAN $SAN $UserAdded
			}
			elseif (($DonorSANs -contains $SAN) -and ($NewcertSANs -notcontains $SAN))
			{
				CompareCertParameters "SAN" $SAN "" $UserAdded
			}
			else
			{
				CompareCertParameters "SAN" "" $SAN $UserAdded
			}
		}
		#Write the Thumbprint last:
		CompareCertParameters "Thumbprint" $SfBThumb $NewCert.Thumbprint "" 1 #Force changes to show as yellow - they're expected
		write-host # A blank line at the end
		
		if ($type -notcontains "OAuthTokenIssuer")
		{
			if ($IsOfflineCsr -eq 1)
			{
$OfflineWarningText=
@"
Don't be alarmed by the values that differ between the existing cert and your offline request. Your issuing CA can and typically does overwrite values like the Key Usages and Sig Algorithm - but best to re-check with my "Compare-PkiCertificates.ps1". Refer https://gallery.technet.microsoft.com/Compare-two-identical-PKI-6dcbfdec
"@			
				write-warning $OfflineWarningText
			}
			else
			{
				try
				{
					"Set-CsCertificate -type $($Type -Join ',') -Thumbprint $($NewCert.Thumbprint) -confirm" | Clip # Paste the "Assign" commandlet to the clipboard}
				} catch {}
			}
		}
		else
		{
			write-warning "Take care assigning a new OAuth cert."
			write-warning "Refer https://technet.microsoft.com/en-us/library/jj660292.aspx for details"
		}
	}
}
else
{
	if ($Thumbprint)
	{
		write-warning "The provided thumbprint is either invalid or for a certificate not installed on this server"
	}
	else
	{
		write-warning "A certificate of type ""$($Type)"" does not exist on this server"
	}
}

#References:
# https://ramblingcookiemonster.wordpress.com/2014/12/01/powershell-splatting-build-parameters-dynamically/
# http://social.technet.microsoft.com/wiki/contents/articles/1447.display-subject-alternative-names-of-a-certificate-with-powershell.aspx

#Code signing certificate kindly provided by Digicert:
# SIG # Begin signature block
# MIIceAYJKoZIhvcNAQcCoIIcaTCCHGUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUQSAwNre7DLdG59GWUV8Q+geF
# 1QCgghenMIIFMDCCBBigAwIBAgIQA1GDBusaADXxu0naTkLwYTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDQxNzAwMDAwMFoXDTIxMDcw
# MTEyMDAwMFowbTELMAkGA1UEBhMCQVUxGDAWBgNVBAgTD05ldyBTb3V0aCBXYWxl
# czESMBAGA1UEBxMJUGV0ZXJzaGFtMRcwFQYDVQQKEw5HcmVpZyBTaGVyaWRhbjEX
# MBUGA1UEAxMOR3JlaWcgU2hlcmlkYW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQC0PMhHbI+fkQcYFNzZHgVAuyE3BErOYAVBsCjZgWFMhqvhEq08El/W
# PNdtlcOaTPMdyEibyJY8ZZTOepPVjtHGFPI08z5F6BkAmyJ7eFpR9EyCd6JRJZ9R
# ibq3e2mfqnv2wB0rOmRjnIX6XW6dMdfs/iFaSK4pJAqejme5Lcboea4ZJDCoWOK7
# bUWkoqlY+CazC/Cb48ZguPzacF5qHoDjmpeVS4/mRB4frPj56OvKns4Nf7gOZpQS
# 956BgagHr92iy3GkExAdr9ys5cDsTA49GwSabwpwDcgobJ+cYeBc1tGElWHVOx0F
# 24wBBfcDG8KL78bpqOzXhlsyDkOXKM21AgMBAAGjggHFMIIBwTAfBgNVHSMEGDAW
# gBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUzBwyYxT+LFH+GuVtHo2S
# mSHS/N0wDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1Ud
# HwRwMG4wNaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3Vy
# ZWQtY3MtZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hh
# Mi1hc3N1cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgG
# CCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEE
# ATCBhAYIKwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMB
# Af8EAjAAMA0GCSqGSIb3DQEBCwUAA4IBAQCtV/Nu/2vgu+rHGFI6gssYWfYLEwXO
# eJqOYcYYjb7dk5sRTninaUpKt4WPuFo9OroNOrw6bhvPKdzYArXLCGbnvi40LaJI
# AOr9+V/+rmVrHXcYxQiWLwKI5NKnzxB2sJzM0vpSzlj1+fa5kCnpKY6qeuv7QUCZ
# 1+tHunxKW2oF+mBD1MV2S4+Qgl4pT9q2ygh9DO5TPxC91lbuT5p1/flI/3dHBJd+
# KZ9vYGdsJO5vS4MscsCYTrRXvgvj0wl+Nwumowu4O0ROqLRdxCZ+1X6a5zNdrk4w
# Dbdznv3E3s3My8Axuaea4WHulgAvPosFrB44e/VHDraIcNCx/GBKNYs8MIIFMDCC
# BBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0Ew
# HhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5n
# IENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfT
# CzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdgl
# rA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRn
# iolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7
# MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPr
# CGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z
# 3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8E
# BAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0g
# BEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nED
# wGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqG
# SIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9
# D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQG
# ivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEeh
# emhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJ
# RZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5
# gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIGajCCBVKgAwIBAgIQAwGa
# Ajr/WLFr1tXq5hfwZjANBgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAw
# WhcNMjQxMDIyMDAwMDAwWjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNl
# cnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBT
# qZ8fZFnmfGt/a4ydVfiS457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWR
# n8YUOawk6qhLLJGJzF4o9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRV
# fRiGBYxVh3lIRvfKDo2n3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3v
# J+P3mvBMMWSN4+v6GYeofs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA
# 8bLOcEaD6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGj
# ggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDCCAb8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIB
# kjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQAS
# KxOYspkH7R7for5XDStnAs0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9
# MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcw
# AoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Q0EtMS5jcnQwDQYJKoZIhvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI
# //+x1GosMe06FxlxF82pG7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7ea
# sGAm6mlXIV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8Oxw
# YtNiS7Dgc6aSwNOOMdgv420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQN
# JsQOfxu19aDxxncGKBXp2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNt
# omHpigtt7BIYvfdVVEADkitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbN
# MIIFtaADAgECAhAG/fkDlgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBD
# QTAeFw0wNjExMTAwMDAwMDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/J
# M/xNRZFcgZ/tLJz4FlnfnrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPs
# i3o2CAOrDDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ
# 8DIhFonGcIj5BZd9o8dD3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNu
# gnM/JksUkK5ZZgrEjb7SzgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJr
# GGWxwXOt1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3ow
# ggN2MA4GA1UdDwEB/wQEAwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUH
# AwIGCCsGAQUFBwMDBggrBgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIB
# xTCCAbQGCmCGSAGG/WwAAQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIw
# ggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQA
# aQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUA
# cAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMA
# UAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEA
# cgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkA
# dAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8A
# cgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIA
# ZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsMB0GA1UdDgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+
# ybcoJKc4HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6
# hnKtOHisdV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5P
# sQXSDj0aqRRbpoYxYqioM+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke
# /MV5vEwSV/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qqu
# AHzunEIOz5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQ
# nHcUwZ1PL1qVCCkQJjGCBDswggQ3AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EANRgwbrGgA18btJ2k5C8GEwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFKrRzXlyyFyXp3jf7gQq
# k3VgNSpNMA0GCSqGSIb3DQEBAQUABIIBAJhUxSGpf+gogIU9RmRbi1kxYNebAhNS
# U2xe8VHsaZxckicPyiXn4OvdS14Wa0C0Z86/qmlWxBKx9K0BW2ugFnuy3LLxBg1x
# AbkqFvRbV+yzHSxJ53k76FQVlFsGVc93VNAmnJZ1lFdcGJ6APxiJBEYWg4xfYlse
# Fj48O6X4WQ+BORFedCn387y96bZdujfshuTCqN9hWgdUUDHGmYc+fwdYyWP1VGCV
# WitS1A4wnDfc0pQsRbf4MdHHCXFGYjR4la9xeHkMQUBL3ly2q/v/E6W0FcqoIB7+
# jzAUgBND4w7rxliHvI6E/vMDjjMxyx2K0/atrUMw71KJ6+tyTC597kihggIPMIIC
# CwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYD
# VQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTECEAMBmgI6/1ixa9bV6uYX8GYw
# CQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcN
# AQkFMQ8XDTIwMDUwNTExMjgyMlowIwYJKoZIhvcNAQkEMRYEFMT9iJg7UQf8PUDN
# ZXzeBCqzb5mnMA0GCSqGSIb3DQEBAQUABIIBAD/MVTyU6WI+aFzqBFBcN7AY7MPC
# XRAubaJCqsFlAOXv+JfmPQJ4vJ0Fr7bZZBaeDJ6eUrC/48Viq2eycN5cVmd5gR6R
# 2dZZOrEnIFCJpa5MgIR6d6E7MRtYHnrzntBOyH/2IZ9xSLWEhPsFBYhquv4ROJHF
# cprVMT4Sy+yGOx7j4q9cPExbGAKLNCld9/3hEs3UF8EufAdIq4IzyAqP1Mcdyeho
# ai6f9YqVInqKHbcghvt/hYKNngeDWtNVh9+m7wnTICguzxHqgeHWz0smC3lcez6x
# r/tpfnNbWrKTvX9IjV+oG6d45MSqUMiti5gel4mvwJIKUljnFdBjYQhp46k=
# SIG # End signature block
