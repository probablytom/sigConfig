<#
The sample scripts are not supported under any Microsoft standard support 
program or service. The sample scripts are provided AS IS without warranty  
of any kind. Microsoft further disclaims all implied warranties including,  
without limitation, any implied warranties of merchantability or of fitness for 
a particular purpose. The entire risk arising out of the use or performance of  
the sample scripts and documentation remains with you. In no event shall 
Microsoft, its authors, or anyone else involved in the creation, production, or 
delivery of the scripts be liable for any damages whatsoever (including, 
without limitation, damages for loss of business profits, business interruption, 
loss of business information, or other pecuniary loss) arising out of the use 
of or inability to use the sample scripts or documentation, even if Microsoft 
has been advised of the possibility of such damages.
#>

#requires -Version 2

#Import Localized Data
Import-LocalizedData -BindingVariable Messages
#Load .NET Assembly for Windows PowerShell V2
Add-Type -AssemblyName System.Core

$webSvcInstallDirRegKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Exchange\Web Services\2.0" -PSProperty "Install Directory" -ErrorAction:SilentlyContinue
if ($webSvcInstallDirRegKey -ne $null) {
	$moduleFilePath = $webSvcInstallDirRegKey.'Install Directory' + 'Microsoft.Exchange.WebServices.dll'
	Import-Module $moduleFilePath
} else {
	$errorMsg = $Messages.InstallExWebSvcModule
	throw $errorMsg
}

Function New-OSCPSCustomErrorRecord
{
	#This function is used to create a PowerShell ErrorRecord
	[CmdletBinding()]
	Param
	(
		[Parameter(Mandatory=$true,Position=1)][String]$ExceptionString,
		[Parameter(Mandatory=$true,Position=2)][String]$ErrorID,
		[Parameter(Mandatory=$true,Position=3)][System.Management.Automation.ErrorCategory]$ErrorCategory,
		[Parameter(Mandatory=$true,Position=4)][PSObject]$TargetObject
	)
	Process
	{
		$exception = New-Object System.Management.Automation.RuntimeException($ExceptionString)
		$customError = New-Object System.Management.Automation.ErrorRecord($exception,$ErrorID,$ErrorCategory,$TargetObject)
		return $customError
	}
}

Function Connect-OSCEXOWebService
{
	#.EXTERNALHELP Connect-OSCEXOWebService-Help.xml

    [cmdletbinding()]
	Param
	(
		#Define parameters
		[Parameter(Mandatory=$true,Position=1)]
		[System.Management.Automation.PSCredential]$Credential,
		[Parameter(Mandatory=$false,Position=2)]
		[Microsoft.Exchange.WebServices.Data.ExchangeVersion]$ExchangeVersion="Exchange2010_SP2",
		[Parameter(Mandatory=$false,Position=3)]
		[string]$TimeZoneStandardName,
		[Parameter(Mandatory=$false)]
		[switch]$Force
	)
	Process
	{
        #Get specific time zone info
        if (-not [System.String]::IsNullOrEmpty($TimeZoneStandardName)) {
            Try
            {
                $tzInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZoneStandardName)
            }
            Catch
            {
                $PSCmdlet.ThrowTerminatingError($_)
            }
        } else {
            $tzInfo = $null
        }

		#Create the callback to validate the redirection URL.
		$validateRedirectionUrlCallback = {
            param ([string]$Url)
			if ($Url -eq "https://autodiscover-s.outlook.com/autodiscover/autodiscover.xml") {
            	return $true
			} else {
				return $false
			}
        }	
	
		#Try to get exchange service object from global scope
		$existingExSvcVar = (Get-Variable -Name exService -Scope Global -ErrorAction:SilentlyContinue) -ne $null
		
		#Establish the connection to Exchange Web Service
		if ((-not $existingExSvcVar) -or $Force) {
			$verboseMsg = $Messages.EstablishConnection
			$PSCmdlet.WriteVerbose($verboseMsg)
            if ($tzInfo -ne $null) {
                $exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
				    		 [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion,$tzInfo)			
            } else {
                $exService = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(`
				    		 [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::$ExchangeVersion)
            }
			
			#Set network credential
			$userName = $Credential.UserName
			$exService.Credentials = $Credential.GetNetworkCredential()
			Try
			{
				#Set the URL by using Autodiscover
				$exService.AutodiscoverUrl($userName,$validateRedirectionUrlCallback)
				$verboseMsg = $Messages.SaveExWebSvcVariable
				$PSCmdlet.WriteVerbose($verboseMsg)
				Set-Variable -Name exService -Value $exService -Scope Global -Force
			}
			Catch [Microsoft.Exchange.WebServices.Autodiscover.AutodiscoverRemoteException]
			{
				$PSCmdlet.ThrowTerminatingError($_)
			}
			Catch
			{
				$PSCmdlet.ThrowTerminatingError($_)
			}
		} else {
			$verboseMsg = $Messages.FindExWebSvcVariable
            $verboseMsg = $verboseMsg -f $exService.Credentials.Credentials.UserName
			$PSCmdlet.WriteVerbose($verboseMsg)            
		}
	}
}

Function Set-OSCEXOEmailSignature
{
	#.EXTERNALHELP Set-OSCEXOEmailSignature-Help.xml

    [cmdletbinding()]
	Param
	(
		#Define parameters
		[Parameter(Mandatory=$false,Position=1)]
		[string]$HtmlSignature,
		[Parameter(Mandatory=$false,Position=2)]
		[string]$TextSignature,
		[Parameter(Mandatory=$false)]
		[switch]$Force
	)
	Begin
	{
        #Verify the existence of exchange service object
        if ($exService -eq $null) {
			$errorMsg = $Messages.RequireConnection
			$customError = New-OSCPSCustomErrorRecord `
			-ExceptionString $errorMsg `
			-ErrorCategory NotSpecified -ErrorID 1 -TargetObject $PSCmdlet
			$PSCmdlet.ThrowTerminatingError($customError)
        }
	}
    Process
    {
        #Get user configuration
        $verboseMsg = $Messages.GetOWAUserOptions
        $PSCmdlet.WriteVerbose($verboseMsg)
        $owaUserOptions = [Microsoft.Exchange.WebServices.Data.UserConfiguration]::Bind(`
                          $exService,"OWA.UserOptions",`
                          [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,`
                          [Microsoft.Exchange.WebServices.Data.UserConfigurationProperties]::All)

        $verboseMsg = $Messages.SetOWASignature
        $PSCmdlet.WriteVerbose($verboseMsg)
		
		#Set HTML signature
        if (-not $owaUserOptions.Dictionary.ContainsKey("signaturehtml")) {
            if (-not [System.String]::IsNullOrEmpty($HtmlSignature)) {
                $owaUserOptions.Dictionary.Add("signaturehtml",$HtmlSignature)
            }
        } else {
            if (-not [System.String]::IsNullOrEmpty($HtmlSignature)) {
                if ($Force) {
                    $owaUserOptions.Dictionary["signaturehtml"] = $HtmlSignature
                } else {
                    $warningMsg = $Messages.ExistingHtmlSignature
                    $PSCmdlet.WriteWarning($warningMsg)
                }
            } else {
                $owaUserOptions.Dictionary.Remove("signaturehtml") | Out-Null
            }
        }
		
		#Set text signature
        if (-not $owaUserOptions.Dictionary.ContainsKey("signaturetext")) {
            if (-not [System.String]::IsNullOrEmpty($TextSignature)) {
                $owaUserOptions.Dictionary.Add("signaturetext",$TextSignature)
            }
        } else {
            if (-not [System.String]::IsNullOrEmpty($TextSignature)) {
                if ($Force) {
                    $owaUserOptions.Dictionary["signaturetext"] = $TextSignature
                } else {
                    $warningMsg = $Messages.ExistingTextSignature
                    $PSCmdlet.WriteWarning($warningMsg)
                }
            } else {
                $owaUserOptions.Dictionary.Remove("signaturetext") | Out-Null
            }
        }
		
		#Update the user configuration by applying local changes to the Exchange server
        $owaUserOptions.Update()
    }
	End {}
}

Export-ModuleMember -Function "Connect-OSCEXOWebService","Set-OSCEXOEmailSignature"