###############################################################################
# M365 PS Profle 
# Installs and Updates the Required PowerShell Modules for M365 Management
###############################################################################
# Install-Module Microsoft.PowerShell.PSResourceGet -Scope CurrentUser
#
# Remove-Module M365PSProfile
# Import-Module C:\GIT_WorkingDir\M365PSProfile\M365PSProfile.psd1
# Install-M365Modules -Scope CurrentUser



##############################################################################
# Function AsciiArt
##############################################################################
Function Invoke-AsciiArt {
Write-Host "__  __ ____    __ _____ _____   _____ _____            __ _ _      "
Write-Host "|  \/  |___ \  / /| ____|  __ \ / ____|  __ \          / _(_) |     "
Write-Host "| \  / | __) |/ /_| |__ | |__) | (___ | |__) | __ ___ | |_ _| | ___ "
Write-Host "| |\/| ||__ <| '_ \___ \|  ___/ \___ \|  ___/ '__/ _ \|  _| | |/ _ \"
Write-Host "| |  | |___) | (_) |__) | |     ____) | |   | | | (_) | | | | |  __/"
Write-Host "|_|  |_|____/ \___/____/|_|    |_____/|_|   |_|  \___/|_| |_|_|\___|"														
}




##############################################################################
# Uninstall-M365Modules
# Remove Modules in -Modules Parameter
##############################################################################
Function Uninstall-M365Modules {
		<#
		.SYNOPSIS
		Uninstall M365 PowerShell Modules

		.DESCRIPTION
		Update and cleanup of all defined M365 modules

		.PARAMETER Modules
		Array of Module Names that will be installed

		.PARAMETER Scope
		Sets the Scope [CurrentUser/AllUsers] for the Installation of the PowerShell Modules

		.EXAMPLE
		Install-M365Modules

		.EXAMPLE
		Install-M365Modules -Modules "Az","MSOnline","PnP.PowerShell","Microsoft.Graph" -Scope [CurrentUser/AllUsers]

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>

	param (
		[Parameter(Mandatory=$True)][array]$Modules,
		[Parameter(Mandatory=$True)][string]$Scope
	)

	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
	$IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	<# Uninstall still has to be coded #>
}

##############################################################################
# Invoke-InstallM365Modules
# Remove old Module instead of only install new Version
##############################################################################


##############################################################################
# Remove existing PS Connections
##############################################################################
Function Disconnect-All {
	Get-PSSession | Remove-PSSession
	Try {
		Disconnect-AzureAD -ErrorAction SilentlyContinue
		Disconnect-MsolService -ErrorAction SilentlyContinue
		Disconnect-SPOService -ErrorAction SilentlyContinue
		Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
		Disconnect-ExchangeOnline -confirm:$false -ErrorAction SilentlyContinue
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	} catch {
		#Missing Error Handling
	}
}


#############################################################################
# Set-WindowTitle Function
#############################################################################
Function Set-WindowTitle {
	PARAM (
		[string]$Title = "Windows PowerShell"
	)
	$host.ui.RawUI.WindowTitle = $Title
}

##############################################################################
# Main Program
##############################################################################

function Install-M365Modules {


<# 
.SYNOPSIS
M365PSProfile installs and keeps the PowerShell Modules needed for Microsoft 365 Management up to date.
It provides a simple way to add it to the PowerShell Profile
	
.DESCRIPTION
M365PSProfile installs and keeps the PowerShell Modules needed for Microsoft 365 Management up to date.
It provides a simple way to add it to the PowerShell Profile

.PARAMETER Modules
[array]$Modules = @(<ModuleName1>,<Modulename2>)
[array]$Modules = @("AZ", "MSOnline", "AzureADPreview", "ExchangeOnlineManagement", "Icewolf.EXO.SpamAnalyze", "MicrosoftTeams", "Microsoft.Online.SharePoint.PowerShell", "PnP.PowerShell" , "ORCA", "O365CentralizedAddInDeployment", "MSCommerce", "WhiteboardAdmin", "Microsoft.Graph", "Microsoft.Graph.Beta", "MSAL.PS", "MSIdentityTools" )

.PARAMETER Scope
Sets the Scope [CurrentUser/AllUsers] for the Installation of the PowerShell Modules

.PARAMETER UpdateCheckDays
Sets the Days for the Update Check

.PARAMETER AsciiArt
[bool]AsciiArt controls the AsciiArt Screen at the Start

.LINK
https://github.com/fabrisodotps1/M365PSProfile

.EXAMPLE
#Installs and updates the Default Modules in CurrentUser Scope
Install-M365Modules

.EXAMPLE
#Installs and updates the specified Modules
Install-M365Modules -Modules @("ExchangeOnlineManagement", "MicrosoftTeams", "Microsoft.Online.SharePoint.PowerShell", "PnP.PowerShell") -Scope [CurrentUser|AllUsers]

.EXAMPLE
#Installs and updates the specified Modules without showing AsciiArt at the Start
Install-M365Modules -Modules @("ExchangeOnlineManagement", "MicrosoftTeams", "Microsoft.Online.SharePoint.PowerShell", "PnP.PowerShell") -Scope [CurrentUser|AllUsers] -AsciiArt $False
#>

	#Parameter for the Module
	param(
		[parameter(mandatory=$false)][array]$Modules = @("AZ", "MSOnline", "AzureADPreview", "ExchangeOnlineManagement", "Icewolf.EXO.SpamAnalyze", "MicrosoftTeams", "Microsoft.Online.SharePoint.PowerShell", "PnP.PowerShell" , "ORCA", "O365CentralizedAddInDeployment", "MSCommerce", "WhiteboardAdmin", "Microsoft.Graph", "Microsoft.Graph.Beta", "MSAL.PS", "MSIdentityTools"),
		[parameter(mandatory=$false)][string]$Scope = "CurrentUser",
		[parameter(mandatory=$false)][bool]$AsciiArt = $true,
		[parameter(mandatory=$false)][int]$UpdateCheckDays = 7

		)


	Write-Host "Starting M365PSProfile..."
	If ($AsciiArt -eq $true)
	{
		#Show AsciArt
		Invoke-AsciiArt
	}

	#Get Current User / Is Admin
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
	$IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	
	#Install-Module Microsoft.PowerShell.PSResourceGet -Scope CurrentUser
	Import-Module  Microsoft.PowerShell.PSResourceGet
	$PSGallery = Get-PSResourceRepository -Name PSGallery	
	If ($PSGallery.Trusted -eq $false)
	{
		Write-Host "Warning: PSGallery is not Trusted" -ForegroundColor Yellow
		#Set-PSResourceRepository -Name PSGallery -Trusted:$true
	}


	#Update Logic
	#Has to be coded
	$UpdateCheck = $true
	If ($UpdateCheck -eq $true)
	{
		
		#Check if VSCode or PowerShell is running
		[array]$process = Get-Process | Where-Object {$_.ProcessName -eq "powershell" -or $_.ProcessName -eq "pwsh" -or $_.ProcessName -eq "code"}
		#$process = Get-Process -Name code -ErrorAction SilentlyContinue
		If ($process.count -gt 1)
		{
			Write-Host "PowerShell or Visual Studio Code running? Please close it, otherwise the Modules sometimes can't be updated..." -ForegroundColor Red
			$process
			#Press any key to continue
			Write-Host 'Press any key to continue...';
			$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
		}


		Write-Host "Checking Modules..."
		Foreach ($Module in $Modules) 
		{
			#Get Array of installed Modules
			[Array]$InstalledModules = Get-InstalledPSResource -Name $Module -Scope $Scope -ErrorAction SilentlyContinue | Sort-Object Version -Descending

			If ($Null -eq $InstalledModules) 
			{
				#Module not found
				Write-Host "$Module Module not found. Try to install..."
				If ($IsAdmin -eq $false -and $Scope -eq "AllUsers") 
				{
					Write-Host "WARNING: PS must be running <As Administrator> to install the Module" -ForegroundColor Red				
				} else {
					#Install-Module $Module -Confirm:$false
					Install-PSResource $Module -Scope $Scope -TrustRepository
				}
			} else {
				#Module found

				#Get Module from PowerShell Gallery
				$PSGalleryModule = Find-PSResource -Name $Module
				$PSGalleryVersion = $PSGalleryModule.Version.ToString()

				#Check if Multiple Modules are installed
				If (($InstalledModules.count) -gt 1) {

					Write-host "WARNIN: $Module > Multiple Versions found. Uninstall old Versions? (Default is Yes)" -ForegroundColor Yellow 
					$Readhost = Read-Host " ( y / n ) " 
					Switch ($ReadHost) 
					{ 
						Y {
							#Uninstall all Modules
							Write-Host "INFO: Uninstall Module"
							Uninstall-PSResource -Name $Module -Scope $Scope -SkipDependencyCheck
							
							#Install newest Module
							Write-Host "INFO: Install newest Module $Module $PSGalleryVersion"
							Install-PSResource -Name $Module -Scope $Scope -TrustRepository
						}				
						N { Write-Host "Skip Uninstall old Modules" }
						Default {
							#Uninstall all Modules
							Write-Host "INFO: Uninstall Module"
							Uninstall-PSResource -Name $Module -Scope $Scope -SkipDependencyCheck

							#Install newest Module
							Write-Host "INFO: Install newest Module $Module $PSGalleryVersion"
							Install-PSResource -Name $Module -Scope $Scope -TrustRepository
						}
					}
				} else {
					#Only one Module found

					#Version Check 
					If ($PSGalleryVersion -gt $InstalledModules.Version.ToString() )
					{
						#Uninstall Module
						Uninstall-PSResource -Name $Module -Scope $Scope -SkipDependencyCheck
						#Install Module
						Install-PSResource -Name $Module -Scope $Scope -TrustRepository
					} else {
						#Write Module Name
						Write-Host "Checking Module: $Module $($InstalledModules.Version.ToString())" -ForegroundColor Green
					}				
				}
			}
		}

	}
}

