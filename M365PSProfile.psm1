###############################################################################
# M365 PS Profle 
# Installs and Updates the Required PowerShell Modules for M365 Management
###############################################################################
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
# Invoke-UninstallM365Modules
# Remove Modules
##############################################################################
Function Invoke-UninstallM365Modules {
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
Function Invoke-InstallM365Modules {
	<#
		.SYNOPSIS
		Install and Update M365 PowerShell Modules

		.DESCRIPTION
		Update and cleanup of all defined M365 PowerShell Modules

		.PARAMETER Modules
		Array of the PowerShell Module Names that will be installed

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

	
	#Install-Module Microsoft.PowerShell.PSResourceGet -Scope CurrentUser
	Import-Module  Microsoft.PowerShell.PSResourceGet
	$PSGallery = Get-PSResourceRepository -Name PSGallery	
	If ($PSGallery.Trusted -eq $false)
	{
		Write-Host "Warning: PSGallery is not Trusted" -ForegroundColor Yellow
		#Set-PSResourceRepository -Name PSGallery -Trusted:$true
	}


	#Update Check = $False
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
				Install-PSResource $Module -Scope $Scope
			}
		} else {
			#Module found

			#Get Module from PowerShell Gallery
			$PSGalleryModule = Find-PSResource -Name $Module
			$PSGalleryVersion = $PSGalleryModule.Version.ToString()

			#Check if Multiple Modules are installed
			If (($InstalledModules.count) -gt 1) {

				Write-host "INFO: Multiple Modules found. Uninstall old Modules? (Default is Yes)" -ForegroundColor Yellow 
				$Readhost = Read-Host " ( y / n ) " 
				Switch ($ReadHost) 
				{ 
					Y {
						#Uninstall all Modules
						For ($i = 0; $i -lt $InstalledModule.count; $i++) {
							$Version = $InstalledModules[$i].Version.ToString()
							Write-Host "INFO: Uninstall Module $Version"
							Uninstall-PSResource -Name $Module -Scope $Scope -Force
						} 
						#Install newest Module
						Write-Host "INFO: Install newest Module $Module $PSGalleryVersion"
						Install-PSResource -Name $Module -Scope $Scope
					}				
					N { Write-Host "Skip Uninstall old Modules" }
					Default {
						#Uninstall all Modules
						For ($i = 0; $i -lt $InstalledModule.count; $i++) {
							$Version = $InstalledModules[$i].Version.ToString()
							Write-Host "INFO: Uninstall Module $Version"
							Uninstall-PSResource -Name $Module -Scope $Scope -Force
						} 
						#Install newest Module
						Write-Host "INFO: Install newest Module $Module $PSGalleryVersion"
						Install-PSResource -Name $Module -Scope $Scope
					}
				}
			} else {
				#Only one Module found

				#Version Check 
				If ($PSGalleryVersion -gt $InstalledModules.Version.ToString() )
				{
					#Uninstall Module
					Uninstall-PSResource -Name $Module -Scope $Scope -Force
					#Install Module
					Install-PSResource -Name $Module -Scope $Scope
				} else {
					#Write Module Name
					Write-Host "Checking Module: $Module $InstalledModules.Version.ToString() "
				}				
			}
		}
	}
}


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
		[parameter(mandatory=$false)][bool]$UpdateCheckDays = "7"

		)


	Write-Host "Starting M365PSProfile..."
	If ($AsciiArt -eq $true)
	{
		#Show AsciArt
		Invoke-AsciiArt
	}

	<#
	$pshost = get-host
	$pswindow = $pshost.ui.rawui
	$LanguageMode = $ExecutionContext.SessionState.LanguageMode
	If ($LanguageMode -eq "Fulllanguage") {
		if ($pswindow.WindowSize.Width -lt 220) {
			if ($env:WT_SESSION) {
				#Windows Terminal
				$Buffersize = $pswindow.buffersize
				$Buffersize.height = 8500
				$Buffersize.width = 220
				$pswindow.buffersize = $Buffersize
				$Windowsize = $pswindow.windowsize
				$Windowsize.width = 150
				$windowsize.height = 40
				$pswindow.windowsize = $windowsize
				#$pswindow
			}
			else {
				if ($env:TERM_PROGRAM -eq 'vscode') {
					#If vscode do nothing
				}
				else {
					#Normal PowerShell Window
					$Buffersize = $pswindow.buffersize
					$Buffersize.height = 8500
					$Buffersize.width = 220
					$pswindow.buffersize = $Buffersize
					$Windowsize = $pswindow.windowsize
					$Windowsize.width = 150
					$windowsize.height = 60
					$pswindow.windowsize = $windowsize
				}
			}
		}
		#>


		#Update Logic / How often the Install and Update Check will be invoked
		#Every x Days?
		#Registry Key > Will work for Windows, what about Linux?
		#File

		#Install-Modules
		Invoke-InstallM365Modules -Modules $Modules -Scope $Scope 	
}

