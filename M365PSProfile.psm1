##############################################################################
# M365 PS Profile 
# Installs and Updates the Required PowerShell Modules for M365 Management
##############################################################################

##############################################################################
# Global variable for standard modules
##############################################################################
[array]$global:M365StandardModules = @(
	"ExchangeOnlineManagement",
	"Icewolf.EXO.SpamAnalyze",
	"MicrosoftTeams",
	"Microsoft.Online.SharePoint.PowerShell",
	"PnP.PowerShell",
	"ORCA",
	"O365CentralizedAddInDeployment",
	"M365PSProfile"
	"MSCommerce",
	"WhiteboardAdmin",
	"Microsoft.Graph",
	"Microsoft.Graph.Beta",
	"MSIdentityTools",
	"PSMSALNet"
)

##############################################################################
# Get-M365StandardModules
# Returns the M365StandardModules global variable 
##############################################################################
Function Get-M365StandardModule {
	<#
		.SYNOPSIS
		Returns the M365StandardModules global variable.

		.DESCRIPTION
		Returns the M365StandardModules global variable which contains the standard modules for M365 Management.

		.EXAMPLE
		Get-M365StandardModule

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>

	return $global:M365StandardModules
}

##############################################################################
# Function AsciiArt
##############################################################################
Function Invoke-AsciiArt {
	<#
		.SYNOPSIS
		Generates M365PSProfile AsciiArt

		.DESCRIPTION
		Generates M365PSProfile AsciiArt

		.EXAMPLE
		Invoke-AsciiArt

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>

	Write-Host " __  __ ____    __ _____ _____   _____ _____            __ _ _      "
	Write-Host "|  \/  |___ \  / /| ____|  __ \ / ____|  __ \          / _(_) |     "
	Write-Host "| \  / | __) |/ /_| |__ | |__) | (___ | |__) | __ ___ | |_ _| | ___ "
	Write-Host "| |\/| ||__ <| '_ \___ \|  ___/ \___ \|  ___/ '__/ _ \|  _| | |/ _ \"
	Write-Host "| |  | |___) | (_) |__) | |     ____) | |   | | | (_) | | | | |  __/"
	Write-Host "|_|  |_|____/ \___/____/|_|    |_____/|_|   |_|  \___/|_| |_|_|\___|"
}

##############################################################################
# Add-M365PSProfile
# Add M365PSProfile, if no PowerShell Profile exists
##############################################################################
Function Add-M365PSProfile {
	<#
		.SYNOPSIS
		Add PowerShell Profile with M365PSProfile setup

		.DESCRIPTION
		Add PowerShell Profile with M365PSProfile setup.

		Needs to be executed separately for PowerShell v5 and v7.

		.EXAMPLE
		Add-M365PSProfile

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>
	[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    	param()

	$ProfileContent = @"

#M365PSProfile: Install or updates the default Modules (what we think every M365 Admin needs) in the CurrentUser Scope
Import-Module -Name M365PSProfile
Install-M365Module
"@

	If (-not(Test-Path -Path $Profile)) {
		Write-Host "No PowerShell Profile exists. A new Profile with the M365PSProfile setup is created."
		$ProfileContent | Out-File -FilePath $Profile -Encoding utf8 -Force
	} else {
		Write-Host "PowerShell Profile already exists. You could manually add the following code to your PowerShell profile: `r`n" -ForegroundColor Yellow
		Write-Host $ProfileContent -ForegroundColor Magenta
		Write-Host "`r`nThis command can also add the code to the end of your profile file ($($Profile))."
		
		if ($PSCmdlet.ShouldContinue($Path, "Add ""Install-M365Module"" to the PowerShell Profile file $($Profile)")) {
			Write-Host "Adding to profile..."
			Add-Content -Path $Profile -Value $ProfileContent -Encoding utf8
		} else {
			Write-Host "Okay, not changing your profile."
		}
	}
}

##############################################################################
# Uninstall-M365Modules
# Remove Modules in -Modules Parameter
##############################################################################
Function Uninstall-M365Module {
	<#
		.SYNOPSIS
		Uninstall M365 PowerShell Modules

		.DESCRIPTION
		Uninstall of all defined M365 modules

		.PARAMETER Modules
		Array of Module Names that will be uninstalled. Default value are the default modules (see Get-M365StandardModule) or an Array with the Modules to uninstall.

		.PARAMETER Scope
		Sets the Scope [CurrentUser/AllUsers] for the Installation of the PowerShell Modules. Default value is CurrentUser.

		.EXAMPLE
		Uninstall-M365Modules

		.EXAMPLE
		Uninstall-M365Modules -Modules "Az","MSOnline","PnP.PowerShell","Microsoft.Graph" -Scope CurrentUser

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>

	param (
		[Parameter(Mandatory = $false)][array]$Modules = $global:M365StandardModules,
		[parameter(mandatory = $false)][ValidateSet("CurrentUser", "AllUsers")][string]$Scope = "CurrentUser"
	)

	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
	$IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	foreach ($Module in $Modules) {
		[Array]$InstalledModules = Get-InstalledPSResource -Name $Module -Scope $Scope -ErrorAction SilentlyContinue | Sort-Object Version -Descending

		if ($InstalledModules) {
			# Module found
			if (($IsAdmin -eq $false) -and ($Scope -eq "AllUsers")) {
				Write-Host "WARNING: PS must be running <As Administrator> to uninstall the Module" -ForegroundColor Red
			} else {
				# Uninstall all versions of the module
				Write-Host "Uninstall Module: $Module $($InstalledModules.Version.ToString())" -ForegroundColor Yellow
				Uninstall-PSResource -Name $Module -Scope $Scope -SkipDependencyCheck -WarningAction SilentlyContinue
			}
		}
	}
}

##############################################################################
# Remove existing PS Connections
##############################################################################
Function Disconnect-All {
	<#
		.SYNOPSIS
		Disconnect all Connections to Microsoft 365 Services

		.DESCRIPTION
		Disconnect all Connections of the Modules MicrosoftTeams, ExchangeOnlineManagement, Microsoft.Online.SharePoint.PowerShell, Microsoft.Graph and removes remote PS Sessions

		.EXAMPLE
		Disconnect-All

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>

	Get-PSSession | Remove-PSSession

	Disconnect-SPOService -ErrorAction SilentlyContinue
	Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
	Disconnect-ExchangeOnline -confirm:$false -ErrorAction SilentlyContinue
	Disconnect-MgGraph -ErrorAction SilentlyContinue
}

#############################################################################
# Set-WindowTitle Function
#############################################################################
Function Set-WindowTitle {
	<#
		.SYNOPSIS
		Set the Window Title

		.DESCRIPTION
		Set the Window Title

		.EXAMPLE
		Set-WindowTitle

		.EXAMPLE
		Set-WindowTitle -Title "My Title"

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>

	PARAM (
		[string]$Title = "Windows PowerShell"
	)
	$host.ui.RawUI.WindowTitle = $Title
}

##############################################################################
# Main Program
##############################################################################

Function Install-M365Module {
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

		.PARAMETER $RunInVSCode
		[bool]$RunInVSCode controls if the Script will run in VSCode [Default is $false]

		.PARAMETER Repository
		[string]Repository specifies which PowerShell Repository should be used [Default is PSGallery]

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

		.EXAMPLE
		#Installs and updates the Default Modules in CurrentUser Scope and use a custom repository called "MyRepo"
		Install-M365Modules -Repository "MyRepo"
	#>

	#Parameter for the Module
	param(
		[parameter(mandatory = $false)][array]$Modules = $global:M365StandardModules,
		[parameter(mandatory = $false)][ValidateSet("CurrentUser", "AllUsers")][string]$Scope = "CurrentUser",
		[parameter(mandatory = $false)][bool]$AsciiArt = $true,
		[parameter(mandatory = $false)][bool]$RunInVSCode = $false,
		[parameter(mandatory = $false)][string]$Repository = "PSGallery"
	)

	#Check if it is running in VSCode
	if ($env:TERM_PROGRAM -eq 'vscode') {
		If ($RunInVSCode -eq $false) {
			Exit
		}
	}	

	If ($AsciiArt -eq $true) {
		#Show AsciArt
		Invoke-AsciiArt
	}

	#Get Current User / Is Admin
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
	$IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	Import-Module  Microsoft.PowerShell.PSResourceGet

	if($Repository -eq "PSGallery") {
		$PSGallery = Get-PSResourceRepository -Name PSGallery
		If ($PSGallery.Trusted -eq $false) {
			Write-Host "Warning: PSGallery is not Trusted" -ForegroundColor Yellow
		}
	}
		
	#Check if VSCode or PowerShell is running
	[array]$process = Get-Process | Where-Object { $_.ProcessName -eq "powershell" -or $_.ProcessName -eq "pwsh" -or $_.ProcessName -eq "code" }
	
	If ($process.count -gt 1) {
		Write-Host "PowerShell or Visual Studio Code running? Please close it, Modules in use can't be updated..." -ForegroundColor Yellow
		$process
		#Press any key to continue
		Write-Host 'Press any key to continue...';
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
	}

	Write-Host "Checking Modules..."
	#Check Microsoft.PowerShell.PSResourceGet
	#Can't uninstall loaded DLL's so you have to uninstall next time you start PowerShell
	#[System.AppDomain]::CurrentDomain.GetAssemblies() | where {$_.Location -match "Microsoft.PowerShell.PSResourceGet"}
	$Module = "Microsoft.PowerShell.PSResourceGet"
	[Array]$InstalledModules = Get-InstalledPSResource -Name $Module -Scope $Scope -ErrorAction SilentlyContinue | Sort-Object Version -Descending

	Write-Host "Checking Module: $Module $($InstalledModules[0].Version.ToString())" -ForegroundColor Green

	If ($InstalledModules.Count -gt 1)
	{
		$Version = $InstalledModules[$InstalledModules.Count - 1].Version
		Write-Host "Uninstall Module $Module $Version" -ForegroundColor Yellow
		Uninstall-PSResource -Name $Module -Scope $Scope -Version $Version -SkipDependencyCheck
	} else {
		#Only one Version found
		[System.Version]$InstalledModuleVersion = $($InstalledModules.Version.ToString())

		#Get Module from PowerShell Gallery (or another repository if specified)
		$PSGalleryModule = Find-PSResource -Name $Module -Repository $Repository
		$PSGalleryVersion = $PSGalleryModule.Version.ToString()

		#Version Check 
		If ($PSGalleryVersion -gt $InstalledModuleVersion) 
		{
			Write-Host "Install newest Module $Module $PSGalleryVersion" -ForegroundColor Yellow
			Install-PSResource $Module -Scope $Scope -TrustRepository -WarningAction SilentlyContinue -Repository $Repository
		}
	}


	Foreach ($Module in $Modules) {
		#Get Array of installed Modules
		[Array]$InstalledModules = Get-InstalledPSResource -Name $Module -Scope $Scope -ErrorAction SilentlyContinue | Sort-Object Version -Descending

		If ($Null -eq $InstalledModules) {
			#Module not found
			Write-Host "$Module Module not found. Try to install..."
			If ($IsAdmin -eq $false -and $Scope -eq "AllUsers") {
				Write-Host "WARNING: PS must be running <As Administrator> to install the Module" -ForegroundColor Red
			} else {
				#Only one Version found
				[System.Version]$InstalledModuleVersion = $($InstalledModules.Version.ToString())

				#Get Module from PowerShell Gallery
				$PSGalleryModule = Find-PSResource -Name $Module -Repository $Repository #-Prerelease
				$PSGalleryVersion = $PSGalleryModule.Version.ToString()
				Write-Host "Install newest Module $Module $PSGalleryVersion" -ForegroundColor Yellow

				Install-PSResource $Module -Scope $Scope -TrustRepository -WarningAction SilentlyContinue -Repository $Repository #-Prerelease
			}
		} else {
			#Module found

			#Get Module from PowerShell Gallery (or another repository if specified)
			$PSGalleryModule = Find-PSResource -Name $Module -Repository $Repository #-Prerelease
			[System.Version]$PSGalleryVersion = $PSGalleryModule.Version.ToString()

			#Check if Multiple Modules are installed
			If (($InstalledModules.count) -gt 1) {

				Write-Host "WARNING: $Module > Multiple Versions found. Uninstall old Versions? (Default is Yes)" -ForegroundColor Yellow 
				$Readhost = Read-Host " ( y / n ) " 
				Switch ($ReadHost) { 
					Y {
						#Uninstall all Modules
						Write-Host "Uninstall Module"
						Uninstall-PSResource -Name $Module -Scope $Scope -SkipDependencyCheck
						
						#Install newest Module
						Write-Host "Install newest Module $Module $PSGalleryVersion" -ForegroundColor Yellow
						Install-PSResource -Name $Module -Scope $Scope -TrustRepository -WarningAction SilentlyContinue -Repository $Repository #-Prerelease
					}				
					N {
						Write-Host "Skip Uninstall old Modules" 
					}
					Default {
						#Uninstall all Modules
						Write-Host "Uninstall Module"
						Uninstall-PSResource -Name $Module -Scope $Scope -SkipDependencyCheck

						#Install newest Module
						Write-Host "Install newest Module $Module $PSGalleryVersion" -ForegroundColor Yellow
						Install-PSResource -Name $Module -Scope $Scope -TrustRepository -WarningAction SilentlyContinue -Repository $Repository #-Prerelease
					}
				}
			} else {
				#Only one Module found
				[System.Version]$InstalledModuleVersion = $($InstalledModules.Version.ToString())

				#Version Check 
				If ($PSGalleryVersion -gt $InstalledModuleVersion) {
					#Uninstall Module
					Write-Host "Uninstall Module: $Module $($InstalledModules.Version.ToString())" -ForegroundColor Yellow
					Uninstall-PSResource -Name $Module -Scope $Scope -SkipDependencyCheck

					#If AZ also Uninstall all AZ.* Modules
					If ($Module -eq "AZ")
					{
						Write-Host "Uninstall AZ.* Modules" -ForegroundColor Yellow
						Uninstall-PSResource AZ.* -Scope $Scope -SkipDependencyCheck
					}

					#If Microsoft.Graph also Uninstall all Microsoft.Graph.* Modules
					If ($Module -eq "Microsoft.Graph")
					{
						Write-Host "Uninstall Microsoft.Graph.* Modules" -ForegroundColor Yellow
						Get-InstalledPSResource -Name Microsoft.Graph.* -Scope CurrentUser | Where-Object {$_.Name -notmatch "Microsoft.Graph.Beta"} | Uninstall-PSResource -SkipDependencyCheck
					}

					#If Microsoft.Graph.Beta also Uninstall all Microsoft.Graph.Beta.* Modules
					If ($Module -eq "Microsoft.Graph.Beta")
					{
						Write-Host "Uninstall Microsoft.Graph.Beta.* Modules" -ForegroundColor Yellow
						Get-InstalledPSResource -Name Microsoft.Graph.Beta* -Scope CurrentUser | Uninstall-PSResource -SkipDependencyCheck
					}

					#Install Module
					Write-Host "Install Module: $Module $PSGalleryVersion" -ForegroundColor Yellow
					Install-PSResource -Name $Module -Scope $Scope -TrustRepository -WarningAction SilentlyContinue -Repository $Repository #-Prerelease
				} else {
					#Write Module Name
					Write-Host "Checking Module: $Module $($InstalledModules.Version.ToString())" -ForegroundColor Green
				}
			}
		}
	}
}

##############################################################################
# Import Module
##############################################################################
If (-not(Test-Path -Path $Profile)) 
{
	Write-Host "No PowerShell Profile exists. You can add the M365Profile Update check with ADD-M365PSProfile" -ForegroundColor Yellow
	
} else {
	$Content = Get-Content -Path $PROFILE
	If ($Content -match "Install-M365Module")
	{
		#Match found
	} else {
		#No Match found
		Write-Host  "You have a PowerShell Profile. You can add the M365Profile Update check by adding: Install-M365Module" -ForegroundColor Yellow
	}
}


