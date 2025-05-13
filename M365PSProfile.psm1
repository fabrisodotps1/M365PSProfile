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
	"MicrosoftPlaces",
	"MSIdentityTools",
	"PSMSALNet"
)

##############################################################################
# Get-M365ModulePath
# Returns the Path for the Modules
##############################################################################
Function Get-M365ModulePath {
	<#
		.SYNOPSIS
		Returns the Path for the Modules based ont the Scope Parameter

		.DESCRIPTION
		Returns the Path for the Modules based ont the Scope Parameter

		.PARAMETER Scope
		Sets the Scope [CurrentUser/AllUsers] for the Installation of the PowerShell Modules. Default value is CurrentUser.

		.EXAMPLE
		Get-M365ModulePath -Scope CurrentUser

		.EXAMPLE
		Get-M365ModulePath -Scope AllUsers

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>

	PARAM(
		[parameter(mandatory = $false)][ValidateSet("CurrentUser", "AllUsers")][string]$Scope = "CurrentUser"
	)

	$Personal = [environment]::getfolderpath("mydocuments")
	$ProgramFiles = [environment]::getfolderpath("ProgramFiles")
	If ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows))
	{
		#Windows
		If ($Host.Version -ge "6.0")
		{
			$Path = "PowerShell"
		} else {
			$path = "WindowsPowerShell"
		}

		If ($Scope -eq "CurrentUser")
		{
			$LocalUserDir = Join-Path -Path $Personal -ChildPath $Path
			return $LocalUserDir + "\Modules\"
		}

		If ($Scope -eq "AllUsers")
		{
			$AllUsersDir = Join-Path -Path $ProgramFiles -ChildPath $Path
			return $AllUsersDir + "\Modules\"
		}
	}

	#Unix / OSX
	#$LocalUserDir = Join-Path -Path $Env:Home -ChildPath ".local", "share", "powershell"
	#$AllUsersDir = Join-Path -Path "/usr" -ChildPath "local", "share", "powershell"
	#Return $LocalUserDir, $AllUsersDir
}

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
	Write-Host "Version:" $MyInvocation.MyCommand.ScriptBlock.Module.Version
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

	$M365PSProfileContent = @"

#M365PSProfile: Install or updates the default Modules (what we think every M365 Admin needs) in the CurrentUser Scope
Import-Module -Name M365PSProfile
Install-M365Module
"@

	If (-not(Test-Path -Path $Profile)) {
		#No Profile found
		Write-Host "No PowerShell Profile exists. A new Profile with the M365PSProfile setup is created."
		$M365PSProfileContent | Out-File -FilePath $Profile -Encoding utf8 -Force
	} else {
		#Profile found
		$ProfileContent = Get-Content -Path $Profile -Encoding utf8
		#$ProfileContent | Where-Object {$_ -match "Import-Module -Name M365PSProfile"}
		$Match = $ProfileContent | Where-Object {$_ -match "Install-M365Module"}
		If ($Null -ne $Match)
		{
			#M365PSProfile already in Profile
			Write-Host "PowerShell Profile already exists. M365PSProfile is already in the Profile." -ForegroundColor Yellow
		} else {
			#M365PSProfile not in Profile
			Write-Host "PowerShell Profile already exists. Adding M365PSProfile to it" -ForegroundColor Yellow
			Add-Content -Path $Profile -Value $M365PSProfileContent -Encoding utf8
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

		.PARAMETER Repository
		[string]Repository specifies which PowerShell Repository should be used [Default is PSGallery]

		.PARAMETER FileMode
		[switch]FileMode uses the File System to remove the Modules

		.EXAMPLE
		Uninstall-M365Modules

		.EXAMPLE
		Uninstall-M365Modules -Modules "Az","MSOnline","PnP.PowerShell","Microsoft.Graph" -Scope CurrentUser

		.LINK
		https://github.com/fabrisodotps1/M365PSProfile
	#>

	param (
		[Parameter(Mandatory = $false)][array]$Modules = $global:M365StandardModules,
		[parameter(mandatory = $false)][ValidateSet("CurrentUser", "AllUsers")][string]$Scope = "CurrentUser",
		[parameter(mandatory = $false)][switch]$FileMode = $false
	)

	If ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows))
	{
		#Windows
		$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent()) -ErrorAction SilentlyContinue
		$IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	}

	foreach ($Module in $Modules) {

		#Unload Module
		If ($Module -ne "M365PSProfile")
		{
			Get-Module -Name $Module | Remove-Module -Force -ErrorAction SilentlyContinue
		}

		#Check if Module is Installed
		[Array]$InstalledModules = Get-InstalledPSResource -Name $Module -Scope $Scope -ErrorAction SilentlyContinue | Sort-Object Version -Descending

		if ($InstalledModules) {
			# Module found
			if (($IsAdmin -eq $false) -and ($Scope -eq "AllUsers")) {
				Write-Host "WARNING: PS must be running <As Administrator> to uninstall the Module" -ForegroundColor Red
			} else {
				If ($FileMode -eq $true)
				{
					Write-Host "Using FileMode. Remove all $Module Modules" -ForegroundColor Yellow
					$ModulesPath = Get-M365ModulePath -Scope $Scope
					Get-ChildItem -Path $ModulesPath -Filter "$Module" -Recurse | Remove-Item -Force -Recurse
				} else {
					Write-Host "Uninstall Module: $Module $($InstalledModules.Version.ToString())" -ForegroundColor Yellow
					Uninstall-PSResource -Name $Module -Scope $Scope -SkipDependencyCheck -ErrorAction Stop
				}
			}

			#If AZ also Uninstall all AZ.* Modules
			If ($Module -eq "AZ")
			{
				Write-Host "Uninstall AZ.* Modules" -ForegroundColor Yellow
				#Get-InstalledPSResource -Name "AZ.*" -Scope $Scope -ErrorAction SilentlyContinue | Uninstall-PSResource -Scope $Scope -SkipDependencyCheck
				$InstalledAZModules = Get-InstalledPSResource -Name "AZ.*" -Scope $Scope -ErrorAction SilentlyContinue
				Foreach ($AZModule in $InstalledAZModules)
				{
					#FileMode
					If ($FileMode -eq $true)
					{
						Write-Host "Using FileMode. Remove Module: $($AZModule.Name)" -ForegroundColor Yellow
						$ModulesPath = Get-M365ModulePath -Scope $Scope
						Get-ChildItem -Path $ModulesPath -Filter "AZ.*" -Recurse | Remove-Item -Force -Recurse
					} else {
						Write-Host "Uninstall Module: $($AZModule.Name) $($AZModule.Version.ToString())" -ForegroundColor Yellow
						Uninstall-PSResource -Name $AZModule.Name -Scope $Scope -SkipDependencyCheck -WarningAction SilentlyContinue
					}
				}
			}

			#If Microsoft.Graph also Uninstall all Microsoft.Graph.* Modules
			If ($Module -eq "Microsoft.Graph")
			{
				#FileMode
				If ($FileMode -eq $true)
				{
					Write-Host "Using FileMode. Remove all Microsoft.Graph.* Modules" -ForegroundColor Yellow
					$ModulesPath = Get-M365ModulePath -Scope $Scope
					Get-ChildItem -Path $ModulesPath -Filter "Microsoft.Graph.*" -Recurse | Remove-Item -Force -Recurse
				} else {
					Write-Host "Uninstall Microsoft.Graph.* Modules" -ForegroundColor Yellow
					Get-InstalledPSResource -Name "Microsoft.Graph.*" -Scope $Scope -ErrorAction SilentlyContinue | Where-Object {$_.Name -notmatch "Microsoft.Graph.Beta"} | Uninstall-PSResource -Scope $Scope -SkipDependencyCheck
				}
			}

			#If Microsoft.Graph.Beta also Uninstall all Microsoft.Graph.Beta.* Modules
			If ($Module -eq "Microsoft.Graph.Beta")
			{
				#FileMode
				If ($FileMode -eq $true)
				{
					Write-Host "Using FileMode. Remove all Microsoft.Graph.Beta* Modules" -ForegroundColor Yellow
					$ModulesPath = Get-M365ModulePath -Scope $Scope
					Get-ChildItem -Path $ModulesPath -Filter "Microsoft.Graph.Beta*" -Recurse | Remove-Item -Force -Recurse
				} else {
					Write-Host "Uninstall Microsoft.Graph.Beta.* Modules" -ForegroundColor Yellow
					Get-InstalledPSResource -Name "Microsoft.Graph.Beta*" -Scope $Scope -ErrorAction SilentlyContinue | Uninstall-PSResource -Scope $Scope -SkipDependencyCheck
				}
			}

		} else {
			#Module Notfound
			If ($FileMode -eq $true)
			{
				Write-Host "Using FileMode. Remove all $Module Modules" -ForegroundColor Yellow
				$ModulesPath = Get-M365ModulePath -Scope $Scope
				Get-ChildItem -Path $ModulesPath -Filter "$Module" -Recurse | Remove-Item -Force -Recurse
			}

			#If AZ also Uninstall all AZ.* Modules
			If ($Module -eq "AZ")
			{
				#FileMode
				If ($FileMode -eq $true)
				{
					Write-Host "Using FileMode. Remove all AZ.* Modules" -ForegroundColor Yellow
					$ModulesPath = Get-M365ModulePath -Scope $Scope
					Get-ChildItem -Path $ModulesPath -Filter "AZ.*" -Recurse | Remove-Item -Force -Recurse
				} else {
					Write-Host "NO AZ Root Module. Uninstall AZ.* Modules" -ForegroundColor Yellow
					Get-InstalledPSResource -Name "AZ.*" -Scope $Scope -ErrorAction SilentlyContinue | Uninstall-PSResource -Scope $Scope -SkipDependencyCheck
				}
			}

			#If Microsoft.Graph also Uninstall all Microsoft.Graph.* Modules
			If ($Module -eq "Microsoft.Graph")
			{
				#FileMode
				If ($FileMode -eq $true)
				{
					Write-Host "Using FileMode. Remove all Microsoft.Graph.* Modules" -ForegroundColor Yellow
					$ModulesPath = Get-M365ModulePath -Scope $Scope
					Get-ChildItem -Path $ModulesPath -Filter "Microsoft.Graph.*" -Recurse | Remove-Item -Force -Recurse
				} else {
					Write-Host "NO Microsoft.Graph Root Module. Uninstall Microsoft.Graph.* Modules" -ForegroundColor Yellow
					Get-InstalledPSResource -Name "Microsoft.Graph.*" -Scope $Scope -ErrorAction SilentlyContinue | Where-Object {$_.Name -notmatch "Microsoft.Graph.Beta"} | Uninstall-PSResource -Scope $Scope -SkipDependencyCheck
				}
			}

			#If Microsoft.Graph.Beta also Uninstall all Microsoft.Graph.Beta.* Modules
			If ($Module -eq "Microsoft.Graph.Beta")
			{
				#FileMode
				If ($FileMode -eq $true)
				{
					Write-Host "Using FileMode. Remove all Microsoft.Graph.Beta* Modules" -ForegroundColor Yellow
					$ModulesPath = Get-M365ModulePath -Scope $Scope
					Get-ChildItem -Path $ModulesPath -Filter "Microsoft.Graph.Beta*" -Recurse | Remove-Item -Force -Recurse
				} else {
					Write-Host "NO Microsoft.Graph.Beta Module. Uninstall Microsoft.Graph.Beta.* Modules" -ForegroundColor Yellow
					Get-InstalledPSResource -Name "Microsoft.Graph.Beta*" -Scope $Scope -ErrorAction SilentlyContinue | Uninstall-PSResource -Scope $Scope -SkipDependencyCheck
				}
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
	Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
	Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

	If (Get-Module -Name "Microsoft.Online.SharePoint.PowerShell")
	{
		Disconnect-SPOService -ErrorAction SilentlyContinue
	}

	If (Get-Module -Name "ExchangeOnlineManagement")
	{
		Disconnect-ExchangeOnline -confirm:$false -ErrorAction SilentlyContinue
	}

	If (Get-Module -Name "PnP.PowerShell")
	{
		Disconnect-PnPOnline -ErrorAction SilentlyContinue
	}
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

	[CmdletBinding(SupportsShouldProcess)]
	PARAM (
		[string]$Title = "Windows PowerShell"
	)
	If ($PSCmdlet.ShouldProcess($Title))
	{
		$host.ui.RawUI.WindowTitle = $Title
	}
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

	If ([System.Runtime.InteropServices.RuntimeInformation]::IsOSPlatform([System.Runtime.InteropServices.OSPlatform]::Windows))
	{
		#Windows
		$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent()) -ErrorAction SilentlyContinue
		$IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	}

	Import-Module  Microsoft.PowerShell.PSResourceGet

	if($Repository -eq "PSGallery") {
		$PSGallery = Get-PSResourceRepository -Name PSGallery
		If ($PSGallery.Trusted -eq $false) {
			Write-Host "Warning: PSGallery is not Trusted (Get-PSResourceRepository -Name PSGallery)" -ForegroundColor Yellow
			Write-Host "Set PSGallery to Trusted (Set-PSResourceRepository -Name PSGallery -Trusted:$true)" -ForegroundColor Yellow
			Set-PSResourceRepository -Name PSGallery -Trusted:$true
		}
	}

	#Check if VSCode or PowerShell is running
	[array]$process = Get-Process | Where-Object { $_.ProcessName -eq "powershell" -or $_.ProcessName -eq "pwsh" -or $_.ProcessName -eq "code" }

	If ($process.count -gt 1) {
		Write-Host "PowerShell or Visual Studio Code running? Please close it, Modules in use can't be updated..." -ForegroundColor Yellow
		$process

		#count back from 5 to 1 and start the update
		5..1 | ForEach-Object {
		Write-Host "M365PSProfile Update starts in $_ seconds... (Hit Ctrl+C to cancel)" -ForegroundColor Yellow
		Start-Sleep -Seconds 1
		}
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
		If ($InstalledModules.Count -eq 0)
		{
			#PSResourceGet Module not found
			Write-Host "Install PSResourceGet: Install-Module -Name Microsoft.PowerShell.PSResourceGet -Scope CurrentUser" -ForegroundColor Red
			#Write-Host "Module $Module not found. Try to install..." -ForegroundColor Yellow
			#Install-PSResource -Name $Module -Scope $Scope -TrustRepository -WarningAction SilentlyContinue -Repository $Repository
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
						#Uninstall-PSResource AZ.* -Scope $Scope -SkipDependencyCheck
						Get-InstalledPSResource -Name "AZ.*" -Scope $Scope | Uninstall-PSResource -SkipDependencyCheck
					}

					#If Microsoft.Graph also Uninstall all Microsoft.Graph.* Modules
					If ($Module -eq "Microsoft.Graph")
					{
						Write-Host "Uninstall Microsoft.Graph.* Modules" -ForegroundColor Yellow
						Get-InstalledPSResource -Name "Microsoft.Graph.*" -Scope $Scope | Where-Object {$_.Name -notmatch "Microsoft.Graph.Beta"} | Uninstall-PSResource -SkipDependencyCheck
					}

					#If Microsoft.Graph.Beta also Uninstall all Microsoft.Graph.Beta.* Modules
					If ($Module -eq "Microsoft.Graph.Beta")
					{
						Write-Host "Uninstall Microsoft.Graph.Beta.* Modules" -ForegroundColor Yellow
						Get-InstalledPSResource -Name "Microsoft.Graph.Beta*" -Scope $Scope | Uninstall-PSResource -SkipDependencyCheck
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
	Write-Host "No PowerShell Profile exists. You can add the M365PSProfile Update check with Add-M365PSProfile" -ForegroundColor Yellow
} else {
	$Content = Get-Content -Path $Profile -Encoding utf8
	If ($Content -match "Install-M365Module")
	{
		#Match found
	} else {
		#No Match found
		Write-Host  "You have a PowerShell Profile. You can add the M365PSProfile Update check with Add-M365PSProfile" -ForegroundColor Yellow
	}
}

###############################################################################
# Function Enable-PIM
###############################################################################
#Add -Groups Parameter to get the PIM Groups
#Connect-MgGraph -Scopes PrivilegedEligibilitySchedule.Read.AzureADGroup,RoleAssignmentSchedule.ReadWrite.Directory -NoWelcome
#Connect-MgGraph -Scopes PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup,RoleAssignmentSchedule.ReadWrite.Directory -NoWelcome
#https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=principalid eq '6db8cdd5-8e93-462d-9907-994406c07f60'
#$PIMGroups = Invoke-MgGraphRequest -URI $uri -Method "GET" -ContentType "application/json"
#$PIMGroupDisplayName = (Get-MgGroup -GroupId 932aab78-78d8-428f-a931-7f064a9a491e).DisplayName
#PrivilegedEligibilitySchedule.Read.AzureADGroup	PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup
#$uri = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?$filter=principalid eq '6db8cdd5-8e93-462d-9907-994406c07f60'"
#Invoke-mgGraphRequest -URI $uri -Method "GET" -ContentType "application/json"
Function Enable-PIM
{
	[CmdletBinding()]
	param(
		 [switch]$Groups
	)

	#Check if MgGraph is connected and has the right scope
	$Context = Get-MgContext -ErrorAction SilentlyContinue
	If ($Null -ne $Context)
	{
		If ($Context.Scopes -match "RoleAssignmentSchedule.ReadWrite.Directory")
		{
			#
		} else {
			Write-Host "Disconnect-MgGraph missing Scope: RoleAssignmentSchedule.ReadWrite.Directory"
			Disconnect-MgGraph
		}
	} else {
		Connect-MgGraph -Scopes "RoleAssignmentSchedule.ReadWrite.Directory" -NoWelcome
	}

	#Get Current User
	$context = Get-MgContext
	Write-Host "Your Account: $($Context.Account)" -ForegroundColor Cyan
	$currentUser = (Get-MgUser -UserId $context.Account).Id

    if ($Groups.IsPresent) 
	{
		#Get PIM Groups
		Write-Host "Getting PIM Groups"

		$INT = 0
		$URI = "https://graph.microsoft.com/v1.0/identityGovernance/privilegedAccess/group/eligibilitySchedules?`$filter=principalid eq '" + $currentUser + 	"'"
		$PIMGroups = Invoke-MgGraphRequest -URI $uri -Method "GET" -ContentType "application/json"
		Foreach ($PIMGroup in $PIMGroups.Value)
		{
			$Int = $Int + 1
			Write-Verbose "GroupID: $($PIMGroup.GroupId)"
			$PIMGroupDisplayName = (Get-MgGroup -GroupId $PIMGroup.GroupId).DisplayName
			Write-Host "$($int). $PIMGroupDisplayName"
		}

		# Prompt the user to select a number
		$selectedNumber = Read-Host "Select a number to proceed"

		$SelectedPIMGroup = $PIMGroups.Value[$SelectedNumber -1]
		$Justification = Read-Host "Justification"

		$params = @{
			accessId = "member"
			principalId = $currentUser
			groupId = $SelectedPIMGroup.GroupId
			action = "selfActivate"
			scheduleInfo = @{
				startDateTime = Get-Date
				expiration = @{
					type = "afterDuration"
					Duration = "PT8H"
				}
			}
		justification = $Justification
		}

		#Debug
		#Write-Debug "Params: $params"
		$params

		# Activate the Group
		New-MgIdentityGovernancePrivilegedAccessGroupEligibilityScheduleRequest -BodyParameter $params
    
    } else {
		# Get all available roles
		Write-Host "Getting Eligible Roles"
		$myRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty RoleDefinition -All -Filter "principalId eq '$currentuser'"
		#$DisplayNames =  $myRoles.RoleDefinition.DisplayName

		$Int =0 
		Foreach ($Role in $MyRoles)
		{
			$Int = $Int + 1
			$RoleDisplayName = $Role.RoleDefinition.DisplayName
			Write-Host "$($int). $RoleDisplayName"
		}

		# Prompt the user to select a number
		$selectedNumber = Read-Host "Select a number to proceed"

		$Role = $myRoles[$SelectedNumber -1]
		$Justification = Read-Host "Justification"

		# Setup parameters for activation
		$params = @{
			Action = "selfActivate"
			PrincipalId = $Role.PrincipalId
			RoleDefinitionId = $Role.RoleDefinitionId
			DirectoryScopeId = $Role.DirectoryScopeId
			Justification = $Justification
			ScheduleInfo = @{
				StartDateTime = Get-Date
				Expiration = @{
					Type = "AfterDuration"
					Duration = "PT8H"
				}
			}
		}

		Write-Debug "Params: $params"

		# Activate the role
		New-MgRoleManagementDirectoryRoleAssignmentScheduleRequest -BodyParameter $params
	}
}

###############################################################################
# Function Get-PIMStatus
###############################################################################
Function Get-PIMStatus
{
	[CmdletBinding()]
	param()

	#Check if MgGraph is connected and has the right scope
	$Context = Get-MgContext -ErrorAction SilentlyContinue
	If ($Null -ne $Context)
	{
		If ($Context.Scopes -match "RoleAssignmentSchedule.ReadWrite.Directory")
		{
			#
		} else {
			Write-Host "Disconnect-MgGraph missing Scope: RoleAssignmentSchedule.ReadWrite.Directory"
			Disconnect-MgGraph
		}
	} else {
		Connect-MgGraph -Scopes "RoleAssignmentSchedule.ReadWrite.Directory" -NoWelcome
	}

	#Get Current User
	$context = Get-MgContext
	Write-Host "Your Account: $($Context.Account)" -ForegroundColor Cyan
	$currentUser = (Get-MgUser -UserId $context.Account).Id

	# Get all available roles
	Write-Host "Getting Eligible Roles"
	$myRoles = Get-MgRoleManagementDirectoryRoleEligibilitySchedule -ExpandProperty RoleDefinition -All -Filter "principalId eq '$currentuser'"
	#$DisplayNames =  $myRoles.RoleDefinition.DisplayName

	$Int =0 
	Foreach ($Role in $MyRoles)
	{
		$Int = $Int + 1
		$Filter = "Roledefinitionid eq '" + $Role.Roledefinitionid +"'"
		Write-Verbose "Filter: $Filter"
		$RoleStatus = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter $Filter -Top 1
		$RoleDisplayName = $Role.RoleDefinition.DisplayName
		$AssignmentType = $RoleStatus.AssignmentType
		$StartDateTime = $RoleStatus.StartDateTime
		$EndDateTime = $RoleStatus.EndDateTime
		Write-Host "$($int). $RoleDisplayName Status: $AssignmentType Start: $StartDateTime End: $EndDateTime"
	}
}