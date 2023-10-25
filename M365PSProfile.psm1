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
# Update-GraphModules
# Remove old Module instead of only install new Version
##############################################################################
Function Update-GraphModules {
	#Remove Loaded Modules
	Remove-Module Microsoft.Graph*


	##############################################################################
	#This is Fast too

	#Uninstall
	$start = Get-Date
	Get-InstalledPSResource Microsoft.Graph -Scope AllUsers -ErrorAction SilentlyContinue | Uninstall-PSResource -Scope AllUsers -SkipDependencyCheck
	Get-InstalledPSResource Microsoft.Graph* -Scope AllUsers -ErrorAction SilentlyContinue | Uninstall-PSResource -Scope AllUsers -SkipDependencyCheck
	$End = Get-Date
	$TimeSpan = New-TimeSpan -Start $start -End $End
	$TimeSpan

	#Install
	$start = Get-Date
	Install-PSResource Microsoft.Graph -Scope AllUsers
	Install-PSResource Microsoft.Graph.Beta -Scope AllUsers
	$End = Get-Date
	$TimeSpan = New-TimeSpan -Start $start -End $End
	$TimeSpan

	##############################################################################
	#Classic Way
	<#
	#Install
	$Start = Get-Date
	Install-Module Microsoft.Graph
	Install-Module Microsoft.Graph.Beta -AllowClobber
	$End = Get-Date
	$TimeSpan = New-TimeSpan -Start $start -End $End
	$TimeSpan

	#Uninstall
	$start = Get-Date
	Get-Module Microsoft.Graph* -ListAvailable | Uninstall-Module -Force
	Get-Module Microsoft.Graph -ListAvailable | Uninstall-Module -Force
	$End = Get-Date
	$TimeSpan = New-TimeSpan -Start $start -End $End
	$TimeSpan
	#>

	##############################################################################

	##############################################################################
	<#
	$start = Get-Date
	#Get Module and Dependency
	$Graph = Find-PSResource Microsoft.Graph
	$GraphBeta = Find-PSResource Microsoft.Graph.Beta
	$Dependencies = $Graph.Dependencies + $GraphBeta.Dependencies
	#$Dependencies += $GraphBeta.Dependencies

	Foreach ($Module in $Dependencies)
	{
		$ModuleName = $Module.Name
		$MinVersion = $Module.VersionRange.MinVersion
		$MaxVersion = $Module.VersionRange.MaxVersion

		Write-Host "Module: $ModuleName MIN: $MinVersion MAX: $MaxVersion"
	}

	#Parallelisierung mit Jobs 
	$Dependencies | ForEach-Object {
		[string]$ModuleName = $_.Name
		[string]$MinVersion = $_.VersionRange.MinVersion.Version.ToString()
		[string]$MaxVersion = $_.VersionRange.MaxVersion.Version.ToString()
		#Write-Host "Module: $Name"
		#Write-Host "Module: $ModuleName MIN: $MinVersion MAX: $MaxVersion"
	
		$ScriptBlock = {
			param(
				[string]$ModuleName,
				[string]$MinVersion,
				[string]$MaxVersion
			)

			#DEBUG: Checking Parameters
			#Write-Host "Module: $ModuleName MIN: $MinVersion MAX: $MaxVersion"
			
			#Only checks for ALLUSERS Scope at the Moment
			[Array]$InstalledModules = Get-InstalledPSResource -Name $ModuleName -Scope AllUsers -ErrorAction SilentlyContinue
			If ($Null -eq $InstalledModules)
			{
				#No Module installed > Install the Module
				Write-Host "Installing $ModuleName $MinVersion"
				Install-PSResource -Name $ModuleName -Scope AllUsers #-SkipDependencyCheck
			} else {

				#At least Module is installed
				foreach ($InstalledModule in $InstalledModules)
				{
					$InstalledModuleVersion = $InstalledModule.Version.ToString()
					If ($InstalledModuleVersion -lt $MinVersion)
					{
						Write-Host "Uninstalling $ModuleName $InstalledModuleVersion"
						Uninstall-PSResource -Name $ModuleName -Scope AllUsers #-SkipDependencyCheck
						Write-Host "Installing $ModuleName $MinVersion"
						Install-PSResource -Name $ModuleName -Scope AllUsers #-SkipDependencyCheck
					}
				}
			}
		}
	  
		# Show the loop variable here is correct
		#Write-Host "processing $ModuleName..."
	  
		# pass the loop variable across the job-context barrier
		Start-Job $ScriptBlock -ArgumentList $ModuleName,$MinVersion,$MaxVersion | Out-Null
	}

	# Wait for all to complete
	Write-Host "Wait for Jobs to complete"
	While (Get-Job -State "Running") { Start-Sleep 2}

	# Display output from all jobs
	Get-Job | Receive-Job

	# Cleanup
	Remove-Job *

	$End = Get-Date
	$TimeSpan = New-TimeSpan -Start $start -End $End
	$TimeSpan
	##############################################################################
	#>


}

##############################################################################
# Uninstall-M365Modules
# Remove Modules
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

	<# Uninstall still has to be coded #>
}

##############################################################################
# Install-M365Modules
# Remove old Module instead of only install new Version
##############################################################################
Function Install-M365Modules {
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

	<#
	#Check Execution Policy
	Write-Host "Check PowerShell ExecutionPolicy"
	If ((Get-ExecutionPolicy) -eq "Restricted") {
		If ($IsAdmin -eq $false) {
			Write-Host "WARNING: PS must be running <As Administrator> to Change Powershell Execution Policy to <RemoteSigned>" -ForegroundColor Red
		}
		else {
			Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
		}
	}
	
	
	#PSGallery Trusted
	Write-Host "Check PowerShell Gallery"
	If ((Get-PSRepository -Name PSGallery).InstallationPolicy -eq "Untrusted")
	{
		If ($IsAdmin -eq $false)
		{
			Write-Host "WARNING: PS must be running <As Administrator> to set PowerShellGallery as Trusted" -ForegroundColor Red
		} else {
			#Set PSGallery to Trusted
			Write-Host "Set PowerShellGallery to Trusted"
			Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
		}

	}


	#Check PowerShellGet Version
	Write-Host "Check PowerShellGet Module"
	$PSGet = Get-Module PowershellGet -ListAvailable | Sort-Object Version -Descending
	If ($PSGet[0].Version.ToString() -eq "1.0.0.1") {
		If ($IsAdmin -eq $false) {
			Write-Host "WARNING: PS must be running <As Administrator> to set Update PowerShellGet" -ForegroundColor Red
		}
		else {
			Write-Host "Old PowerShellGet Module found... Installing new Version"
			Install-Module PowershellGet -Force -AllowClobber
		}
	}

	#Uninstall Skype for Business Online Powershell Module (msi based)
	Write-Host "Check SkypeForBusiness MSI"
	$Skype4B = Get-Package -Provider Programs -IncludeWindowsInstaller -Name "Skype for Business Online, Windows PowerShell Module" -ErrorAction SilentlyContinue
	If ($Null -ne $Skype4B) {
		If ($IsAdmin -eq $false) {
			Write-Host "You need to uninstall 'Skype for Business Online, Windows PowerShell Module' or run PS as Administrator"
		}
		else {
			Write-Host "Uninstalling 'Skype for Business Online, Windows PowerShell Module'"
			$Skype4B | Uninstall-Package
		}
	}
	
	#>
	

	#Install-Module Microsoft.PowerShell.PSResourceGet -Scope CurrentUser
	Import-Module  Microsoft.PowerShell.PSResourceGet
	$PSGallery = Get-PSResourceRepository -Name PSGallery
	If ($PSGallery.Trusted -eq $false)
	{
		Write-Host "Set PowerShellGallery to Trusted"
		Set-PSResourceRepository -Name PSGallery -Trusted:$true
	}

	#Updating Modules
	If ($IsAdmin -eq $true) {
		#Check for newer Versions of PS Modules
		Write-host "Would you like to Check for newer Versions of PS Modules? (Default is Yes)" -ForegroundColor Yellow 
		$Readhost = Read-Host " ( y / n ) "
		Switch ($ReadHost) { 
			Y { Write-Host "You selected: Updating Modules"; $UpdateCheck = $true }
			N { Write-Host "You selected: Skip Updating Modules" }
			Default { Write-Host "You selected: Updating Modules"; $UpdateCheck = $true } 
		} 

		If ($UpdateCheck -eq $true) {
			#Check if VSCode or PowerShell is running
			[array]$process = Get-Process | Where-Object { $_.ProcessName -eq "powershell" -or $_.ProcessName -eq "pwsh" -or $_.ProcessName -eq "code" }
			#$process = Get-Process -Name code -ErrorAction SilentlyContinue
			If ($process.count -gt 1) {
				Write-Host "PowerShell or Visual Studio Code running? Please close it, otherwise the Modules sometimes can't be updated..." -ForegroundColor Red
				$process
				#Press any key to continue
				Write-Host 'Press any key to continue...';
				$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
			}

			Write-Host "Checking Modules..." 
			Foreach ($Module in $Modules) {
				#Write-Host "Checking Module: $Module"
				If ($Module -eq "AZ" -and $($PSVersionTable.PSVersion.Major) -eq "5") {
					#Sonderfall AZ
					[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending
				}
				else {
					[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
				}

				$Gallerymodule = Find-Module $Module
				$VersionGallerymodule = $Gallerymodule.Version

				If ($Null -ne $InstalledModule) {
					#Module is installed
					$Version = $InstalledModule[0].Version.ToString()
					Write-Host "Checking Module: $Module $Version"

					#Check if Multiple Modules are installed
					If (($InstalledModule.count) -gt 1) {

						Write-host "INFO: Multiple Modules found. Uninstall old Modules? (Default is Yes)" -ForegroundColor Yellow 
						$Readhost = Read-Host " ( y / n ) " 
						Switch ($ReadHost) { 
							Y {
								#Special Cases for AZ and Microsoft.Graph
								switch ($Module) {
									#Custom Update for AZ Modules
									"AZ" { Update-AZModules }
									#Custom Update for Graph Modules
									"Microsoft.Graph" { Update-GraphModules }
									#Default Handling
									default {
										#Uninstall all Modules
										For ($i = 0; $i -lt $InstalledModule.count; $i++) {
											$Version = $InstalledModule[$i].Version.ToString()
											Write-Host "INFO: Uninstall Module $Version"
											Uninstall-Module $Module -Force
										} 
										#Install newest Module
										Write-Host "INFO: Install newest Module $VersionGallerymodule"
										Install-Module $Module
									}
								}

							}
							N { Write-Host "Skip Uninstall old Modules" }
							Default {
								#Special Cases for AZ and Microsoft.Graph
								switch ($Module) {
									#Custom Update for AZ Modules
									"AZ" { Update-AZModules }
									#Custom Update for Graph Modules
									"Microsoft.Graph" { Update-GraphModules }
									#Default Handling
									default {
										#Uninstall all Modules
										For ($i = 0; $i -lt $InstalledModule.count; $i++) {
											$Version = $InstalledModule[$i].Version.ToString()
											Write-Host "INFO: Uninstall Module $Version"
											Uninstall-Module $Module -Force
										} 
										#Install newest Module			  
										Write-Host "INFO: Install newest Module $VersionGallerymodule"
										Install-Module $Module
									}
								}
							} 
						} 

					}
					else {
						#only one Module Version found
						#If Module is newer install it
						If ($Gallerymodule.Version -gt $InstalledModule[0].Version.ToString()) {
							switch ($Module) {
								#Custom Update for AZ Modules
								"AZ" {
									Write-Host "INFO: Uninstall Module $Version"
									Update-AZModules
								}
								#Custom Update for Graph Modules
								"Microsoft.Graph" {
									Write-Host "INFO: Uninstall Module $Version"
									Update-GraphModules
								}
								#Default Handling
								default {
									#Uninstall all Modules
									For ($i = 0; $i -lt $InstalledModule.count; $i++) {
										$Version = $InstalledModule[$i].Version.ToString()
										Write-Host "INFO: Uninstall Module $Version"
										Uninstall-Module $Module -Force
									} 
									#Install newest Module
									Write-Host "INFO: Install newest Module $VersionGallerymodule"
									Install-Module $Module
								}
							}

						} 
					}
				}
				else {
					#Module not found
					Write-Host "Install Module $Module $VersionGallerymodule"
					Install-Module $Module -Scope AllUsers
				}
			}
		}
		else {
			#Update Check = $False
			Write-Host "Checking Modules..."
			Foreach ($Module in $Modules) {
				#Write-Host "Checking Module: $Module"
				#[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
				#[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending
	
				<#
				#Write-Host "Checking Module: $Module"
				If ($Module -eq "AZ" -and $($PSVersionTable.PSVersion.Major) -eq "5") {
					#Sonderfall AZ
					[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending
				}
				else {
					[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
				}
				#>

				[Array]$InstalledModule = Get-InstalledPSResource $Module -Scope $Scope -ErrorAction SilentlyContinue | Sort-Object Version -Descending

				If ($Null -eq $InstalledModule) 
				{
					Write-Host "$Module Module not found. Try to install..."
					If ($IsAdmin -eq $false -and $Scope -eq "AllUsers") 
					{
						Write-Host "WARNING: PS must be running <As Administrator> to install the Module" -ForegroundColor Red
					} else {
						#Install-Module $Module -Confirm:$false
						Install-PSResource $Module -Scope $Scope
					}
				}
				else {
					#Write-Host "Loading Module: MicrosoftGraph"
					#Import-Module MicrosoftGraph
					$Version = $InstalledModule[0].Version.ToString()
					Write-Host "Checking Module: $Module $Version"
					#If ($InstalledModule[0].InstalledLocation -match "OneDrive")
					If ($InstalledModule[0].InstalledLocation -match "OneDrive") {
						Write-Host "Module might be installed in OneDrive Folder - this can lead to Problems" -ForegroundColor Yellow
					}
				}
			}
		}
	}
 else {
		#IsAdmin = $False
		Write-Host "Checking Modules... (NoAdmin)"
		Foreach ($Module in $Modules) {
			#Write-Host "Checking Module: $Module"
			#[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
			#[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending

			<#
			#Write-Host "Checking Module: $Module"		
			If ($Module -eq "AZ" -and $($PSVersionTable.PSVersion.Major) -eq "5") {
				#Sonderfall AZ
				[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending
			}
			else {
				[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
			}
			#>

			[Array]$InstalledModules = Get-InstalledPSResource $Module -Scope $Scope -ErrorAction SilentlyContinue | Sort-Object Version -Descending

			If ($Null -eq $InstalledModule) 
			{
				Write-Host "$Module Module not found. Try to install..."
				If ($IsAdmin -eq $false -and $Scope -eq "AllUsers") 
				{
					Write-Host "WARNING: PS must be running <As Administrator> to install the Module" -ForegroundColor Red
				} else {
					#Install-Module $Module -Confirm:$false
					$Version = (Find-PSResource -Name $Module).Version.ToString()
					Write-Host "INFO: Install Module $Module $Version"
					Install-PSResource $Module -Scope $Scope
				}
			} else {
				$Version = $InstalledModule[0].Version.ToString()
				Write-Host "Checking Module: $Module $Version"
				If ($InstalledModule[0].InstalledLocation -match "OneDrive") 
				{
					Write-Host "Module is be installed in OneDrive Folder - this can lead to Problems" -ForegroundColor Yellow
				}

				foreach ($InstalledModule in $InstalledModules)
				{
					#Uninstall all Modules
					$Version = $InstalledModule.Version.ToString()
					Write-Host "INFO: Uninstall Module $Module $Version"
					Uninstall-PSResource $Module -Scope $Scope -Force
				}
				#Install newest Module
				$Version = (Find-PSResource -Name $Module).Version.ToString()
				Write-Host "INFO: Install Module $Module $Version"
				Install-PSResource $Module -Scope $Scope
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
	}
 catch {
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
		[parameter(mandatory=$false)][bool]$AsciiArt = $true
		)


	Write-Host "Starting M365PSProfile..."
	If ($AsciiArt -eq $true)
	{
		#Show AsciArt
		Invoke-AsciiArt
	}

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


		#Update Logic / How often the Install and Update Check will be invoked
		#Every x Days?
		#Registry Key > Will work for Windows, what about Linux?
		#File

		#Install-Modules
		Install-M365Modules -Modules $Modules -Scope $Scope 
	}
}

