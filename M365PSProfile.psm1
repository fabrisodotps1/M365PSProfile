###############################################################################
# M365 PS Profle 
###############################################################################
#Zielvorstellung
#•	Im Profil nur noch das Modul laden
#•	Parameter für die zu installierenden Module
#•	Soll PowerShell 5 / 7 Unterstützen (Allenfalls https://github.com/JustinGrote/ModuleFast für PowerShell 7)
#•	https://blog.icewolf.ch/archive/2023/08/12/PowerShellGet-will-be-renamed-to-Microsoft-PowerShell-PSResourceGet/
#•	Soll keine Admin Rechte brauchten > Sondern im CurrentUser Kontext funktionieren
#•	Aktuell dauert AZ Installation ewig. Evtl müssen wir da die installierten Module mit den Dependencys prüfen. Allenfalls kann man die Installation auch paralellisieren
###############################################################################

##############################################################################
# Update-AZModules
# Remove old Module instead of only install new Version
##############################################################################
Function Update-AZModules {
	#Remove loaded az.* Modules
	Remove-Module az.*

	#Removing all AZ Modules and installing them takes a long time
	#Get-Module AZ.* -ListAvailable | Uninstall-Module -Force

	$NewAZModule = Find-Module -Name AZ
	
	#Iterate through Modules and uninstall
	$Modules = Get-Module AZ.* -ListAvailable | Where-Object {$_.Name -ne "Az.Accounts"} | Select-Object Name -Unique
	Foreach ($Module in $Modules)
	{
		$ModuleName = $Module.Name
		$Versions = Get-Module $ModuleName -ListAvailable
		#$Versions = Get-InstalledModule $ModuleName -AllVersions
		Foreach ($Version in $Versions)
		{
			$ModuleVersion = $Version.Version

			#Check Version of Submodule if it's older than the dependency of the new AZ Module > Uninstall
			$NewDependency = $NewAZModule.Dependencies | Where-Object {$_.Name -eq $ModuleName}
			If ($NewDependency.Version -ne $ModuleVersion)
			{
				Write-Host "Uninstall-Module $ModuleName $ModuleVersion"
				Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion
			}
		}
	}
	#Uninstall Az.Accounts
	$ModuleName = "Az.Accounts"
	$Versions = Get-Module $ModuleName -ListAvailable
	#$Versions = Get-InstalledModule $ModuleName -AllVersions
	Foreach ($Version in $Versions)
	{
		$ModuleVersion = $Version.Version

		#Check Version of Submodule if it's older than the dependency of the new AZ Module > Uninstall
		$NewDependency = $NewAZModule.Dependencies | Where-Object {$_.Name -eq $ModuleName}
		If ($NewDependency.Version -ne $ModuleVersion)
		{
			Write-Host "Uninstall-Module $ModuleName $ModuleVersion"
			Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion
		}
	}
	<#
	#Uninstall Az
	$ModuleName = "Az"
	$Versions = Get-Module $ModuleName -ListAvailable
	#$Versions = Get-InstalledModule $ModuleName -AllVersions
	Foreach ($Version in $Versions)
	{
		$ModuleVersion = $Version.Version
		Write-Host "Uninstall-Module $ModuleName $ModuleVersion"
		Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion -Force
	}
	#>
	
	#Uninstall AZ Module
	Get-Module AZ -ListAvailable | Uninstall-Module -Force
	Get-InstalledModule AZ | Uninstall-Module -Force

	#Install newest Module
	Write-Host "Install newest AZ Module"
	Install-Module AZ
	Write-Host "AZ Module Cleanup finished"
}

##############################################################################
# Update-GraphModules
# Remove old Module instead of only install new Version
##############################################################################
Function Update-GraphModules {
	#Remove loaded Microsoft.Graph* Modules
	Remove-Module Microsoft.Graph*

	Get-Module Microsoft.Graph* -ListAvailable | Uninstall-Module -Force
	Get-Module Microsoft.Graph -ListAvailable | Uninstall-Module -Force

	#Get Microsoft.Graph Modules except Microsoft.Graph.Authentication because the other Modules have dependencys
	<#
	$Modules = Get-Module Microsoft.Graph* -ListAvailable | Where-Object {$_.Name -ne "Microsoft.Graph.Authentication"} | Select-Object Name -Unique
	Foreach ($Module in $Modules)
	{
		$ModuleName = $Module.Name
		#Get Installed Versions of that specific Module
		$Versions = Get-Module $ModuleName -ListAvailable
		Foreach ($Version in $Versions)
		{
			#Uninstall Module
			$ModuleVersion = $Version.Version
			Write-Host "Uninstall-Module $ModuleName $ModuleVersion"
			Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion
		}
	}
	#Uninstall Microsoft.Graph.Authentication
	$ModuleName = "Microsoft.Graph.Authentication"
	$Versions = Get-Module $ModuleName -ListAvailable
	#$Versions = Get-InstalledModule $ModuleName -AllVersions
	Foreach ($Version in $Versions)
	{
		$ModuleVersion = $Version.Version
		Write-Host "Uninstall-Module $ModuleName $ModuleVersion"
		Uninstall-Module $ModuleName -RequiredVersion $ModuleVersion -Force
	}
	#>

	#Finally install the newest Version
	Write-Host "Install newest Microsoft.Graph Module"
	Install-Module Microsoft.Graph
	Install-Module Microsoft.Graph.Beta -AllowClobber
	Write-Host "Cleanup finished"
}

##############################################################################
# Update-ModuleCustom
# Remove old Module instead of only install new Version
##############################################################################
Function Update-ModuleCustom {

	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
	$IsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	#Check Execution Policy
	Write-Host "Check PowerShell ExecutionPolicy"
	If ((Get-ExecutionPolicy) -eq "Restricted")
	{
		If ($IsAdmin -eq $false)
		{
			Write-Host "WARNING: PS must be running <As Administrator> to Change Powershell Execution Policy to <RemoteSigned>" -ForegroundColor Red
		} else {
			Set-ExecutionPolicy -ExecutionPolicy RemoteSigned
		}
	}
	
	<#
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
	#>

	#Check PowerShellGet Version
	Write-Host "Check PowerShellGet Module"
	$PSGet = Get-Module PowershellGet -ListAvailable | Sort-Object Version -Descending
	If ($PSGet[0].Version.ToString() -eq "1.0.0.1")
	{
		If ($IsAdmin -eq $false)
		{
			Write-Host "WARNING: PS must be running <As Administrator> to set Update PowerShellGet" -ForegroundColor Red
		} else {
			Write-Host "Old PowerShellGet Module found... Installing new Version"
			Install-Module PowershellGet -Force -AllowClobber
		}
	}

	#Uninstall Skype for Business Online Powershell Module (msi based)
	Write-Host "Check SkypeForBusiness MSI"
	$Skype4B = Get-Package -Provider Programs -IncludeWindowsInstaller -Name "Skype for Business Online, Windows PowerShell Module" -ErrorAction SilentlyContinue
	If ($Null -ne $Skype4B)
	{
		If ($IsAdmin -eq $false)
		{
			Write-Host "You need to uninstall 'Skype for Business Online, Windows PowerShell Module' or run PS as Administrator"
		} else {
			Write-Host "Uninstalling 'Skype for Business Online, Windows PowerShell Module'"
			$Skype4B | Uninstall-Package
		}
	}
	

	#Updating Modules
	If ($IsAdmin -eq $true)
	{
		#Check for newer Versions of PS Modules
		Write-host "Would you like to Check for newer Versions of PS Modules? (Default is Yes)" -ForegroundColor Yellow 
		$Readhost = Read-Host " ( y / n ) " 
		Switch ($ReadHost) 
		{ 
			Y {Write-Host "You selected: Updating Modules"; $UpdateCheck = $true}
			N {Write-Host "You selected: Skip Updating Modules"}
			Default {Write-Host "You selected: Updating Modules"; $UpdateCheck = $true} 
		} 

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
				#Write-Host "Checking Module: $Module"
				If ($Module -eq "AZ" -and $($PSVersionTable.PSVersion.Major) -eq "5")
				{
					#Sonderfall AZ
					[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending
				} else {
					[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
				}

				$Gallerymodule = Find-Module $Module
				$VersionGallerymodule = $Gallerymodule.Version

				If ($Null -ne $InstalledModule)
				{
					#Module is installed
					$Version = $InstalledModule[0].Version.ToString()
					Write-Host "Checking Module: $Module $Version"

					#Check if Multiple Modules are installed
					If (($InstalledModule.count) -gt 1)
					{

						Write-host "INFO: Multiple Modules found. Uninstall old Modules? (Default is Yes)" -ForegroundColor Yellow 
						$Readhost = Read-Host " ( y / n ) " 
						Switch ($ReadHost) 
						{ 
							Y {
								#Special Cases for AZ and Microsoft.Graph
								switch ($Module)
								{
									#Custom Update for AZ Modules
									"AZ" {Update-AZModules}
									#Custom Update for Graph Modules
									"Microsoft.Graph" {Update-GraphModules}
									#Default Handling
									default {
										#Uninstall all Modules
										For ($i=0; $i -lt $InstalledModule.count; $i++) 
										{
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
							N {Write-Host "Skip Uninstall old Modules"}
							Default {
								#Special Cases for AZ and Microsoft.Graph
								switch ($Module)
								{
									#Custom Update for AZ Modules
									"AZ" {Update-AZModules}
									#Custom Update for Graph Modules
									"Microsoft.Graph" {Update-GraphModules}
									#Default Handling
									default {
										#Uninstall all Modules
										For ($i=0; $i -lt $InstalledModule.count; $i++) 
										{
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

					} else {
						#only one Module Version found
						#If Module is newer install it
						If ($Gallerymodule.Version -gt $InstalledModule[0].Version.ToString())
						{
							switch ($Module)
							{
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
									For ($i=0; $i -lt $InstalledModule.count; $i++) 
									{
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
				} else {
					#Module not found
					Write-Host "Install Module $Module $VersionGallerymodule"
					Install-Module $Module -Scope AllUsers
				}
			}
		} else {
			#Update Check = $False
			Write-Host "Checking Modules..."
			Foreach ($Module in $Modules)
			{
				#Write-Host "Checking Module: $Module"
				#[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
				#[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending
	
				#Write-Host "Checking Module: $Module"
				If ($Module -eq "AZ" -and $($PSVersionTable.PSVersion.Major) -eq "5")
				{
					#Sonderfall AZ
					[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending
				} else {
					[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
				}
	
				If ($Null -eq $InstalledModule)
				{
					Write-Host "$Module Module not found. Try to install..."
					If ($IsAdmin -eq $false)
					{
						Write-Host "WARNING: PS must be running <As Administrator> to install the Module" -ForegroundColor Red
					} else {
						#Install-Module $Module -Confirm:$false
					}
				} else {
					#Write-Host "Loading Module: MicrosoftGraph"
					#Import-Module MicrosoftGraph
					$Version = $InstalledModule[0].Version.ToString()
					Write-Host "Checking Module: $Module $Version"
					#If ($InstalledModule[0].InstalledLocation -match "OneDrive")
					If ($InstalledModule[0].Path -match "OneDrive")
					{
						Write-Host "Module might be installed in OneDrive Folder - this can lead to Problems" -ForegroundColor Yellow
					}
				}
			}
		}
	} else {
		#IsAdmin = $False
		Write-Host "Checking Modules... (NoAdmin)"
		Foreach ($Module in $Modules)
		{
			#Write-Host "Checking Module: $Module"
			#[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
			#[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending

			#Write-Host "Checking Module: $Module"		
			If ($Module -eq "AZ" -and $($PSVersionTable.PSVersion.Major) -eq "5")
			{
				#Sonderfall AZ
				[Array]$InstalledModule = Get-InstalledModule -Name $Module -AllVersions -ErrorAction SilentlyContinue | Sort-Object Version -Descending
			} else {
				[Array]$InstalledModule = Get-Module $Module -ListAvailable | Sort-Object Version -Descending
			}

			If ($Null -eq $InstalledModule)
			{
				Write-Host "$Module Module not found. Try to install..."
				If ($IsAdmin -eq $false)
				{
					Write-Host "WARNING: PS must be running <As Administrator> to install the Module" -ForegroundColor Red
				} else {
					#Install-Module $Module -Confirm:$false
				}
			} else {
				#Write-Host "Loading Module: MicrosoftGraph"
				#Import-Module MicrosoftGraph
				$Version = $InstalledModule[0].Version.ToString()
				Write-Host "Checking Module: $Module $Version"
				#If ($InstalledModule[0].InstalledLocation -match "OneDrive")
				If ($InstalledModule[0].Path -match "OneDrive")
				{
					Write-Host "Module might be installed in OneDrive Folder - this can lead to Problems" -ForegroundColor Yellow
				}
			}
		}

	}


}

##############################################################################
# Remove existing PS Connections
##############################################################################
Function dcon-all {
	Get-PSSession | Remove-PSSession
	Try {		
		Disconnect-AzureAD -ErrorAction SilentlyContinue
		Disconnect-MsolService -ErrorAction SilentlyContinue
		Disconnect-SPOService -ErrorAction SilentlyContinue
		Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue
		Disconnect-ExchangeOnline -confirm:$false -ErrorAction SilentlyContinue
		Disconnect-MgGraph -ErrorAction SilentlyContinue
	} catch {
	}
}


#############################################################################
# Set-WindowTitle Function
#############################################################################
Function Set-WindowTitle
{
	PARAM (
			[string]$Title="Windows PowerShell"
	 )
	$host.ui.RawUI.WindowTitle = $Title
}

##############################################################################
# Main Program
##############################################################################
$pshost = get-host
$pswindow = $pshost.ui.rawui
$LanguageMode = $ExecutionContext.SessionState.LanguageMode
If ($LanguageMode -eq "Fulllanguage")
{
	if ($pswindow.WindowSize.Width -lt 220)
	{
		if ($env:WT_SESSION) 
		{
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
		} else {
			if ($env:TERM_PROGRAM -eq 'vscode') 
			{
				#If vscode do nothing
			} else {
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

		#Call Function to Load/Install Modules
		#$Modules = @("MSOnline", "AzureADPreview", "ExchangeOnlineManagement", "MicrosoftTeams", "Microsoft.Online.SharePoint.PowerShell","SharePointPnPPowerShellOnline" , "ORCA", "O365CentralizedAddInDeployment", "MSCommerce", "WhiteboardAdmin", "Microsoft.Graph", "MSAL.PS" )
		$Modules = @("AZ","MSOnline", "AzureADPreview", "ExchangeOnlineManagement", "Icewolf.EXO.SpamAnalyze", "MicrosoftTeams", "Microsoft.Online.SharePoint.PowerShell","PnP.PowerShell" , "ORCA", "O365CentralizedAddInDeployment", "MSCommerce", "WhiteboardAdmin", "Microsoft.Graph","Microsoft.Graph.Beta", "MSAL.PS", "MSIdentityTools" )
		#Install-Modules
		Update-ModuleCustom
}

