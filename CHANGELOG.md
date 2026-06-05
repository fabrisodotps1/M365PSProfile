# Changelog

## V0.9.2

- Added $PSDefaultParameterValues

## V0.9.1

- Changed: Replace Uninstall-PSResource with Uninstall-M365Module -FileMode to cover all aspects of uninstallation
- Changed: Multiple Modules will be uninstalled after 5 seconds delay with User Warning (Resolves: #26)
- Added: -KeepMultipleVersions switch to Install-M365Module to keep multiple Versions of Modules installed (Resolves: #26)
- Added: Microsoft.Entra and Microsoft.Entra.Beta to Standard Modules List
- Enabled Linux/macOS support in Get-M365ModulePath by @spiddeer
- Added -ProfileType parameter to Add-M365PSProfile by @spiddeer
- Changed Logic to detect the Microsoft.PowerShell.PSResourceGet since it is installed in the Program Files folder since PowerShell 7.6.0
- Updated the PSResourceGet dependency to version 1.2.0 to support the new features and improvements in PSResourceGet (Resolves: #25)
- Added Pester Tests for the Module

## V0.9.0

- Added MicrosoftPlaces Module to the Standard Modules (Get-M365StandardModule)
- Bugfix: Uninstall-M365Module -FileMode did not cover all aspects
- Updated to Microsoft.PowerShell.PSResourceGet 1.1.1

## V0.8.0

- Bugfix: Removed unnecessary variable $InstalledModuleVersion in "module is not installed" area
- Added Parameter -FileMode to uninstall Modules using the Filesystem
- Bumped Requirement for Microsoft.PowerShell.PSResourceGet to 1.0.6
- Improved the Disconnect-All Function

## V0.7.0

- Added Version to the Install-M365Module Function
- Added optional Parameter -Repository (default PSGallery) if using multiple Repositorys by @diecknet
- The Function Add-M365PSProfile now adds the needed commands to the $Profile by @diecknet and @BohrenAn
- Changed from "Press any key to continue..." to a Counter from 5 to 1 when other PS Processes are running

## V0.6.0

- Fixed Bug when Modules are not installed
- Updated required Modules to Microsoft.PowerShell.PSResourceGet 1.0.5

## V0.5.0

- Added Info for adding M365PSProfile to the Profile when loading the Module
- Updated required Modules to Microsoft.PowerShell.PSResourceGet 1.0.3
- Added Code to update Microsoft.PowerShell.PSResourceGet to the latest Version
- Added Icon on PowerShell Gallery

## V0.4.0

- Added Code to fix Modules like AZ and Microsoft.Graph

## V0.3.0

Functions:

- Install-M365Module
- Uninstall-M365Module
- Get-M365StandardModule
- Add-M365Profile
- Disconnect-All
- Set-WindowTitle

## v0.2.0-Preview2

- Fixed Add-M365PSProfile
- Updated Readme and added Images