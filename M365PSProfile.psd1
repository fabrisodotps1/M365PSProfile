﻿# Module manifest for module 'M365PSProfile'
# Generated by: Andres Bohren
# Generated on: 9/13/2023

@{

# Script module or binary module file associated with this manifest.
RootModule = 'M365PSProfile.psm1'

# Version number of this module.
ModuleVersion = '0.7.0'

# Supported PSEditions
CompatiblePSEditions = @('Core', 'Desktop')

# ID used to uniquely identify this module
GUID = '36bed064-7e84-4df2-b6d5-a3346c217051'

# Author of this module
Author = 'Andres Bohren, Fabrice Reiser'

# Company or vendor of this module
CompanyName = ''

# Copyright statement for this module
Copyright = '(c) 2024 Andres Bohren & Fabrice Reiser'

# Description of the functionality provided by this module
Description = 'This PowerShell Module helps M365 Administrators to keep the needed PowerShell Modules up to date'

# Minimum version of the Windows PowerShell engine required by this module
# PowerShellVersion = ''
PowerShellVersion = '5.1'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()
RequiredModules = @(@{ModuleName = 'Microsoft.PowerShell.PSResourceGet'; GUID = 'e4e0bda1-0703-44a5-b70d-8fe704cd0643'; ModuleVersion = '1.0.5'; })

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @('Install-M365Module', 'Uninstall-M365Module', 'Get-M365StandardModule', 'Add-M365PSProfile', 'Disconnect-All', 'Set-WindowTitle')

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

	PSData = @{

		# Tags applied to this module. These help with module discovery in online galleries.
		Tags = @('Office365', 'Microsoft365', 'O365', 'M365', 'Admin' )

		# A URL to the license for this module.
		# LicenseUri = ''

		# A URL to the main website for this project.
		ProjectUri = 'https://github.com/fabrisodotps1/M365PSProfile'

		# A URL to an icon representing this module.
		IconUri = 'https://raw.githubusercontent.com/fabrisodotps1/M365PSProfile/develop/M365PSProfile.png'

 		# Set to a prerelease string value if the release should be a prerelease.
 		#Prerelease = 'Preview2'

		# ReleaseNotes of this module
		ReleaseNotes = '
---------------------------------------------------------------------------------------------
Whats new in this release:
V0.7.0
- Added Version to the Install-M365Module Function
- Added optional Parameter -Repository (default PSGallery) if using multiple Repositorys by @diecknet
- The Function Add-M365PSProfile now adds the needed commands to the $Profile by @diecknet and @bohrenan
- Changed from "Press any key to continue..." to a Counter from 5 to 1 when other PS Processes are running
---------------------------------------------------------------------------------------------
'
	} # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}

