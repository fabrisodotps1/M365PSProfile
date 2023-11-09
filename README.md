# M365PSProfile

If you're an Microsoft 365 Administrator, you need to have several PowerShell Modules for Managing an M365 Environement..

We've created a flexible Module, that simplifies the Installation and Update of these Modules.

## ToDo
### Andres
- Anpassen der Moduleliste
- Testing PS5/PS7 mit und ohne Admin Rights
- Readme überarbeiten
  - Erklärung Profile und was rein muss

### Fabrice
- Uninstall Function coden
- Vorschlag für Update Mechanismus überlegen


## Goals
- Simple One-Liner in the [PowerShell Profile](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_profiles?view=powershell-7.3)
- No Admin Rights required
- Fast and configurable
- Support for PowerShell 5 and 7 (Install in CurrentUser Scope)
- Parameter for Modules that should be installed and updated
- Use the Microsoft.PowerShell.PSResourceGet

## Installation

You need to install the Module

```pwsh
#PowerShellGet
Install-Module -Name M365PSProfile

#Microsoft.PowerShell.PSResourceGet
Install-PSResource -Name M365PSProfile
```

## Usage

### PowerShell Profiles
- Diffrent Scopes [CurrentUser / AllUsers]
- CurrentUser / CurrentHost
- Diffrent Locations in PowerShell 5/7


Open the PowerShell Profile add the following Line

```
#Current user, Current Host
#Windows - $HOME\Documents\PowerShell\Microsoft.PowerShell_profile.ps1
#Linux - ~/.config/powershell/Microsoft.PowerShell_profile.ps1
#macOS - ~/.config/powershell/Microsoft.PowerShell_profile.ps1
```

### What you have to put into your Profile

> Note:  Command might still Change

```pwsh
Import-Module -Name M365PSProfile
#Install or updates the default Modules (what we think every M365 Admin needs) in the CurrentUser Scope
Install-M365Module

#Install or Updates the Modules in the Array
Install-M365Module -Modules @("MSOnline", "AzureADPreview", "ExchangeOnlineManagement", "Icewolf.EXO.SpamAnalyze", "MicrosoftTeams", "Microsoft.Online.SharePoint.PowerShell", "PnP.PowerShell" , "ORCA", "O365CentralizedAddInDeployment", "MSCommerce", "WhiteboardAdmin", "Microsoft.Graph", "Microsoft.Graph.Beta", "MSAL.PS", "MSIdentityTools" )
```

### Parameters
```pwsh
-Modules @(ArrayOfModulenames)
-Scope [Default:CurrentUser/AllUsers]
-AsciiArt [Default:true/false]
-UpdateCheckDays [Default:7]
-RunInVSCode [Default:false/true]
```

### Modules 

| Module | Description |
| --- | --- |
| #MSOnline | AzureAD / Entra ID Module > Depreciated 30.03.2024 |
| #AzureADPreview | AzureAD / Entra ID Module > Depreciated 30.03.2024 |
| ExchangeOnlineManagement | Exchange Online |
| Icewolf.EXO.SpamAnalyze | Exchange Online Message Tracking / SpamAnalyze | 
| MicrosoftTeams | Microsoft Teams |
| Microsoft.Online.SharePoint.PowerShell | Microsoft Sharepoint | 
| PnP.PowerShell | SharePoint / Microsoft Teams |
| ORCA | Defender for Office 365 Recommended Configuration Analyzer |
| O365CentralizedAddInDeployment | Deploy Office Add-Ins | 
| MSCommerce | Manage M365 SelfServicePurchase | 
| WhiteboardAdmin | Manage Whiteboards |
| Microsoft.Graph | Microsoft.Graph Modules https://graph.microsoft.com/v1.0 | 
| Microsoft.Graph.Beta | Microsoft.Graph Modules https://graph.microsoft.com/beta |
| #MSAL.PS | Microsoft Authentication Library (Depreciated)| 
| PSMSALNet| PowerShell 7.2 MSAL.NET wrapper| 
| MSIdentityTools | Additional Functions for Identity |

## Contribution
How can you contribute?

- Create an Issue
- Fork the Repo and create a Pull Request

## Maintainer
- Fabrice Reiser [@fabrisodotps1](https://twitter.com/fabrisodotps1)
- Andres Bohren [@andresbohren](https://twitter.com/andresbohren)
