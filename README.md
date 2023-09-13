# M365PSProfile

If you're an Microsoft 365 Administrator, you need to have several PowerShell Modules for Managing an M365 Environement..

We've created a flexible Module, that simplifies the Installation and Updatemanagement of these Modules.

## Goals
- Simple One-Liner in the [PowerShell Profile](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_profiles?view=powershell-7.3)
- No Admin Rights required
- Support for PowerShell 5 and 7 (Install in CurrentUser)
- Parameter for Modules that should be installed and updated
- Fast and configurable

```pwsh
Update-ModuleCustom -Modules @("MSOnline", "AzureADPreview", "ExchangeOnlineManagement", "Icewolf.EXO.SpamAnalyze", "MicrosoftTeams", "Microsoft.Online.SharePoint.PowerShell", "PnP.PowerShell" , "ORCA", "O365CentralizedAddInDeployment", "MSCommerce", "WhiteboardAdmin", "Microsoft.Graph", "Microsoft.Graph.Beta", "MSAL.PS", "MSIdentityTools" )
```

> Note:  Command might still Change

## Modules 

| Module | Description |
| --- | --- |
| MSOnline | AzureAD / Entra ID Module > Depreciated 30.03.2024 |
| AzureADPreview | AzureAD / Entra ID Module > Depreciated 30.03.2024 |
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
| MSAL.PS | Microsoft Authentication Library | 
| MSIdentityTools | Additional Functions for Identity |

## Contribution
How can you contribute?

- Create an Issue
- Fork the Repo and create a Pull Request

## Maintainer
- Fabrice Reiser [@fabrisodotps1](https://twitter.com/fabrisodotps1)
- Andres Bohren [@andresbohren](https://twitter.com/andresbohren)
