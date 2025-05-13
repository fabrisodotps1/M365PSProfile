## Script Analyzer

```pwsh
Install-PSResource -Name PSScriptAnalyzer -Scope CurrentUser
Invoke-ScriptAnalyzer -Path C:\GIT_WorkingDir\M365PSProfile\M365PSProfile.psm1 -ExcludeRule PSAvoidUsingWriteHost
```

## Testing

```pwsh
Import-Module C:\Git_WorkingDir\M365PSProfile\M365PSProfile.psd1
Get-Command -Module M365PSProfile
Get-M365StandardModule
Uninstall-M365Module
Install-M365Module
Add-M365PSProfile
Disconnect-All
Set-WindowTitle -Title "M365PSProfile"
```

## Pull Request to Main

Create a Pull Request to Main and Approve

## Add GIT Tag

Adding Tags in git

```pwsh
# List Tags
git tag

# Add Tag
git tag -a v0.9.0 -m "Release v0.9.0"

# Push Tag to Repo
git push origin --tags
```

## Create Release

Create Release from GitHub Portal based on Tag

## Deploy to Release Folder

```pwsh
##############################################################################
# Deploy to Release Folder
##############################################################################
Set-Location -Path "C:\GIT_WorkingDir\M365PSProfile\"

$CurrentDirectory = (Get-Location).Path
$Path = (Get-Location).Path + "\Release"

#Delete Folder Release if it exist
If (Test-Path -Path $Path)
{
    Remove-Item -Path $Path -Recurse -confirm:$false
}

#Create Folder Release
New-Item -Path $Path -Type "Directory"

# Copy PoweShell Files to Release Folder
$PSFiles = Get-ChildItem -Path $CurrentDirectory -File | Where-Object {$_.Name -match ".psd1" -or $_.Name -match ".psm1"}
Copy-Item $PSFiles -Destination $Path
```

## Deploy to PowerShell Gallery

Deploy Files from the Release Directory to the PowerShell Gallery

```pwsh
##############################################################################
# Deploy to PowerShell Gallery
##############################################################################
Set-Location -Path "C:\GIT_WorkingDir\M365PSProfile\"
$APIKey = Read-Host "YourSecretApiKey"
$Path = (Get-Location).Path + "\Release\"
Publish-PSResource -Path $Path -ApiKey $APIKey -Repository PSGallery
```
