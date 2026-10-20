# Pester tests for M365PSProfile
# Run from repo root: Invoke-Pester

$ErrorActionPreference = 'Stop'

$script:ExpectedExports = @(
    'Install-M365Module',
    'Uninstall-M365Module',
    'Get-M365StandardModule',
    'Add-M365PSProfile',
    'Disconnect-All',
    'Set-WindowTitle'
)

BeforeAll {
    $repoRoot = Resolve-Path (Join-Path $PSScriptRoot '..')

    $script:ManifestPath = Join-Path $repoRoot 'M365PSProfile.psd1'
    $script:ModulePath = Join-Path $repoRoot 'M365PSProfile.psm1'
}

Describe 'M365PSProfile' {
    It 'Manifest file exists' {
        $ManifestPath | Should -Exist
    }

    It 'Module file exists' {
        $ModulePath | Should -Exist
    }

    It 'Module script parses without syntax errors' {
        $tokens = $null
        $errors = $null
        [void][System.Management.Automation.Language.Parser]::ParseFile($ModulePath, [ref]$tokens, [ref]$errors)
        $errors.Count | Should -Be 0
    }

    It 'Manifest exports the expected functions' {
        $data = Import-PowerShellDataFile -Path $ManifestPath
        $data.FunctionsToExport | Should -Not -BeNullOrEmpty

        foreach ($name in $script:ExpectedExports) {
            $data.FunctionsToExport | Should -Contain $name
        }
    }

    Context 'Import and command surface' {
        BeforeAll {
            $SkipImportBecause = $null

            $requiredModuleName = 'Microsoft.PowerShell.PSResourceGet'
            $requiredModuleMinVersion = [version]'1.2.0'

            $available = Get-Module -Name $requiredModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
            if (-not $available -or $available.Version -lt $requiredModuleMinVersion) {
                $SkipImportBecause = "Required dependency '$requiredModuleName' >= $requiredModuleMinVersion is not available on this machine."
            }
        }

        AfterAll {
            Remove-Module -Name 'M365PSProfile' -Force -ErrorAction SilentlyContinue
        }

        It 'Can import the local module manifest' {
            if ($SkipImportBecause) {
                Skip -Because $SkipImportBecause
                return
            }

            Import-Module -Name $ManifestPath -Force -ErrorAction Stop
            (Get-Module -Name 'M365PSProfile') | Should -Not -BeNullOrEmpty
        }

        It 'Exports expected commands after import' {
            if ($SkipImportBecause) {
                Skip -Because $SkipImportBecause
                return
            }

            foreach ($name in $script:ExpectedExports) {
                Get-Command -Name $name -Module 'M365PSProfile' -ErrorAction Stop | Should -Not -BeNullOrEmpty
            }
        }

        It 'Get-M365StandardModule returns a non-empty array' {
            if ($SkipImportBecause) {
                Skip -Because $SkipImportBecause
                return
            }

            $modules = @(Get-M365StandardModule)
            $modules | Should -Not -BeNullOrEmpty
            $modules.Count | Should -BeGreaterThan 0
            $modules | Should -Contain 'Microsoft.Graph'
            $modules | Should -Contain 'M365PSProfile'
        }

        It 'Each exported function has help available' {
            if ($SkipImportBecause) {
                Skip -Because $SkipImportBecause
                return
            }

            foreach ($name in $script:ExpectedExports) {
                $help = Get-Help -Name $name -ErrorAction Stop
                $help.Synopsis | Should -Not -BeNullOrEmpty
            }
        }
    }
}

