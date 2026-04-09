#Requires -Version 5.1
<#
.SYNOPSIS
    BuildMaster - Machine Rebuild Manager
.DESCRIPTION
    Manages the end-to-end process of staging, rebooting, and monitoring
    machine rebuilds via the BuildMaster and VirtualWorks APIs.

    DEPLOYMENT LAYOUT (environment determined by PARENT folder name):
        \\FileServer\ITTools\BuildMaster\dev\BuildMaster.ps1
        \\FileServer\ITTools\BuildMaster\qa\BuildMaster.ps1
        \\FileServer\ITTools\BuildMaster\uat\BuildMaster.ps1
        \\FileServer\ITTools\BuildMaster\prod\BuildMaster.ps1

    Each env folder contains its own copy of BuildMaster.ps1 + Modules\.
    The parent folder name is auto-detected as the environment stamp.

.NOTES
    PowerShell 5.1 | WPF GUI | No external module dependencies
    API Endpoints:
        BuildMaster : http://bm.zz.com/new-api/v1
        VirtualWorks: http://vw.zz.com/vwapi/Desktop
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
#  ENVIRONMENT DETECTION  (parent folder name  dev / qa / uat / prod)
# -----------------------------------------------------------------------------
$script:BMEnvironment = Split-Path -Leaf $PSScriptRoot
$validEnvs = @('dev', 'qa', 'uat', 'prod')

if ($script:BMEnvironment -notin $validEnvs) {
    Write-Warning (("Parent folder '{0}' is not a recognised environment " +
                   "(dev/qa/uat/prod). Treating as 'dev'.") -f $script:BMEnvironment)
    $script:BMEnvironment = 'dev'
}

# -----------------------------------------------------------------------------
#  PATHS
# -----------------------------------------------------------------------------
$script:ModulesPath = Join-Path $PSScriptRoot 'Modules'
$script:LogPath     = Join-Path $PSScriptRoot 'Logs'

if (-not (Test-Path $script:LogPath)) {
    New-Item -ItemType Directory -Path $script:LogPath -Force | Out-Null
}

# -----------------------------------------------------------------------------
#  WPF ASSEMBLIES  (must load before modules that reference WPF types)
# -----------------------------------------------------------------------------
Add-Type -AssemblyName PresentationFramework  -ErrorAction Stop
Add-Type -AssemblyName PresentationCore       -ErrorAction Stop
Add-Type -AssemblyName WindowsBase            -ErrorAction Stop
Add-Type -AssemblyName System.Windows.Forms   -ErrorAction Stop   # MessageBox helper

# -----------------------------------------------------------------------------
#  IMPORT MODULES  (order matters - Auth/API before Engine, all before GUI)
# -----------------------------------------------------------------------------
$moduleOrder = @('BM-Auth', 'BM-API', 'BM-Engine', 'BM-GUI')

foreach ($mod in $moduleOrder) {
    $modPath = Join-Path $script:ModulesPath "$mod.psm1"
    if (-not (Test-Path $modPath)) {
        throw "Required module not found: $modPath"
    }
    Import-Module $modPath -Force -Global -DisableNameChecking
}

# -----------------------------------------------------------------------------
#  SESSION CONTEXT DETECTION
#  (must run after WPF assemblies are loaded so MessageBox is available,
#   and after modules are imported so auth functions exist)
# -----------------------------------------------------------------------------
Set-BMLogPath -Path $script:LogPath

try {
    Initialize-BMSessionContext -Environment $script:BMEnvironment | Out-Null
}
catch {
    # Context detection failure is non-fatal - GUI will surface the issue
    Write-Warning ("Session context detection error: {0}" -f $_.Exception.Message)
}

# -----------------------------------------------------------------------------
#  LAUNCH GUI
# -----------------------------------------------------------------------------
try {
    Start-BMGui -Environment $script:BMEnvironment -LogPath $script:LogPath
}
catch {
    # Fatal error that escaped the GUI's own try/catch (e.g. XAML load failure)
    $errMsg = $_.Exception.Message
    try {
        [System.Windows.MessageBox]::Show(
            ("BuildMaster encountered a fatal error and cannot continue:`n`n{0}`n`n" +
             "Check the console window for the full error details.") -f $errMsg,
            'BuildMaster - Fatal Error',
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error) | Out-Null
    }
    catch { }
    Write-Error $errMsg
}
