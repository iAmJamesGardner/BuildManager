#Requires -Version 5.1
#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Registers the BuildMaster JEA (Just Enough Administration) endpoint.
.DESCRIPTION
    +-------------------------------------------------------------------------+
    |  SIDE QUEST: JEA + CyberArk EPM                                         |
    |                                                                         |
    |  Two complementary privilege elevation strategies:                      |
    |                                                                         |
    |  1. Microsoft JEA - Locks down *what* can be done via a constrained     |
    |     PS remoting endpoint. Users connect via New-PSSession -ConfigName   |
    |     'BuildMasterAdmin', and can only run the whitelisted commands        |
    |     defined in RoleCapabilities\RebuildAdmin.psrc.                      |
    |     The endpoint runs as a Group Managed Service Account (gMSA) or       |
    |     virtual account - so no human privileged password is ever needed.   |
    |                                                                         |
    |  2. CyberArk EPM policy - Controls *who* can trigger the elevation.     |
    |     EPM's "Application Elevation Policy" intercepts the PSRemoting       |
    |     connection attempt and auto-approves it based on:                   |
    |       * the calling user's group membership                             |
    |       * the target endpoint name / hash                                 |
    |       * optional MFA step-up (CyberArk PVWA integration)               |
    |                                                                         |
    |  Together: EPM decides if the user *can* connect; JEA decides what      |
    |  they *can do* once connected. No privileged passwords are exposed.     |
    +-------------------------------------------------------------------------+

    PREREQUISITES (run once per target machine or via GPO):
      1. The gMSA 'DOMAIN\svc_bm_jea$' must exist in AD and be authorized
         on the target machine:
           Add-ADComputerServiceAccount -Identity <TARGET> -ServiceAccount svc_bm_jea$
           Install-ADServiceAccount     -Identity svc_bm_jea$
      2. CyberArk EPM policy "BuildMaster-JEA-Allow" must exist in the EPM
         console targeting the endpoint name 'BuildMasterAdmin'.
      3. The calling user must be in the AD group defined in $AllowedGroup.

    DEPLOYMENT:
      Run this script on each *target* machine (the machines being rebuilt),
      not on the operator's workstation.

        \\FileServer\ITTools\BuildMaster\<env>\JEA\Register-JEAEndpoint.ps1

    USAGE AFTER REGISTRATION:
        # From operator workstation (via BM-JEA module):
        $session = New-BMJEASession -TargetComputer 'PC001'
        Invoke-Command -Session $session -ScriptBlock { Restart-Computer -Force }
        Remove-PSSession $session
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    # AD group whose members are allowed to use the JEA endpoint
    [string]$AllowedGroup = 'DOMAIN\BuildMaster-Operators',

    # Group Managed Service Account (gMSA) to run the JEA session as.
    # Set to $null to use a virtual account instead.
    [string]$gMSAAccount = 'DOMAIN\svc_bm_jea$',

    # JEA endpoint registration name
    [string]$EndpointName = 'BuildMasterAdmin',

    # Where to deploy the role capability files on this machine
    [string]$RoleCapabilitiesRoot = "$env:ProgramFiles\WindowsPowerShell\Modules\BuildMasterJEA\RoleCapabilities"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host '[ JEA Registration ] BuildMaster JEA Endpoint' -ForegroundColor Cyan
Write-Host ('  Endpoint    : {0}' -f $EndpointName)        -ForegroundColor Gray
Write-Host ('  Allowed group: {0}' -f $AllowedGroup)       -ForegroundColor Gray
Write-Host ('  Run-as gMSA  : {0}' -f $gMSAAccount)        -ForegroundColor Gray

# -----------------------------------------------------------------------------
#  STEP 1: Deploy role capability file to module path
# -----------------------------------------------------------------------------

if (-not (Test-Path $RoleCapabilitiesRoot)) {
    New-Item -ItemType Directory -Path $RoleCapabilitiesRoot -Force | Out-Null
    Write-Host "  [+] Created: $RoleCapabilitiesRoot" -ForegroundColor Green
}

$psrcSource = Join-Path $PSScriptRoot 'RoleCapabilities\RebuildAdmin.psrc'
$psrcDest   = Join-Path $RoleCapabilitiesRoot 'RebuildAdmin.psrc'

if (-not (Test-Path $psrcSource)) {
    throw "Role capability file not found: $psrcSource"
}

Copy-Item -Path $psrcSource -Destination $psrcDest -Force
Write-Host "  [+] Deployed role capability to: $psrcDest" -ForegroundColor Green

# -----------------------------------------------------------------------------
#  STEP 2: Create Session Configuration file (.pssc) in a temp location
# -----------------------------------------------------------------------------

$psscPath = Join-Path $env:TEMP 'BuildMasterAdmin.pssc'

# Build the role definition hashtable: maps the allowed group to the role capability
$roleDefinition = @{
    $AllowedGroup = @{ RoleCapabilities = 'RebuildAdmin' }
}

$psscParams = @{
    Path                 = $psscPath
    SessionType          = 'RestrictedRemoteServer'   # No interactive shell; commands only
    LanguageMode         = 'ConstrainedLanguage'
    ExecutionPolicy      = 'AllSigned'
    RoleDefinitions      = $roleDefinition
    TranscriptDirectory  = 'C:\BuildMaster\JEA-Transcripts'   # Audit trail
}

# Prefer gMSA; fall back to virtual account
if (-not [string]::IsNullOrWhiteSpace($gMSAAccount)) {
    $psscParams.GroupManagedServiceAccount = $gMSAAccount
} else {
    $psscParams.RunAsVirtualAccount = $true
    Write-Host '  [!] No gMSA specified - using virtual account (less auditable).' -ForegroundColor Yellow
}

New-PSSessionConfigurationFile @psscParams
Write-Host "  [+] Session config file created: $psscPath" -ForegroundColor Green

# -----------------------------------------------------------------------------
#  STEP 3: Ensure transcript directory exists
# -----------------------------------------------------------------------------

$transcriptDir = $psscParams.TranscriptDirectory
if (-not (Test-Path $transcriptDir)) {
    New-Item -ItemType Directory -Path $transcriptDir -Force | Out-Null
    # Restrict write access - only SYSTEM/gMSA writes transcripts; admins can read
    $acl = Get-Acl $transcriptDir
    $acl.SetAccessRuleProtection($true, $false)
    $adminRule  = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    'BUILTIN\Administrators', 'ReadAndExecute', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                    'NT AUTHORITY\SYSTEM', 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
    $acl.AddAccessRule($adminRule)
    $acl.AddAccessRule($systemRule)
    Set-Acl -Path $transcriptDir -AclObject $acl
    Write-Host "  [+] Transcript directory secured: $transcriptDir" -ForegroundColor Green
}

# -----------------------------------------------------------------------------
#  STEP 4: Register (or re-register) the endpoint
# -----------------------------------------------------------------------------

# Remove existing registration with the same name
$existing = Get-PSSessionConfiguration -Name $EndpointName -ErrorAction SilentlyContinue
if ($null -ne $existing) {
    Write-Host "  [~] Removing existing endpoint '$EndpointName'..." -ForegroundColor Yellow
    Unregister-PSSessionConfiguration -Name $EndpointName -Force
}

if ($PSCmdlet.ShouldProcess($env:COMPUTERNAME, "Register JEA endpoint '$EndpointName'")) {
    Register-PSSessionConfiguration `
        -Name    $EndpointName `
        -Path    $psscPath `
        -Force   `
        -NoServiceRestart   # We'll restart WinRM manually below

    Write-Host "  [+] Endpoint '$EndpointName' registered." -ForegroundColor Green
}

# -----------------------------------------------------------------------------
#  STEP 5: Restart WinRM to activate the new endpoint
# -----------------------------------------------------------------------------

Write-Host '  [~] Restarting WinRM service...' -ForegroundColor Yellow
Restart-Service WinRM -Force
Write-Host '  [+] WinRM restarted.' -ForegroundColor Green

# -----------------------------------------------------------------------------
#  STEP 6: Verify
# -----------------------------------------------------------------------------

$registered = Get-PSSessionConfiguration -Name $EndpointName -ErrorAction SilentlyContinue
if ($null -ne $registered) {
    Write-Host "`n  [+] JEA endpoint '$EndpointName' is active." -ForegroundColor Green
    Write-Host "    Enabled    : $($registered.Enabled)"    -ForegroundColor Gray
    Write-Host "    Permission : $($registered.Permission)" -ForegroundColor Gray
} else {
    Write-Error "Registration appeared to succeed but endpoint not found. Check WinRM event log."
}

# -----------------------------------------------------------------------------
#  STEP 7: CyberArk EPM integration note
# -----------------------------------------------------------------------------

Write-Host @'

  ===================================================================
  CYBERARK EPM POLICY SETUP (manual - done in EPM management console)
  ===================================================================

  Create an "Application Elevation Policy" in EPM with:
    Policy name  : BuildMaster-JEA-Allow
    Target       : C:\Windows\System32\wsmprovhost.exe
                   (the WinRM host process that JEA runs under)
    Conditions   :
      * Called by: powershell.exe
      * Argument contains: -ConfigurationName BuildMasterAdmin
      * Caller user in AD group: DOMAIN\BuildMaster-Operators
    Action       : Elevate (run as gMSA svc_bm_jea$ or virtual account)
    Audit        : Full - log to CyberArk PVWA

  This means:
    - The regular operator account triggers wsmprovhost.exe via PSRemoting
    - EPM sees the call matches the policy conditions
    - EPM transparently elevates the process without exposing any password
    - JEA then further constrains what the elevated session can do

  To test the endpoint from the operator workstation:
    $s = New-PSSession -ComputerName <target> -ConfigurationName BuildMasterAdmin
    Invoke-Command -Session $s { Get-Command }
    Remove-PSSession $s

'@ -ForegroundColor Cyan

Write-Host '  Registration complete.' -ForegroundColor Green
