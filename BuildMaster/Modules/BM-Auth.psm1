#Requires -Version 5.1
<#
.SYNOPSIS
    BM-Auth - Authentication module for BuildMaster
.DESCRIPTION
    Manages regular and privileged account credentials for the session.

    Regular account  : Always known; obtained via Get-Credential prompt.
    Privileged account: Password unknown to the operator.
                        Two elevation paths are supported:
                        (A) CyberArk EPM  - if the EPM agent is installed and
                            a policy permits elevation, the regular account's
                            token is automatically elevated; no priv password needed.
                        (B) Explicit priv credential - operator supplies the
                            privileged account credentials at runtime (e.g. via
                            a PSM / CyberArk session or a secondary sign-in).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- In-session credential store -----------------------------------------------
$script:RegularCred     = $null   # [PSCredential]
$script:PrivilegedInfo  = $null   # @{ Method='EPM'|'Credential'; Credential=$null|[PSCredential] }

# -----------------------------------------------------------------------------
#  REGULAR ACCOUNT
# -----------------------------------------------------------------------------

function Get-BMRegularCredential {
    <#
    .SYNOPSIS  Returns the stored regular credential, prompting if not yet set.
    .PARAMETER Force  Re-prompt even if a credential is already stored.
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCredential])]
    param([switch]$Force)

    if ($null -eq $script:RegularCred -or $Force) {
        $script:RegularCred = Get-Credential -Message 'Enter your REGULAR account credentials'
    }
    return $script:RegularCred
}

function Set-BMRegularCredential {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential)
    $script:RegularCred = $Credential
}

function Test-BMRegularCredentialSet {
    return ($null -ne $script:RegularCred)
}

# -----------------------------------------------------------------------------
#  PRIVILEGED ACCOUNT
# -----------------------------------------------------------------------------

function Get-BMPrivilegedInfo {
    <#
    .SYNOPSIS
        Returns the privileged auth info hashtable, prompting / detecting if needed.
    .DESCRIPTION
        Returns: @{ Method = 'EPM' | 'Credential'
                    Credential = $null | [PSCredential] }

        EPM path     regular account token is elevated transparently by
                      the CyberArk EPM agent; no priv password is stored.
        Credential   operator provided explicit priv credentials at startup.
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param([switch]$Force)

    if ($null -ne $script:PrivilegedInfo -and -not $Force) {
        return $script:PrivilegedInfo
    }

    $epmAvailable = Test-BMCyberArkEPMAvailable

    if ($epmAvailable) {
        Write-Host '[Auth] CyberArk EPM agent detected - using policy-based elevation.' -ForegroundColor Cyan
        $script:PrivilegedInfo = @{ Method = 'EPM'; Credential = $null }
    }
    else {
        Write-Host '[Auth] EPM not detected - prompting for privileged credentials.' -ForegroundColor Yellow
        $privCred = Get-Credential -Message 'Enter PRIVILEGED account credentials (e.g. DOMAIN\adm_jsmith)'
        $script:PrivilegedInfo = @{ Method = 'Credential'; Credential = $privCred }
    }

    return $script:PrivilegedInfo
}

function Set-BMPrivilegedInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable]$PrivInfo)
    $script:PrivilegedInfo = $PrivInfo
}

function Test-BMPrivilegedInfoSet {
    return ($null -ne $script:PrivilegedInfo)
}

# -----------------------------------------------------------------------------
#  CYBERARK EPM DETECTION
# -----------------------------------------------------------------------------

function Test-BMCyberArkEPMAvailable {
    <#
    .SYNOPSIS
        Returns $true if the CyberArk EPM agent appears to be installed and running.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    # Registry hive check (EPM installs a key under CyberArk)
    $regPaths = @(
        'HKLM:\SOFTWARE\CyberArk\Endpoint Privilege Manager',
        'HKLM:\SOFTWARE\WOW6432Node\CyberArk\Endpoint Privilege Manager'
    )
    foreach ($p in $regPaths) {
        if (Test-Path $p) { return $true }
    }

    # Service check
    $svc = Get-Service -Name 'CyberArkEPMService' -ErrorAction SilentlyContinue
    if ($null -ne $svc -and $svc.Status -eq 'Running') { return $true }

    # Alternative service name used in some EPM versions
    $svc2 = Get-Service -Name 'CyberArk Endpoint Privilege Manager' -ErrorAction SilentlyContinue
    if ($null -ne $svc2 -and $svc2.Status -eq 'Running') { return $true }

    return $false
}

# -----------------------------------------------------------------------------
#  PRIVILEGED COMMAND EXECUTION
# -----------------------------------------------------------------------------

function Invoke-BMPrivilegedCommand {
    <#
    .SYNOPSIS
        Executes a scriptblock on a remote machine using the privileged auth method.
    .DESCRIPTION
        EPM path     : Uses Negotiate auth; the EPM agent on the *calling* machine
                       intercepts the outbound PSRemoting request and attaches an
                       elevated token per the applicable policy.
        Credential   : Uses the explicit priv [PSCredential] for PSRemoting.
    .PARAMETER ScriptBlock  The code to run remotely.
    .PARAMETER TargetComputer  The machine to run it on.
    .PARAMETER PrivInfo  Auth info from Get-BMPrivilegedInfo (optional; fetched if omitted).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,

        [Parameter(Mandatory)]
        [string]$TargetComputer,

        [Parameter()]
        [hashtable]$PrivInfo
    )

    if ($null -eq $PrivInfo) {
        $PrivInfo = Get-BMPrivilegedInfo
    }

    $invokeParams = @{
        ComputerName = $TargetComputer
        ScriptBlock  = $ScriptBlock
        ErrorAction  = 'Stop'
    }

    switch ($PrivInfo.Method) {
        'EPM' {
            # EPM handles elevation transparently via Negotiate
            $invokeParams.Authentication = 'Negotiate'
            return Invoke-Command @invokeParams
        }
        'Credential' {
            $invokeParams.Credential     = $PrivInfo.Credential
            $invokeParams.Authentication = 'Negotiate'
            return Invoke-Command @invokeParams
        }
        default {
            throw "Unknown privileged auth method: $($PrivInfo.Method)"
        }
    }
}

function Clear-BMCredentials {
    <#  .SYNOPSIS  Clears all stored credentials (e.g. on session end).  #>
    $script:RegularCred    = $null
    $script:PrivilegedInfo = $null
    Write-Host '[Auth] Credentials cleared.' -ForegroundColor DarkGray
}

# -----------------------------------------------------------------------------
#  EXPORTS
# -----------------------------------------------------------------------------
Export-ModuleMember -Function @(
    'Get-BMRegularCredential',
    'Set-BMRegularCredential',
    'Test-BMRegularCredentialSet',
    'Get-BMPrivilegedInfo',
    'Set-BMPrivilegedInfo',
    'Test-BMPrivilegedInfoSet',
    'Test-BMCyberArkEPMAvailable',
    'Invoke-BMPrivilegedCommand',
    'Clear-BMCredentials'
)
