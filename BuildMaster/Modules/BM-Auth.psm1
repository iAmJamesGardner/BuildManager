#Requires -Version 5.1
<#
.SYNOPSIS
    BM-Auth - Authentication and session-context module for BuildMaster
.DESCRIPTION
    Determines HOW the current operator can perform privileged actions by
    inspecting the running account domain and AD group membership.

    ACCOUNT DETECTION LOGIC:
    +----- Running as PRIVDOMAIN account? -----+
    |  YES -> IsPrivilegedSession = $true       |
    |         API calls need explicit regular   |
    |         creds (different account).        |
    |         Admin commands run as current     |
    |         session (token has the rights).   |
    +-------------------------------------------+
    |  NO  -> Regular account. Check EPM.       |
    |   +-- Member of EPM mail group            |
    |   |   AND EPM agent present?              |
    |   |  YES -> ElevationMethod = 'EPM'       |
    |   |         API calls use SSO (no prompt) |
    |   |         Admin via EPM (transparent)   |
    |   +-- NOT a member / agent missing?       |
    |      -> ElevationMethod = 'PrivAccount'   |
    |         QA/UAT/PROD: BLOCKED              |
    |         DEV: WhatIf mode available        |
    +-------------------------------------------+

    CREDENTIAL RULES:
    - Regular cred  : needed ONLY in a privileged session (priv user calling
                      API endpoints that require a regular account).
                      Regular sessions use UseDefaultCredentials (SSO).
    - Priv cred     : never stored manually. Elevation is via current session
                      token (priv session) or CyberArk EPM (regular session).
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
#  CONFIGURATION  (edit these to match your environment)
# -----------------------------------------------------------------------------

# NetBIOS domain name for privileged accounts (e.g. 'MYDOM_PRIV')
$script:PrivDomain = 'PRIVDOMAIN'

# PLACEHOLDER: Replace with the actual DN or CN of the CyberArk EPM
# authorisation mail group. Example:
#   'CN=GRP-CyberArk-EPM-Authorized,OU=Mail Groups,DC=corp,DC=example,DC=com'
$script:EPMMailGroup = 'PLACEHOLDER-EPM-MAIL-GROUP'

# -----------------------------------------------------------------------------
#  MODULE STATE
# -----------------------------------------------------------------------------
$script:SessionContext = $null   # [hashtable] - set by Initialize-BMSessionContext
$script:RegularCred    = $null   # [PSCredential] - only stored in priv sessions
$script:WhatIfMode     = $false  # DEV-only bypass; set via Set-BMWhatIfMode

# -----------------------------------------------------------------------------
#  SESSION CONTEXT INITIALISATION
# -----------------------------------------------------------------------------

function Initialize-BMSessionContext {
    <#
    .SYNOPSIS
        Detects the current operator's session context.
        Must be called once at startup before any other auth function.
    .PARAMETER Environment  dev / qa / uat / prod  (enforces block on restricted envs)
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('dev','qa','uat','prod')]
        [string]$Environment
    )

    $domain   = $env:USERDOMAIN
    $userName = $env:USERNAME
    $isPriv   = $domain -ieq $script:PrivDomain

    $ctx = [ordered]@{
        UserName            = $userName
        Domain              = $domain
        FullUser            = ('{0}\{1}' -f $domain, $userName)
        IsPrivilegedSession = $isPriv
        ElevationMethod     = 'Unknown'   # DirectSession | EPM | PrivAccount
        EPMAgentPresent     = $false
        EPMGroupMember      = $false
        Environment         = $Environment
        Blocked             = $false
        BlockReason         = ''
    }

    if ($isPriv) {
        $ctx.ElevationMethod = 'DirectSession'
        Write-Host ('[Auth] Privileged session detected ({0})' -f $ctx.FullUser) -ForegroundColor Cyan
    }
    else {
        $ctx.EPMAgentPresent = Test-BMCyberArkEPMAvailable
        $ctx.EPMGroupMember  = Test-BMEPMGroupMembership -UserName $userName

        if ($ctx.EPMGroupMember -and $ctx.EPMAgentPresent) {
            $ctx.ElevationMethod = 'EPM'
            Write-Host ('[Auth] Regular session with CyberArk EPM ({0})' -f $ctx.FullUser) -ForegroundColor Cyan
        }
        else {
            $ctx.ElevationMethod = 'PrivAccount'

            $reason = ''
            if (-not $ctx.EPMGroupMember -and -not $ctx.EPMAgentPresent) {
                $reason = 'Not a member of the EPM authorisation group and EPM agent is not running.'
            }
            elseif (-not $ctx.EPMGroupMember) {
                $reason = 'Account is not a member of the CyberArk EPM authorisation group.'
            }
            else {
                $reason = 'CyberArk EPM agent is not running on this machine.'
            }

            if ($Environment -in @('qa','uat','prod')) {
                $ctx.Blocked     = $true
                $ctx.BlockReason = (
                    "This tool requires a privileged account or CyberArk EPM in the " +
                    "$($Environment.ToUpper()) environment.`n`n$reason`n`n" +
                    "Close this window and relaunch the tool while logged on as your " +
                    "$script:PrivDomain account."
                )
                Write-Warning ('[Auth] Session BLOCKED in {0}: {1}' -f $Environment, $reason)
            }
            else {
                # DEV - not blocked but WhatIf mode is available
                Write-Host ('[Auth] Regular session without EPM in DEV ({0}) - WhatIf available.' -f $ctx.FullUser) -ForegroundColor Yellow
            }
        }
    }

    $script:SessionContext = $ctx
    return $ctx
}

function Get-BMSessionContext {
    <#  .SYNOPSIS  Returns the current session context hashtable.  #>
    if ($null -eq $script:SessionContext) {
        throw 'Session context not initialised. Call Initialize-BMSessionContext first.'
    }
    return $script:SessionContext
}

function Test-BMSessionBlocked {
    <#  .SYNOPSIS  Returns $true when the session is blocked from admin actions.  #>
    if ($null -eq $script:SessionContext) { return $false }
    return [bool]$script:SessionContext.Blocked
}

# -----------------------------------------------------------------------------
#  WHAT-IF MODE  (DEV environment only - no commands are actually sent)
# -----------------------------------------------------------------------------

function Set-BMWhatIfMode {
    [CmdletBinding()]
    param([Parameter(Mandatory)][bool]$Enabled)
    $script:WhatIfMode = $Enabled
    $msg = if ($Enabled) { 'WHAT-IF mode ENABLED - privileged commands will be simulated.' }
           else          { 'WHAT-IF mode disabled.' }
    Write-Host ('[Auth] {0}' -f $msg) -ForegroundColor $(if ($Enabled) { 'DarkYellow' } else { 'DarkGray' })
}

function Get-BMWhatIfMode {
    return $script:WhatIfMode
}

# -----------------------------------------------------------------------------
#  CREDENTIAL POPUP HELPER  (always a GUI dialog - never the console)
# -----------------------------------------------------------------------------

function Invoke-BMCredentialPopup {
    <#
    .SYNOPSIS
        Shows a Windows Security credential dialog.
        Uses $host.UI.PromptForCredential so the popup always appears
        regardless of how the script was launched (console, shortcut, etc.).
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCredential])]
    param(
        [string]$Title    = 'BuildMaster - Credentials Required',
        [string]$Message  = 'Enter your credentials.',
        [string]$UserHint = ''
    )

    $cred = $host.UI.PromptForCredential(
        $Title,
        $Message,
        $UserHint,
        ''   # targetName - left blank; appears as domain label in the dialog
    )

    if ($null -eq $cred) {
        throw 'Credential entry was cancelled.'
    }
    return $cred
}

# -----------------------------------------------------------------------------
#  REGULAR ACCOUNT CREDENTIALS
# -----------------------------------------------------------------------------

function Get-BMRegularCredential {
    <#
    .SYNOPSIS
        Returns the regular-account credential for API calls.
    .DESCRIPTION
        Privileged session : returns an explicitly stored PSCredential
                             (prompted via GUI popup if not yet stored).
                             The priv user's own token cannot reach BM/VW APIs.
        Regular session    : returns $null.
                             Callers pass $null to Invoke-BMRestCall which then
                             uses -UseDefaultCredentials (current Windows token).
    .PARAMETER Force  Re-prompt even if a credential is already stored.
    #>
    [CmdletBinding()]
    [OutputType([System.Management.Automation.PSCredential])]
    param([switch]$Force)

    $ctx = Get-BMSessionContext

    if (-not $ctx.IsPrivilegedSession) {
        # Regular session - current Windows token IS the regular account.
        # Return $null to signal "use SSO / UseDefaultCredentials".
        return $null
    }

    # Privileged session - needs a separate regular-account credential.
    if ($null -eq $script:RegularCred -or $Force) {
        # Strip common privileged-account suffixes to offer a sensible username hint.
        $hint = ($env:USERNAME -replace '_SUP$|_ADM$|_PA$|_PRIV$', '')
        $hint = ('CORP\{0}' -f $hint)

        $script:RegularCred = Invoke-BMCredentialPopup `
            -Title    'Regular Account Credentials' `
            -Message  ("Enter your REGULAR (non-admin) account credentials.`n`nThese are used for BuildMaster and VirtualWorks API calls.`nYou are currently running as the privileged account $($ctx.FullUser).") `
            -UserHint $hint
    }
    return $script:RegularCred
}

function Set-BMRegularCredential {
    [CmdletBinding()]
    param([Parameter(Mandatory)][System.Management.Automation.PSCredential]$Credential)
    $script:RegularCred = $Credential
}

function Test-BMRegularCredentialSet {
    <#
    .SYNOPSIS
        Returns $true when a usable regular credential exists.
        Regular sessions always return $true (SSO is always available).
        Privileged sessions return $true only once the credential has been stored.
    #>
    $ctx = Get-BMSessionContext
    if (-not $ctx.IsPrivilegedSession) { return $true }
    return ($null -ne $script:RegularCred)
}

# -----------------------------------------------------------------------------
#  CYBERARK EPM DETECTION
# -----------------------------------------------------------------------------

function Test-BMCyberArkEPMAvailable {
    <#  .SYNOPSIS  Returns $true if the CyberArk EPM agent is installed and running.  #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $regPaths = @(
        'HKLM:\SOFTWARE\CyberArk\Endpoint Privilege Manager',
        'HKLM:\SOFTWARE\WOW6432Node\CyberArk\Endpoint Privilege Manager'
    )
    foreach ($p in $regPaths) {
        if (Test-Path $p) { return $true }
    }

    $svc = Get-Service -Name 'CyberArkEPMService' -ErrorAction SilentlyContinue
    if ($null -ne $svc -and $svc.Status -eq 'Running') { return $true }

    $svc2 = Get-Service -Name 'CyberArk Endpoint Privilege Manager' -ErrorAction SilentlyContinue
    if ($null -ne $svc2 -and $svc2.Status -eq 'Running') { return $true }

    return $false
}

function Test-BMEPMGroupMembership {
    <#
    .SYNOPSIS
        Uses ADSISearcher to check if the user is a member of the
        CyberArk EPM authorisation mail group.
    .DESCRIPTION
        PLACEHOLDER - set $script:EPMMailGroup at the top of this file.

        Checks the 'memberOf' attribute of the user's AD object.
        For nested group membership you may need LDAP_MATCHING_RULE_IN_CHAIN
        (1.2.840.113556.1.4.1941) in the filter instead.
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$UserName = $env:USERNAME
    )

    try {
        if ($script:EPMMailGroup -eq 'PLACEHOLDER-EPM-MAIL-GROUP') {
            Write-Warning '[Auth] EPM mail group not configured. Edit $script:EPMMailGroup in BM-Auth.psm1.'
            return $false
        }

        # Direct memberOf attribute search
        $searcher = [adsisearcher]("(&(objectCategory=person)(objectClass=user)(sAMAccountName=$UserName))")
        $searcher.PropertiesToLoad.Add('memberOf') | Out-Null
        $result = $searcher.FindOne()

        if ($null -eq $result) {
            Write-Warning ('[Auth] ADSISearcher could not locate user ''{0}'' in the directory.' -f $UserName)
            return $false
        }

        $memberOf   = @($result.Properties['memberOf'])
        $matchFound = @($memberOf | Where-Object { $_ -like "*$($script:EPMMailGroup)*" })
        return ($matchFound.Count -gt 0)
    }
    catch {
        Write-Warning ('[Auth] EPM group membership check failed: {0}' -f $_.Exception.Message)
        return $false
    }
}

# -----------------------------------------------------------------------------
#  PRIVILEGED COMMAND EXECUTION
# -----------------------------------------------------------------------------

function Invoke-BMPrivilegedCommand {
    <#
    .SYNOPSIS
        Executes a scriptblock on a remote machine using the appropriate
        elevation method for the current session.
    .DESCRIPTION
        DirectSession  : current token already has admin rights; runs via
                         Invoke-Command with Negotiate auth.
        EPM            : current token is a regular account; the CyberArk EPM
                         agent intercepts the outbound WinRM request and attaches
                         an elevated token per policy. Code is identical to
                         DirectSession - EPM elevation is completely transparent.
        WhatIf mode    : logs what WOULD be run but does NOT execute anything.
                         Enabled by Set-BMWhatIfMode; available in DEV only.
        Blocked        : throws immediately (GUI should prevent reaching here).
    .PARAMETER PrivInfo  Legacy parameter accepted but no longer used.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory)][string]$TargetComputer,
        [Parameter()][hashtable]$PrivInfo   # retained for backward compatibility
    )

    $ctx = Get-BMSessionContext

    if ($script:WhatIfMode) {
        # WhatIf - simulate without executing
        $cmdText = $ScriptBlock.ToString().Trim() -replace '\s+', ' '
        Write-BMLog -MachineName $TargetComputer `
                    -Message ('[WHAT-IF] Would execute on {0}: {1}' -f $TargetComputer, $cmdText) `
                    -Level Warning
        return $null
    }

    if ($ctx.Blocked) {
        throw ('Session is blocked from performing privileged actions. {0}' -f $ctx.BlockReason)
    }

    # DirectSession and EPM are mechanically identical from PowerShell's
    # perspective - the EPM agent does its work at the OS/WinRM level.
    return Invoke-Command `
        -ComputerName   $TargetComputer `
        -ScriptBlock    $ScriptBlock `
        -Authentication Negotiate `
        -ErrorAction    Stop
}

# -----------------------------------------------------------------------------
#  CLEANUP
# -----------------------------------------------------------------------------

function Clear-BMCredentials {
    <#  .SYNOPSIS  Clears stored credentials at session end.  #>
    $script:RegularCred = $null
    $script:WhatIfMode  = $false
    Write-Host '[Auth] Credentials cleared.' -ForegroundColor DarkGray
}

# -----------------------------------------------------------------------------
#  EXPORTS
# -----------------------------------------------------------------------------
Export-ModuleMember -Function @(
    # Session context
    'Initialize-BMSessionContext',
    'Get-BMSessionContext',
    'Test-BMSessionBlocked',
    # WhatIf mode
    'Set-BMWhatIfMode',
    'Get-BMWhatIfMode',
    # Regular credentials
    'Get-BMRegularCredential',
    'Set-BMRegularCredential',
    'Test-BMRegularCredentialSet',
    'Invoke-BMCredentialPopup',
    # EPM detection
    'Test-BMCyberArkEPMAvailable',
    'Test-BMEPMGroupMembership',
    # Privileged execution
    'Invoke-BMPrivilegedCommand',
    # Cleanup
    'Clear-BMCredentials'
)
