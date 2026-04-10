#Requires -Version 5.1
<#
.SYNOPSIS
    BM-API - API wrapper module for BuildMaster and VirtualWorks
.DESCRIPTION
    All outbound HTTP calls live here. No external modules - uses only
    the built-in Invoke-RestMethod (PS 5.1 / .NET 4.x).

    BuildMaster API  : http://bm.zz.com/new-api
    VirtualWorks API : http://vw.zz.com/vwapi/Desktop

    ENDPOINT MAP:
        Stage machine    : POST  /v1/machinebuild/StageMachineByName
        Get build by PC  : GET   /v1/machinebuild?computerName=<name>   (to get BuildId only)
        Get build inst.  : GET   /v1/machinebuild/instance/<BuildId>    (for BuildTimes / stage)

        VW desktop check : GET   /Desktops?hostnames=<FQDN>
                           Response: .Results[].DesktopType
                               "VM" | "Moonshot" -> virtual/DiDC (use VW reboot)
                               "Non-DiDC" | no results -> physical (use privileged reboot)
        VW reboot        : POST  /Desktops   body: { "Id": "<FQDN>" }

    CREDENTIAL PASSING:
        Passing a [PSCredential] -> Invoke-RestMethod -Credential
        Passing $null            -> Invoke-RestMethod -UseDefaultCredentials (Windows SSO)
        Regular sessions use SSO. Privileged sessions pass an explicit credential.

    BuildMaster instance JSON shape (from real sample):
    {
      "BuildId"          : "251637f8-...",
      "InstanceId"       : "63a9f946-...",
      "ComputerName"     : "MACHINENAME",
      "BuildTimes"       : [
        { "BuildTime":"2026-03-23T23:34:49Z", "IsCompletedTime":true,  ... },
        { "BuildTime":"2026-03-23T18:28:32Z", "IsOSCompletedTime":true,...},
        { "BuildTime":"2026-03-23T17:59:50Z", "IsStartedTime":true,    ...},
        { "BuildTime":"2026-03-23T13:40:19Z", "IsStagedTime":true,     ...}
      ],
      "CreatedBy":"...", "CreatedOn":"..."
    }
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Base URLs / FQDN config ---------------------------------------------------
$script:BMBaseUrl    = 'http://bm.zz.com/new-api'
$script:VWBaseUrl    = 'http://vw.zz.com/vwapi/Desktop'
# VirtualWorks requires FQDN; BuildMaster uses short computer name only.
# Override at runtime with Set-BMAPIConfig if the suffix differs per environment.
$script:VWFqdnSuffix = 'xxxx.yy.com'

# -----------------------------------------------------------------------------
#  CONFIG (runtime override)
# -----------------------------------------------------------------------------

function Set-BMAPIConfig {
    <#
    .SYNOPSIS  Overrides base URL or FQDN suffix at runtime (e.g. per environment).
    #>
    [CmdletBinding()]
    param(
        [string]$VWFqdnSuffix,
        [string]$BMBaseUrl,
        [string]$VWBaseUrl
    )
    if ($PSBoundParameters.ContainsKey('VWFqdnSuffix')) { $script:VWFqdnSuffix = $VWFqdnSuffix }
    if ($PSBoundParameters.ContainsKey('BMBaseUrl'))    { $script:BMBaseUrl    = $BMBaseUrl    }
    if ($PSBoundParameters.ContainsKey('VWBaseUrl'))    { $script:VWBaseUrl    = $VWBaseUrl    }
}

# -- Private hostname helpers --------------------------------------------------

function Get-VWFqdn {
    # Returns a FQDN for VirtualWorks API calls.
    # If $ComputerName already contains a dot it is returned unchanged;
    # otherwise $script:VWFqdnSuffix is appended.
    param([string]$ComputerName)
    if ($ComputerName -match '\.') { return $ComputerName }
    return ('{0}.{1}' -f $ComputerName, $script:VWFqdnSuffix)
}

function Get-BMHostname {
    # Returns the short (pre-dot) hostname for BuildMaster API calls.
    param([string]$ComputerName)
    return $ComputerName.Split('.')[0]
}

# -----------------------------------------------------------------------------
#  INTERNAL REST HELPER
# -----------------------------------------------------------------------------

function Invoke-BMRestCall {
    <#
    .SYNOPSIS  Internal REST helper with PS 5.1 error handling.
    .DESCRIPTION
        Pass a PSCredential for explicit auth.
        Pass $null to use -UseDefaultCredentials (Windows SSO for regular sessions).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method   = 'GET',
        [object]$Body,
        [System.Management.Automation.PSCredential]$Credential
    )

    $params = @{
        Uri             = $Uri
        Method          = $Method
        ContentType     = 'application/json'
        UseBasicParsing = $true
        ErrorAction     = 'Stop'
    }

    if ($null -ne $Body)       { $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress) }

    if ($null -ne $Credential) {
        $params.Credential = $Credential
    }
    else {
        # $null credential -> use the current Windows session token (SSO / Kerberos / NTLM)
        $params.UseDefaultCredentials = $true
    }

    try {
        return Invoke-RestMethod @params
    }
    catch [System.Net.WebException] {
        $resp       = $_.Exception.Response
        $statusCode = if ($null -ne $resp) { [int]$resp.StatusCode } else { 0 }

        $errorBody = ''
        if ($null -ne $resp) {
            try {
                $stream    = $resp.GetResponseStream()
                $reader    = New-Object System.IO.StreamReader($stream)
                $errorBody = $reader.ReadToEnd()
                $reader.Dispose()
            }
            catch { }
        }

        throw ('API Error [{0} {1}] HTTP {2} - {3}' -f $Method, $Uri, $statusCode, $errorBody).TrimEnd()
    }
    catch {
        throw ('API Error [{0} {1}] - {2}' -f $Method, $Uri, $_.Exception.Message)
    }
}

# -----------------------------------------------------------------------------
#  BUILDMASTER - STAGING
# -----------------------------------------------------------------------------

function Invoke-BMStage {
    <#
    .SYNOPSIS  Submits a staging request for a machine to BuildMaster.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $uri  = ('{0}/v1/machinebuild/StageMachineByName' -f $script:BMBaseUrl)
    $body = @{ ComputerName = (Get-BMHostname -ComputerName $ComputerName) }
    return Invoke-BMRestCall -Uri $uri -Method POST -Body $body -Credential $Credential
}

# -----------------------------------------------------------------------------
#  BUILDMASTER - BUILD DATA (used only to retrieve BuildId)
# -----------------------------------------------------------------------------

function Get-BMBuildData {
    <#
    .SYNOPSIS
        Queries BuildMaster for a machine's build record.
        Use this ONLY to obtain the BuildId - then use Get-BMBuildInstance
        for all subsequent status polling (see notes in module header).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $uri = ('{0}/v1/machinebuild?computerName={1}' -f
            $script:BMBaseUrl, [Uri]::EscapeDataString((Get-BMHostname -ComputerName $ComputerName)))
    return Invoke-BMRestCall -Uri $uri -Method GET -Credential $Credential
}

# -----------------------------------------------------------------------------
#  BUILDMASTER - BUILD INSTANCE (used for all stage monitoring)
# -----------------------------------------------------------------------------

function Get-BMBuildInstance {
    <#
    .SYNOPSIS
        Returns the build instance record for the given BuildId.
        This is the primary polling call during monitoring - the returned
        object contains the BuildTimes array used to determine build stage.
    .NOTES
        TODO: Confirm exact URI path for the build instance endpoint.
              Current assumption: /v1/machinebuild/instance/{BuildId}
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$BuildId,
        [System.Management.Automation.PSCredential]$Credential
    )
    $uri = ('{0}/v1/machinebuild/instance/{1}' -f
            $script:BMBaseUrl, [Uri]::EscapeDataString($BuildId))
    return Invoke-BMRestCall -Uri $uri -Method GET -Credential $Credential
}

# -----------------------------------------------------------------------------
#  BUILDMASTER - STATUS PARSING
# -----------------------------------------------------------------------------

function Get-BMCurrentStage {
    <#
    .SYNOPSIS
        Parses a BuildData / BuildInstance object and returns the current stage.
    .OUTPUTS  'Staged' | 'Started' | 'OSComplete' | 'Completed' | 'Unknown'
    .DESCRIPTION
        Inspects BuildTimes sorted newest-first.
        Priority (highest wins): Completed > OSComplete > Started > Staged.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][object]$BuildData)

    if ($null -eq $BuildData -or $null -eq $BuildData.BuildTimes) { return 'Unknown' }

    $sorted = @($BuildData.BuildTimes) | Sort-Object { [datetime]$_.BuildTime } -Descending

    foreach ($entry in $sorted) {
        if ($entry.IsCompletedTime   -eq $true) { return 'Completed'  }
        if ($entry.IsOSCompletedTime -eq $true) { return 'OSComplete' }
        if ($entry.IsStartedTime     -eq $true) { return 'Started'    }
        if ($entry.IsStagedTime      -eq $true) { return 'Staged'     }
    }
    return 'Unknown'
}

function Get-BMStageEntryTime {
    <#
    .SYNOPSIS  Returns the [datetime] when a specific stage was first recorded.
    #>
    [CmdletBinding()]
    [OutputType([nullable[datetime]])]
    param(
        [Parameter(Mandatory)][object]$BuildData,
        [Parameter(Mandatory)]
        [ValidateSet('Staged','Started','OSComplete','Completed')]
        [string]$Stage
    )

    if ($null -eq $BuildData.BuildTimes) { return $null }

    $flagMap = @{
        'Staged'     = 'IsStagedTime'
        'Started'    = 'IsStartedTime'
        'OSComplete' = 'IsOSCompletedTime'
        'Completed'  = 'IsCompletedTime'
    }
    $flag = $flagMap[$Stage]

    $match = @($BuildData.BuildTimes) |
             Where-Object { $_.$flag -eq $true } |
             Sort-Object  { [datetime]$_.BuildTime } |
             Select-Object -First 1

    if ($null -eq $match) { return $null }
    return [datetime]$match.BuildTime
}

# -----------------------------------------------------------------------------
#  VIRTUALWORKS - VM / DiDC DETECTION
# -----------------------------------------------------------------------------

function Get-VWDesktopInfo {
    <#
    .SYNOPSIS
        Queries VirtualWorks for a machine's desktop record by FQDN.
        Returns the raw response object (XmlDocument or PSObject depending on
        server Content-Type), or $null when the machine is not found (HTTP 404).

    .DESCRIPTION
        Actual XML response shape:
            <PagedResult xmlns="http://...com/api">
              <Result>
                <Desktop>
                  <DesktopType>VM</DesktopType>   <!-- or Moonshot / Non-DiDC -->
                  <HostName>machine.domain.com</HostName>
                  <Id>4259915</Id>
                  ...
                </Desktop>
              </Result>
            </PagedResult>

        Each hostname query returns exactly one Desktop record.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    try {
        $fqdn = Get-VWFqdn -ComputerName $ComputerName
        $uri  = ('{0}/Desktops?hostnames={1}' -f $script:VWBaseUrl, [Uri]::EscapeDataString($fqdn))
        return Invoke-BMRestCall -Uri $uri -Method GET -Credential $Credential
    }
    catch {
        # HTTP 404 means the machine is not registered in VirtualWorks (physical)
        if ($_.Exception.Message -match 'HTTP 404') { return $null }
        throw
    }
}

function Test-IsVirtualMachine {
    <#
    .SYNOPSIS
        Returns $true if VirtualWorks reports this machine as a VM or DiDC/Moonshot.
        Returns $false for physical machines, unknown machines, or on API error.

    .DESCRIPTION
        VW returns XML with a default namespace.  Direct dot-notation property access
        is unreliable against a default-namespaced XmlDocument in PowerShell 5.1, so
        we use XPath local-name() queries which are namespace-agnostic.

        DesktopType "VM" or "Moonshot"  -> $true  (reboot via VW API)
        DesktopType "Non-DiDC" / absent -> $false (reboot requires elevation)
        API unreachable / any error     -> $false (fail-safe: treat as Physical)
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )

    try {
        $info = Get-VWDesktopInfo -ComputerName $ComputerName -Credential $Credential

        # 404 / empty body -> not a VW-managed machine -> physical
        if ($null -eq $info) { return $false }

        # ------------------------------------------------------------------
        # Parse DesktopType from the XML response.
        #
        # The VW API returns XML with a default namespace
        # (xmlns="http://...com/api").  Dot-notation on XmlDocument can silently
        # return $null when elements are in a default namespace, so we use
        # XPath with local-name() which matches regardless of namespace.
        #
        # Target path:  /PagedResult/Result/Desktop/DesktopType
        # Each hostname query returns exactly one Desktop record.
        # ------------------------------------------------------------------
        $desktopType = $null

        if ($info -is [System.Xml.XmlNode]) {
            # Invoke-RestMethod returned an XmlDocument (XML Content-Type)
            $dtNode = $info.SelectSingleNode('//*[local-name()="DesktopType"]')
            if ($null -ne $dtNode) {
                $desktopType = $dtNode.InnerText.Trim()
            }
        }
        elseif ($null -ne $info.Result) {
            # Fallback: PSObject / JSON response with Result.Desktop structure
            $desktopType = $info.Result.Desktop.DesktopType
        }

        if ([string]::IsNullOrEmpty($desktopType)) {
            Write-Verbose ("No VW Desktop record found for '{0}'" -f $ComputerName)
            return $false
        }

        Write-Verbose ("VW DesktopType for '{0}': {1}" -f $ComputerName, $desktopType)
        return ($desktopType -in @('VM', 'Moonshot'))
    }
    catch {
        # VirtualWorks unreachable or returned an unexpected error.
        # Default to $false (treat as Physical) so the staging step is not blocked.
        Write-Warning ("VW machine-type check failed for '{0}': {1}  [Defaulting to Physical]" -f
                       $ComputerName, $_.Exception.Message)
        return $false
    }
}

# -----------------------------------------------------------------------------
#  VIRTUALWORKS - REBOOT
# -----------------------------------------------------------------------------

function Invoke-VWReboot {
    <#
    .SYNOPSIS  Sends a reboot command for a VM / DiDC through VirtualWorks.
    .DESCRIPTION
        POSTs to /Desktops with body { "Id": "<ComputerName>" }.
        Use only for machines where Test-IsVirtualMachine returns $true.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $uri  = ('{0}/Desktops' -f $script:VWBaseUrl)
    $body = @{ Id = (Get-VWFqdn -ComputerName $ComputerName) }
    return Invoke-BMRestCall -Uri $uri -Method POST -Body $body -Credential $Credential
}

# -----------------------------------------------------------------------------
#  EXPORTS
# -----------------------------------------------------------------------------
Export-ModuleMember -Function @(
    'Set-BMAPIConfig',
    'Invoke-BMStage',
    'Get-BMBuildData',
    'Get-BMBuildInstance',
    'Get-BMCurrentStage',
    'Get-BMStageEntryTime',
    'Get-VWDesktopInfo',
    'Test-IsVirtualMachine',
    'Invoke-VWReboot'
)
