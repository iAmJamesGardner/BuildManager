#Requires -Version 5.1
<#
.SYNOPSIS
    BM-API - API wrapper module for BuildMaster and VirtualWorks
.DESCRIPTION
    All outbound HTTP calls live here. No external modules - uses only
    the built-in Invoke-RestMethod (PS 5.1 / .NET 4.x).

    BuildMaster API  : http://bm.zz.com/new-api/v1
    VirtualWorks API : http://vw.zz.com/vwapi/Desktop

    BuildMaster JSON shape (from real sample):
    {
      "BuildId"          : "guid",
      "InstanceId"       : "guid",
      "ComputerName"     : "MACHINENAME",
      "Settings"         : 0,
      "InstanceSettings" : 2097168,
      "BuildTimes"       : [
        {
          "Id"               : "guid",
          "BuildId"          : "guid",
          "Message"          : "Build has completed",
          "Settings"         : 2097152,
          "BuildTime"        : "2026-03-23T23:34:49Z",
          "IsStagedTime"     : false,
          "IsStartedTime"    : false,
          "IsOSCompletedTime": false,
          "IsCompletedTime"  : true
        },
        ...
      ],
      "CreatedBy" : "username",
      "CreatedOn" : "2026-03-23T13:40:19Z"
    }
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Base URLs -----------------------------------------------------------------
$script:BMBaseUrl = 'http://bm.zz.com/new-api/v1'
$script:VWBaseUrl = 'http://vw.zz.com/vwapi/Desktop'

# -----------------------------------------------------------------------------
#  INTERNAL HELPER
# -----------------------------------------------------------------------------

function Invoke-BMRestCall {
    <#
    .SYNOPSIS  Internal REST helper - wraps Invoke-RestMethod with PS 5.1 error handling.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
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

    if ($null -ne $Body)       { $params.Body       = ($Body | ConvertTo-Json -Depth 10 -Compress) }
    if ($null -ne $Credential) { $params.Credential = $Credential }

    try {
        return Invoke-RestMethod @params
    }
    catch [System.Net.WebException] {
        $resp       = $_.Exception.Response
        $statusCode = if ($null -ne $resp) { [int]$resp.StatusCode } else { 0 }

        $errorBody  = ''
        if ($null -ne $resp) {
            try {
                $stream = $resp.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $errorBody = $reader.ReadToEnd()
                $reader.Dispose()
            } catch { }
        }

        throw "API Error [$Method $Uri] HTTP $statusCode - $errorBody".Trim()
    }
    catch {
        throw "API Error [$Method $Uri] - $($_.Exception.Message)"
    }
}

# -----------------------------------------------------------------------------
#  BUILDMASTER - BUILD DATA
# -----------------------------------------------------------------------------

function Get-BMBuildData {
    <#
    .SYNOPSIS
        Queries BuildMaster for the current build record of a machine.
        Returns the full object including the BuildTimes array.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $uri = '{0}/build?computer={1}' -f $script:BMBaseUrl, [Uri]::EscapeDataString($ComputerName)
    return Invoke-BMRestCall -Uri $uri -Method GET -Credential $Credential
}

function Get-BMLastBuildInstance {
    <#
    .SYNOPSIS
        Returns the most recent build instance record for a machine.
        Used as the "second scriptblock" query described in the requirements.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $uri = '{0}/build/instance/last?computer={1}' -f $script:BMBaseUrl, [Uri]::EscapeDataString($ComputerName)
    return Invoke-BMRestCall -Uri $uri -Method GET -Credential $Credential
}

function Invoke-BMStage {
    <#
    .SYNOPSIS  Submits a staging request for a machine to BuildMaster.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $uri  = '{0}/stage' -f $script:BMBaseUrl
    $body = @{ ComputerName = $ComputerName }
    return Invoke-BMRestCall -Uri $uri -Method POST -Body $body -Credential $Credential
}

# -----------------------------------------------------------------------------
#  BUILDMASTER - STATUS PARSING
# -----------------------------------------------------------------------------

function Get-BMCurrentStage {
    <#
    .SYNOPSIS
        Parses a BuildData object and returns the current stage name string.
    .OUTPUTS  'Staged' | 'Started' | 'OSComplete' | 'Completed' | 'Unknown'
    .DESCRIPTION
        Inspects BuildTimes sorted descending by BuildTime.
        Priority (highest wins): Completed  OSComplete  Started  Staged.
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param([Parameter(Mandatory)][object]$BuildData)

    if ($null -eq $BuildData -or $null -eq $BuildData.BuildTimes) { return 'Unknown' }

    $sorted = @($BuildData.BuildTimes) |
              Sort-Object { [datetime]$_.BuildTime } -Descending

    foreach ($entry in $sorted) {
        if ($entry.IsCompletedTime    -eq $true) { return 'Completed'  }
        if ($entry.IsOSCompletedTime  -eq $true) { return 'OSComplete' }
        if ($entry.IsStartedTime      -eq $true) { return 'Started'    }
        if ($entry.IsStagedTime       -eq $true) { return 'Staged'     }
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
    $flagName = $flagMap[$Stage]

    $match = @($BuildData.BuildTimes) |
             Where-Object { $_.$flagName -eq $true } |
             Sort-Object  { [datetime]$_.BuildTime } |
             Select-Object -First 1

    if ($null -eq $match) { return $null }
    return [datetime]$match.BuildTime
}

# -----------------------------------------------------------------------------
#  VIRTUALWORKS
# -----------------------------------------------------------------------------

function Get-VWDesktop {
    <#
    .SYNOPSIS
        Queries VirtualWorks for a desktop VM record.
        Returns $null (not an error) when the machine is not found (HTTP 404).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    try {
        $uri = '{0}/{1}' -f $script:VWBaseUrl, [Uri]::EscapeDataString($ComputerName)
        return Invoke-BMRestCall -Uri $uri -Method GET -Credential $Credential
    }
    catch {
        if ($_ -match 'HTTP 404') { return $null }
        throw
    }
}

function Invoke-VWReboot {
    <#
    .SYNOPSIS  Sends a reboot command for a DiDC / VM through VirtualWorks.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $uri = '{0}/{1}/reboot' -f $script:VWBaseUrl, [Uri]::EscapeDataString($ComputerName)
    return Invoke-BMRestCall -Uri $uri -Method POST -Credential $Credential
}

function Test-IsVirtualMachine {
    <#
    .SYNOPSIS  Returns $true if the machine exists in VirtualWorks (DiDC / VM).
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [System.Management.Automation.PSCredential]$Credential
    )
    $record = Get-VWDesktop -ComputerName $ComputerName -Credential $Credential
    return ($null -ne $record)
}

# -----------------------------------------------------------------------------
#  EXPORTS
# -----------------------------------------------------------------------------
Export-ModuleMember -Function @(
    'Get-BMBuildData',
    'Get-BMLastBuildInstance',
    'Invoke-BMStage',
    'Get-BMCurrentStage',
    'Get-BMStageEntryTime',
    'Get-VWDesktop',
    'Invoke-VWReboot',
    'Test-IsVirtualMachine'
)
