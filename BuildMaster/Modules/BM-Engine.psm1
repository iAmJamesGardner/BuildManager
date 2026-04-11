#Requires -Version 5.1
<#
.SYNOPSIS
    BM-Engine - Build orchestration state machine
.DESCRIPTION
    Manages a collection of MachineJob objects through the full rebuild lifecycle.
    A WPF DispatcherTimer drives the engine on the UI thread, polling BuildMaster
    and advancing each job's state.

    STATE MACHINE:
        Scheduled   (optional entry point when a future run time is set)
          +-> [scheduled time reached] -> Pending
        Pending
          +-> Staging         (Invoke-BMStage API call issued)
                +-> StagingWait   (success; waiting 30-45 min)
                      +-> Rebooting     (wait elapsed; reboot command sent)
                            +-> Monitoring    (reboot confirmed; polling BuildMaster)
                                  +-> [BuildStage = Staged]     ( 1 hr limit)
                                  +-> [BuildStage = Started]    ( 45 min limit)
                                  +-> [BuildStage = OSComplete] ( 5 hr limit)
                                  +-> Completed
        (any stage failure)
          +-> [RetryCount < 3]  back to Pending (re-stage)
          +-> [RetryCount  3]  Failed

    TIMEOUT RULES (per spec):
        Staged     stage   must advance within 1 hour
        Started    stage   must advance within 45 minutes
        OSComplete stage   must advance within 5 hours
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
#  CONSTANTS
# -----------------------------------------------------------------------------
$script:MaxRetries          = 3
$script:StagingWaitMinutes  = 35   # Default: 35 min (configurable via Set-BMEngineConfig)
$script:PollIntervalSeconds = 30   # How often the engine ticks

$script:StageTimeouts = [ordered]@{
    'Staged'     = [timespan]::FromHours(1)
    'Started'    = [timespan]::FromMinutes(45)
    'OSComplete' = [timespan]::FromHours(5)
}

# -----------------------------------------------------------------------------
#  JOB COLLECTION  (ObservableCollection so WPF ItemsSource gets add/remove events)
# -----------------------------------------------------------------------------
$script:Jobs = New-Object 'System.Collections.ObjectModel.ObservableCollection[Object]'

# Engine timer (WPF DispatcherTimer - runs on UI thread)
$script:EngineTimer = $null

# -----------------------------------------------------------------------------
#  BACKGROUND STAGING WORKER
#  Staging API calls run in a separate runspace so the UI thread stays responsive.
#  The worker modifies job PSCustomObject properties directly (local runspace
#  passes objects by reference - no serialisation).  Log output is marshalled
#  back to the UI thread via a ConcurrentQueue drained on every engine tick.
# -----------------------------------------------------------------------------
$script:BgPS          = $null   # [powershell] instance
$script:BgRunspace    = $null   # [Runspace] for the worker
$script:BgAsyncResult = $null   # IAsyncResult from BeginInvoke
$script:BgModulePath  = ''      # Path to BM-API.psm1 (set via Set-BMEngineConfig)
$script:BgLogQueue    = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# Scriptblock executed in the background runspace for every staging batch.
# Parameters are passed by name via PowerShell.AddParameter(); objects are
# passed by reference because it is a LOCAL (not remote) runspace.
$script:BgStagingScript = {
    param(
        [object[]]$PendingJobs,
        [System.Management.Automation.PSCredential]$Credential,
        [string]$ApiModulePath,
        [string]$BMBaseUrl,
        [string]$VWBaseUrl,
        [string]$VWFqdnSuffix,
        [int]$StagingWaitMinutes,
        [System.Collections.Concurrent.ConcurrentQueue[string]]$LogQueue
    )

    Import-Module $ApiModulePath -Force -DisableNameChecking
    Set-BMAPIConfig -BMBaseUrl $BMBaseUrl -VWBaseUrl $VWBaseUrl -VWFqdnSuffix $VWFqdnSuffix

    foreach ($job in $PendingJobs) {
        if ($job.CancelRequested) { continue }

        $job.Status     = 'Staging'
        $job.Message    = 'Checking build status...'
        $job.StatusIcon = [char]0x25B6

        try {
            # ------------------------------------------------------------------
            # Step 1: Build-status pre-check
            # If BuildMaster already reports the machine as Staged, skip the
            # stage API call entirely and jump straight to StagingWait.
            # ------------------------------------------------------------------
            $currentStatus = $null
            try {
                $currentStatus = Get-BMBuildStatus -ComputerName $job.MachineName -Credential $Credential
            }
            catch {
                $LogQueue.Enqueue(('DEBUG|{0}|Build status pre-check failed (will stage normally): {1}' -f $job.MachineName, $_.Exception.Message))
            }

            $alreadyStaged = (-not [string]::IsNullOrWhiteSpace($currentStatus) -and
                              $currentStatus -ieq 'Staged')

            # ------------------------------------------------------------------
            # Step 2: VM detection  (needed for the reboot step either way)
            # ------------------------------------------------------------------
            $job.Message = 'Detecting machine type (VirtualWorks lookup)...'
            $LogQueue.Enqueue(('DEBUG|{0}|Checking VirtualWorks for machine type' -f $job.MachineName))
            try {
                $job.IsVM = Test-IsVirtualMachine -ComputerName $job.MachineName -Credential $Credential
            }
            catch {
                $job.IsVM = $false
                $LogQueue.Enqueue(('WARNING|{0}|VW check error (defaulting to Physical): {1}' -f $job.MachineName, $_.Exception.Message))
            }
            $job.MachineType = if ($job.IsVM) { 'VM/DiDC' } else { 'Physical' }
            $LogQueue.Enqueue(('INFO|{0}|Machine type resolved: {1} (IsVM={2})' -f $job.MachineName, $job.MachineType, $job.IsVM))

            if ($alreadyStaged) {
                # Already staged - skip the API call, go straight to wait
                if ($job.CancelRequested) { continue }
                $job.StagedAt   = [datetime]::Now
                $job.Status     = 'StagingWait'
                $job.StatusIcon = [char]0x23F0
                $job.Message    = ('Already staged in BuildMaster. Waiting {0} min before reboot...' -f $StagingWaitMinutes)
                $LogQueue.Enqueue(('INFO|{0}|Build status is Staged - skipping stage API call. Waiting {1} min.' -f $job.MachineName, $StagingWaitMinutes))
            }
            else {
                # ------------------------------------------------------------------
                # Step 3: Stage the machine
                # ------------------------------------------------------------------
                $job.Message = 'Contacting BuildMaster to stage machine...'
                $stageResult = Invoke-BMStage -ComputerName $job.MachineName -Credential $Credential

                if ($job.CancelRequested) { continue }

                $job.StagedAt   = [datetime]::Now
                $job.Status     = 'StagingWait'
                $job.StatusIcon = [char]0x23F0

                if ($stageResult.AlreadyStarted) {
                    $job.Message = ('Build already in progress. Waiting {0} min before reboot...' -f $StagingWaitMinutes)
                    $LogQueue.Enqueue(('WARNING|{0}|BuildMaster: build already started - treating as staged. Waiting {1} min.' -f $job.MachineName, $StagingWaitMinutes))
                }
                else {
                    $job.Message = ('Staged. Waiting {0} min before reboot...' -f $StagingWaitMinutes)
                    $LogQueue.Enqueue(('INFO|{0}|Machine staged successfully. Waiting {1} min.' -f $job.MachineName, $StagingWaitMinutes))
                }

                # ------------------------------------------------------------------
                # Step 4: Pre-fetch BuildId (non-fatal)
                # ------------------------------------------------------------------
                try {
                    $buildData = Get-BMBuildData -ComputerName $job.MachineName -Credential $Credential
                    if ($null -ne $buildData -and -not [string]::IsNullOrEmpty($buildData.BuildId)) {
                        $job.BuildId    = $buildData.BuildId
                        $job.InstanceId = if ($null -ne $buildData.InstanceId) { $buildData.InstanceId } else { '' }
                        $LogQueue.Enqueue(('DEBUG|{0}|BuildId acquired: {1}' -f $job.MachineName, $job.BuildId))
                    }
                }
                catch {
                    $LogQueue.Enqueue(('DEBUG|{0}|Could not pre-fetch BuildId: {1}' -f $job.MachineName, $_.Exception.Message))
                }
            }
        }
        catch {
            if (-not $job.CancelRequested) {
                $job.LastError  = $_.Exception.Message
                $job.Message    = ('Stage failed: {0}' -f $_.Exception.Message)
                $job.Status     = 'Error'
                $job.StatusIcon = [char]0x26A0
                $LogQueue.Enqueue(('ERROR|{0}|Stage API failed: {1}' -f $job.MachineName, $_.Exception.Message))
            }
        }
    }
}

# -----------------------------------------------------------------------------
#  LOGGING  (shared with GUI - GUI module registers a UI sink)
# -----------------------------------------------------------------------------
$script:LogPath    = $null
$script:UILogSink  = $null   # [scriptblock] - called by Write-BMLog for GUI output

function Set-BMLogPath {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $script:LogPath = $Path
}

function Register-BMUILogSink {
    <#  .SYNOPSIS  GUI calls this to receive log lines in the log textbox.  #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][scriptblock]$Sink)
    $script:UILogSink = $Sink
}

function Write-BMLog {
    [CmdletBinding()]
    param(
        [string]$MachineName = 'SYSTEM',
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warning','Error','Debug')][string]$Level = 'Info'
    )

    $ts   = (Get-Date).ToString('HH:mm:ss')
    $line = ('[{0}] [{1,-7}] [{2,-15}] {3}' -f $ts, $Level.ToUpper(), $MachineName.ToUpper(), $Message)

    # Console
    $fgColor = switch ($Level) {
        'Error'   { 'Red'     }
        'Warning' { 'Yellow'  }
        'Debug'   { 'DarkGray'}
        default   { 'Cyan'    }
    }
    Write-Host $line -ForegroundColor $fgColor

    # File
    if ($script:LogPath) {
        $logFile = Join-Path $script:LogPath ("BuildMaster_{0}.log" -f (Get-Date -Format 'yyyyMMdd'))
        Add-Content -Path $logFile -Value $line -ErrorAction SilentlyContinue
    }

    # GUI sink
    if ($null -ne $script:UILogSink) {
        try { & $script:UILogSink $line } catch { }
    }
}

# -----------------------------------------------------------------------------
#  JOB FACTORY
# -----------------------------------------------------------------------------

function New-BMJobObject {
    param([Parameter(Mandatory)][string]$MachineName)

    return [PSCustomObject][ordered]@{
        MachineName     = $MachineName.Trim().ToUpper()
        # -- State ----------------------------------------------------------
        Status          = 'Pending'      # Engine state (string for simple binding)
        BuildStage      = 'N/A'          # Last BuildMaster stage seen
        # -- Timestamps ----------------------------------------------------
        StartedAt       = [datetime]::Now
        StagedAt        = $null          # When Invoke-BMStage completed
        StageEnteredAt  = $null          # When current BuildStage was entered
        CompletedAt     = $null
        # -- BuildMaster data ----------------------------------------------
        BuildId         = ''
        InstanceId      = ''
        # -- Metadata ------------------------------------------------------
        IsVM            = $null          # $null=unknown, $true=VM/DiDC, $false=Physical
        RetryCount      = 0
        CancelRequested = $false
        # -- Schedule ------------------------------------------------------
        ScheduledFor    = $null          # [datetime] or $null = run immediately
        ScheduleLabel   = 'Now'          # Display string shown in grid column
        # -- Display -------------------------------------------------------
        ProgressValue   = 0
        Message         = 'Queued'
        StatusIcon      = [char]0x23F3   # [wait]
        MachineType     = 'Unknown'
        ElapsedDisplay  = '00:00:00'
        LastError       = ''
    }
}

# -----------------------------------------------------------------------------
#  PUBLIC JOB MANAGEMENT
# -----------------------------------------------------------------------------

function Add-BMJob {
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory)][string]$MachineName,
        [nullable[datetime]]$ScheduledFor = $null
    )

    $name = $MachineName.Trim().ToUpper()
    $existing = $script:Jobs | Where-Object { $_.MachineName -eq $name }
    if ($null -ne $existing) {
        Write-BMLog -MachineName $name -Message "Job already exists (Status: $($existing.Status))" -Level Warning
        return $existing
    }
    $job = New-BMJobObject -MachineName $name

    # If a future time is specified, park the job in Scheduled state
    if ($null -ne $ScheduledFor -and $ScheduledFor -gt [datetime]::Now) {
        $job.ScheduledFor  = $ScheduledFor
        $job.Status        = 'Scheduled'
        $job.ScheduleLabel = $ScheduledFor.ToString('MM/dd HH:mm')
        $job.StatusIcon    = [char]0x23F0   # [alarm]
        $job.Message       = "Scheduled for $($ScheduledFor.ToString('MM/dd/yyyy HH:mm'))"
        Write-BMLog -MachineName $name -Message ("Job scheduled for {0}" -f $ScheduledFor.ToString('MM/dd/yyyy HH:mm')) -Level Info
    } else {
        Write-BMLog -MachineName $name -Message 'Job created and queued' -Level Info
    }

    $script:Jobs.Add($job)
    return $job
}

function Stop-BMJob {
    <#  .SYNOPSIS  Requests cancellation of a specific job.  #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$MachineName)
    $job = Get-BMJob -MachineName $MachineName
    if ($null -ne $job) {
        $job.CancelRequested = $true
        $job.Status          = 'Cancelled'
        $job.Message         = 'Cancelled by user'
        $job.StatusIcon      = [char]0x274C   # [X]
        Write-BMLog -MachineName $MachineName -Message 'Job cancelled by user' -Level Warning
    }
}

function Remove-BMJob {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$MachineName)
    $job = Get-BMJob -MachineName $MachineName
    if ($null -ne $job) {
        $script:Jobs.Remove($job) | Out-Null
    }
}

function Get-BMJob {
    [CmdletBinding()]
    param([string]$MachineName)
    if ($MachineName) {
        return $script:Jobs | Where-Object { $_.MachineName -eq $MachineName.Trim().ToUpper() }
    }
    return @($script:Jobs)
}

function Get-BMJobs {
    # Returns job objects as an enumerable sequence (for filtering, counting, piping).
    # Do NOT use this for WPF ItemsSource - use Get-BMJobCollection instead.
    $script:Jobs
}

function Get-BMJobCollection {
    # Returns the ObservableCollection[Object] reference itself, NOT its contents.
    #
    # PowerShell enumerates IEnumerable objects when they pass through the output
    # pipeline, so  $ctrl.ItemsSource = Get-BMJobs  sets ItemsSource to $null
    # (empty collection) or a plain object[] (non-empty) — neither fires
    # CollectionChanged when items are later added.
    #
    # The leading comma operator wraps $script:Jobs in a 1-element array.
    # The pipeline enumerates that wrapper array (yielding the collection object),
    # but does NOT recurse into the collection itself, so the caller receives the
    # live ObservableCollection reference.
    ,$script:Jobs
}

function Clear-BMCompletedJobs {
    $toRemove = @($script:Jobs | Where-Object { $_.Status -in @('Completed','Failed','Cancelled') })
    foreach ($j in $toRemove) { $script:Jobs.Remove($j) | Out-Null }
}

function Stop-AllBMJobs {
    <#
    .SYNOPSIS  Cancels every non-terminal job in the queue.
    .OUTPUTS   [int] number of jobs that were cancelled.
    .DESCRIPTION
        Jobs already dispatched to BuildMaster will continue on the server;
        only this tool's local tracking is stopped.
    #>
    [CmdletBinding()]
    [OutputType([int])]
    param()

    $active = @($script:Jobs | Where-Object {
        $_.Status -notin @('Completed','Failed','Cancelled')
    })

    foreach ($j in $active) {
        $j.CancelRequested = $true
        $j.Status          = 'Cancelled'
        $j.Message         = 'Cancelled by user (Stop All Builds)'
        $j.StatusIcon      = [char]0x274C   # [X]
        Write-BMLog -MachineName $j.MachineName -Message 'Job cancelled (Stop All Builds)' -Level Warning
    }

    return $active.Count
}

# -----------------------------------------------------------------------------
#  ENGINE LIFECYCLE
# -----------------------------------------------------------------------------

function Set-BMEngineConfig {
    [CmdletBinding()]
    param(
        [int]$StagingWaitMinutes,
        [int]$PollIntervalSeconds,
        [string]$ModulePath          # Path to BM-API.psm1 for the background staging runspace
    )
    if ($PSBoundParameters.ContainsKey('StagingWaitMinutes'))  { $script:StagingWaitMinutes  = $StagingWaitMinutes  }
    if ($PSBoundParameters.ContainsKey('PollIntervalSeconds')) { $script:PollIntervalSeconds = $PollIntervalSeconds }
    if ($PSBoundParameters.ContainsKey('ModulePath'))          { $script:BgModulePath        = $ModulePath          }

    # Update live timer interval if running
    if ($null -ne $script:EngineTimer -and $script:EngineTimer.IsEnabled) {
        $script:EngineTimer.Interval = [timespan]::FromSeconds($script:PollIntervalSeconds)
    }
}

function Request-BMQuickTick {
    <#
    .SYNOPSIS
        Shortens the DispatcherTimer to 500 ms so the first engine tick fires
        almost immediately after the current UI event returns.
        Invoke-BMEngineTick restores the normal interval on the next tick.
        Call this instead of Invoke-BMEngineTick from UI event handlers to keep
        the UI thread responsive during API calls.
    #>
    if ($null -ne $script:EngineTimer -and $script:EngineTimer.IsEnabled) {
        $script:EngineTimer.Stop()
        $script:EngineTimer.Interval = [timespan]::FromMilliseconds(500)
        $script:EngineTimer.Start()
    }
}

function Start-BMStagingBackground {
    <#
    .SYNOPSIS
        Submits all currently-Pending jobs to a background runspace for staging.
        Returns immediately; the runspace modifies job properties asynchronously.
        The DispatcherTimer tick drains the log queue and collects the result.
    #>

    # If a run is already in-flight, do nothing - pending jobs will be picked
    # up on the next call once the current run completes.
    if ($null -ne $script:BgPS -and -not $script:BgAsyncResult.IsCompleted) { return }

    $pendingJobs = @($script:Jobs | Where-Object {
        $_.Status -eq 'Pending' -and -not $_.CancelRequested
    })
    if ($pendingJobs.Count -eq 0) { return }

    # Clean up any previously completed (but not yet collected) runner
    if ($null -ne $script:BgPS) {
        try   { $script:BgPS.EndInvoke($script:BgAsyncResult) } catch { }
        $script:BgPS.Dispose()
        $script:BgPS = $null
        $script:BgRunspace.Close()
        $script:BgRunspace.Dispose()
        $script:BgRunspace    = $null
        $script:BgAsyncResult = $null
    }

    # Fallback to inline staging if the API module path is not configured
    if ([string]::IsNullOrEmpty($script:BgModulePath) -or
        -not (Test-Path $script:BgModulePath)) {
        Write-BMLog -Message ('Background staging unavailable: BM-API path not configured ({0}). Falling back to inline staging.' -f $script:BgModulePath) -Level Warning
        foreach ($job in $pendingJobs) {
            Invoke-BMStep_Stage -Job $job
        }
        return
    }

    $regCred = Get-BMRegularCredential
    $apiCfg  = Get-BMAPIConfig

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'MTA'
    $rs.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $rs
    [void]$ps.AddScript($script:BgStagingScript)
    [void]$ps.AddParameter('PendingJobs',        $pendingJobs)
    [void]$ps.AddParameter('Credential',         $regCred)
    [void]$ps.AddParameter('ApiModulePath',      $script:BgModulePath)
    [void]$ps.AddParameter('BMBaseUrl',          $apiCfg.BMBaseUrl)
    [void]$ps.AddParameter('VWBaseUrl',          $apiCfg.VWBaseUrl)
    [void]$ps.AddParameter('VWFqdnSuffix',       $apiCfg.VWFqdnSuffix)
    [void]$ps.AddParameter('StagingWaitMinutes', $script:StagingWaitMinutes)
    [void]$ps.AddParameter('LogQueue',           $script:BgLogQueue)

    $script:BgRunspace    = $rs
    $script:BgPS          = $ps
    $script:BgAsyncResult = $ps.BeginInvoke()

    Write-BMLog -Message ("Background staging started for {0} machine(s)" -f $pendingJobs.Count) -Level Debug
}

function Start-BMEngine {
    [CmdletBinding()]
    param()

    if ($null -ne $script:EngineTimer -and $script:EngineTimer.IsEnabled) {
        Write-BMLog -Message 'Engine already running' -Level Warning
        return
    }

    $script:EngineTimer          = New-Object System.Windows.Threading.DispatcherTimer
    $script:EngineTimer.Interval = [timespan]::FromSeconds($script:PollIntervalSeconds)
    $script:EngineTimer.Add_Tick({
        try { Invoke-BMEngineTick }
        catch {
            Write-BMLog -Message ("Engine tick unhandled error: {0}" -f $_.Exception.Message) -Level Error
        }
    })
    $script:EngineTimer.Start()

    Write-BMLog -Message ("Engine started - poll interval: {0}s  staging wait: {1}m" -f
                           $script:PollIntervalSeconds, $script:StagingWaitMinutes) -Level Info
}

function Stop-BMEngine {
    [CmdletBinding()]
    param()
    if ($null -ne $script:EngineTimer) {
        $script:EngineTimer.Stop()
        $script:EngineTimer = $null
        Write-BMLog -Message 'Engine stopped' -Level Info
    }
}

function Test-BMEngineRunning {
    return ($null -ne $script:EngineTimer -and $script:EngineTimer.IsEnabled)
}

# -----------------------------------------------------------------------------
#  ENGINE TICK  (called every PollInterval by DispatcherTimer)
# -----------------------------------------------------------------------------

function Invoke-BMEngineTick {
    # -- Restore normal poll interval if a quick-tick shortened it -------------
    if ($null -ne $script:EngineTimer -and
        $script:EngineTimer.Interval.TotalSeconds -lt $script:PollIntervalSeconds) {
        $script:EngineTimer.Interval = [timespan]::FromSeconds($script:PollIntervalSeconds)
    }

    # -- Drain background staging log queue ------------------------------------
    $logEntry = $null
    while ($script:BgLogQueue.TryDequeue([ref]$logEntry)) {
        $parts = $logEntry -split '\|', 3
        if ($parts.Count -eq 3) {
            Write-BMLog -Level $parts[0] -MachineName $parts[1] -Message $parts[2]
        }
    }

    # -- Collect completed background worker -----------------------------------
    if ($null -ne $script:BgPS -and $script:BgAsyncResult.IsCompleted) {
        try   { $script:BgPS.EndInvoke($script:BgAsyncResult) }
        catch { Write-BMLog -Message ("Background staging runner error: {0}" -f $_.Exception.Message) -Level Error }
        $script:BgPS.Dispose()
        $script:BgPS = $null
        $script:BgRunspace.Close()
        $script:BgRunspace.Dispose()
        $script:BgRunspace    = $null
        $script:BgAsyncResult = $null
        Write-BMLog -Message 'Background staging batch complete' -Level Debug
    }

    $activeStatuses = @('Pending','Staging','StagingWait','Rebooting','Monitoring','Error','Scheduled')

    $activeJobs = @($script:Jobs | Where-Object {
        $_.Status -in $activeStatuses -and -not $_.CancelRequested
    })

    foreach ($job in $activeJobs) {
        # Update elapsed display (also updates in Completed/Failed rows)
        Update-BMJobElapsed -Job $job

        try {
            Invoke-BMJobStep -Job $job
        }
        catch {
            $job.LastError = $_.Exception.Message
            $job.Message   = "Unexpected error: $($_.Exception.Message)"
            $job.Status    = 'Error'
            $job.StatusIcon = [char]0x26A0  # [!]
            Write-BMLog -MachineName $job.MachineName `
                        -Message "Unhandled error: $($_.Exception.Message)" `
                        -Level Error
            Invoke-BMHandleFailure -Job $job
        }
    }

    # Also refresh elapsed on non-active jobs (for display)
    $script:Jobs | Where-Object { $_.Status -notin $activeStatuses } | ForEach-Object {
        Update-BMJobElapsed -Job $_
    }
}

function Get-BMJobProgress {
    <#
    .SYNOPSIS  Returns a 0-100 integer progress value for the job's current state.
    .DESCRIPTION
        Terminal states (Failed/Cancelled/Error) return the current ProgressValue
        unchanged so the bar freezes at the point of failure.
    #>
    param([object]$Job)
    switch ($Job.Status) {
        'Scheduled'   { return 0  }
        'Pending'     { return 2  }
        'Staging'     { return 8  }
        'StagingWait' {
            if ($null -ne $Job.StagedAt) {
                $pct = ([datetime]::Now - $Job.StagedAt).TotalMinutes / [math]::Max(1, $script:StagingWaitMinutes)
                $pct = [math]::Min($pct, 1.0)
                return [int](10 + $pct * 20)   # 10 -> 30 %
            }
            return 10
        }
        'Rebooting'   { return 32 }
        'Monitoring'  {
            switch ($Job.BuildStage) {
                'Waiting'    { return 35 }
                'Staged'     { return 42 }
                'Started'    { return 58 }
                'OSComplete' { return 82 }
                default      { return 38 }
            }
        }
        'Completed'   { return 100 }
        # Terminal states - freeze at wherever the bar was
        'Failed'      { return $Job.ProgressValue }
        'Cancelled'   { return $Job.ProgressValue }
        'Error'       { return $Job.ProgressValue }
        default       { return 0 }
    }
}

function Update-BMJobElapsed {
    param([object]$Job)
    if ($null -ne $Job.StartedAt) {
        $e = [datetime]::Now - $Job.StartedAt
        $Job.ElapsedDisplay = ('{0:hh\:mm\:ss}' -f $e)
    }
    $Job.ProgressValue = Get-BMJobProgress -Job $Job
}

# -----------------------------------------------------------------------------
#  STATE MACHINE STEPS
# -----------------------------------------------------------------------------

function Invoke-BMJobStep {
    param([Parameter(Mandatory)][object]$Job)

    switch ($Job.Status) {
        'Scheduled'   { Invoke-BMStep_Scheduled   -Job $Job }
        'Pending'     { Start-BMStagingBackground }   # submits batch to bg runspace; no-op if already in-flight
        'Staging'     { <# in-flight: background runspace is working #> }
        'StagingWait' { Invoke-BMStep_StagingWait  -Job $Job }
        'Rebooting'   { Invoke-BMStep_Reboot        -Job $Job }
        'Monitoring'  { Invoke-BMStep_Monitor        -Job $Job }
        'Error'       {
            # Auto-retry after error unless max retries exceeded
            if ($Job.RetryCount -lt $script:MaxRetries) {
                Write-BMLog -MachineName $Job.MachineName `
                            -Message ("Auto-retry after error (attempt {0}/{1})" -f ($Job.RetryCount+1), $script:MaxRetries) `
                            -Level Warning
                Invoke-BMResetForRetry -Job $Job
            }
        }
    }
}

# -- Step: Scheduled -----------------------------------------------------------
function Invoke-BMStep_Scheduled {
    param([object]$Job)

    if ($null -eq $Job.ScheduledFor) {
        # No time set - just release immediately
        $Job.Status        = 'Pending'
        $Job.ScheduleLabel = 'Now'
        return
    }

    $remaining = $Job.ScheduledFor - [datetime]::Now

    if ($remaining.TotalSeconds -le 0) {
        $Job.Status        = 'Pending'
        $Job.ScheduleLabel = 'Now'
        $Job.StatusIcon    = [char]0x23F3   # [wait]
        $Job.Message       = 'Scheduled time reached - starting rebuild...'
        Write-BMLog -MachineName $Job.MachineName -Message 'Scheduled time reached - transitioning to Pending' -Level Info
    } else {
        $h = [math]::Floor($remaining.TotalHours)
        $m = $remaining.Minutes
        $s = $remaining.Seconds
        $Job.Message = ('Scheduled - starts in {0:D2}:{1:D2}:{2:D2}' -f $h, $m, $s)
    }
}

# -- Step: Stage (inline fallback - used when background runspace unavailable) -
function Invoke-BMStep_Stage {
    param([object]$Job)
    # Normal path: Start-BMStagingBackground runs this work off the UI thread.
    # This inline version is only reached when BgModulePath is not configured.
    # It blocks the UI thread briefly - prefer the background path.

    $regCred = Get-BMRegularCredential

    $Job.Status     = 'Staging'
    $Job.Message    = 'Staging (inline - background path unavailable)...'
    $Job.StatusIcon = [char]0x25B6

    try {
        # VM detection
        try   { $Job.IsVM = Test-IsVirtualMachine -ComputerName $Job.MachineName -Credential $regCred }
        catch { $Job.IsVM = $false }
        $Job.MachineType = if ($Job.IsVM) { 'VM/DiDC' } else { 'Physical' }
        Write-BMLog -MachineName $Job.MachineName `
                    -Message ("Machine type: {0} (IsVM={1})" -f $Job.MachineType, $Job.IsVM) -Level Info

        # Build-status pre-check
        $currentStatus = $null
        try { $currentStatus = Get-BMBuildStatus -ComputerName $Job.MachineName -Credential $regCred } catch { }

        if (-not [string]::IsNullOrWhiteSpace($currentStatus) -and $currentStatus -ieq 'Staged') {
            $Job.StagedAt   = [datetime]::Now
            $Job.Status     = 'StagingWait'
            $Job.StatusIcon = [char]0x23F0
            $Job.Message    = ("Already staged. Waiting {0} min before reboot..." -f $script:StagingWaitMinutes)
            Write-BMLog -MachineName $Job.MachineName -Message 'Already staged - skipping stage API call.' -Level Info
        }
        else {
            $stageResult    = Invoke-BMStage -ComputerName $Job.MachineName -Credential $regCred
            $Job.StagedAt   = [datetime]::Now
            $Job.Status     = 'StagingWait'
            $Job.StatusIcon = [char]0x23F0
            $msg = if ($stageResult.AlreadyStarted) {
                "Build already in progress. Waiting {0} min before reboot..."
            } else {
                "Staged. Waiting {0} min before reboot..."
            }
            $Job.Message = ($msg -f $script:StagingWaitMinutes)
            Write-BMLog -MachineName $Job.MachineName `
                        -Message ("Staged (inline). Waiting {0} min." -f $script:StagingWaitMinutes) -Level Info
            try {
                $buildData = Get-BMBuildData -ComputerName $Job.MachineName -Credential $regCred
                if ($null -ne $buildData -and $buildData.BuildId) { $Job.BuildId = $buildData.BuildId }
            } catch { }
        }
    }
    catch {
        $Job.LastError  = $_.Exception.Message
        $Job.Message    = "Stage failed: $($_.Exception.Message)"
        $Job.Status     = 'Error'
        $Job.StatusIcon = [char]0x26A0
        Write-BMLog -MachineName $Job.MachineName `
                    -Message ("Stage API failed: {0}" -f $_.Exception.Message) -Level Error
        Invoke-BMHandleFailure -Job $Job
    }
}

# -- Step: Staging Wait --------------------------------------------------------
function Invoke-BMStep_StagingWait {
    param([object]$Job)

    if ($null -eq $Job.StagedAt) { $Job.StagedAt = [datetime]::Now; return }

    $elapsed   = [datetime]::Now - $Job.StagedAt
    $waitSpan  = [timespan]::FromMinutes($script:StagingWaitMinutes)
    $remaining = $waitSpan - $elapsed

    if ($remaining.TotalSeconds -le 0) {
        $Job.Status     = 'Rebooting'
        $Job.Message    = 'Staging wait complete. Sending reboot...'
        $Job.StatusIcon = [char]0x21BB   # (o>) clockwise arrow - rebooting
        Write-BMLog -MachineName $Job.MachineName -Message 'Staging wait period complete. Initiating reboot.' -Level Info
    }
    else {
        $mins = [math]::Ceiling($remaining.TotalMinutes)
        $Job.Message = ("Staging wait - ~{0} min remaining before reboot..." -f $mins)
    }
}

# -- Step: Reboot --------------------------------------------------------------
function Invoke-BMStep_Reboot {
    param([object]$Job)

    $regCred = Get-BMRegularCredential  # $null for regular sessions (SSO)

    try {
        if ($Job.IsVM -eq $true) {
            # -- DiDC / VM: VirtualWorks API reboot (regular/SSO credentials) --
            Write-BMLog -MachineName $Job.MachineName -Message 'Rebooting via VirtualWorks API' -Level Info
            Invoke-VWReboot -ComputerName $Job.MachineName -Credential $regCred | Out-Null
        }
        else {
            # -- Physical: requires elevation (direct priv session or EPM) ----
            $ctx = Get-BMSessionContext

            # If session has no elevation, fail this job immediately.
            # No point retrying - it will never succeed without a different session.
            # VM/DiDC jobs in the same batch are unaffected (they took the branch above).
            if ($ctx.Blocked -and -not (Get-BMWhatIfMode)) {
                $failMsg = ("Physical machine reboot requires elevation. " +
                            "Session '{0}' has no CyberArk EPM or PRIVDOMAIN rights. " +
                            "Relaunch the tool as a PRIVDOMAIN account to rebuild this machine.") -f $ctx.FullUser
                $Job.Status     = 'Failed'
                $Job.StatusIcon = [char]0x274C
                $Job.Message    = $failMsg
                $Job.LastError  = $failMsg
                $Job.RetryCount = $script:MaxRetries   # prevent any retry loop
                Write-BMLog -MachineName $Job.MachineName -Message $failMsg -Level Error
                return
            }

            Write-BMLog -MachineName $Job.MachineName `
                        -Message ("Rebooting physical machine [Session: {0}  Elevation: {1}]" -f `
                                  $ctx.FullUser, $ctx.ElevationMethod) `
                        -Level Info

            Invoke-BMPrivilegedCommand -TargetComputer $Job.MachineName -ScriptBlock {
                Restart-Computer -Force -ErrorAction Stop
            }
        }

        $Job.Status         = 'Monitoring'
        $Job.BuildStage     = 'Waiting'
        $Job.StageEnteredAt = [datetime]::Now
        $Job.StatusIcon     = [char]0x25CF   # (*) black circle - monitoring
        $Job.Message        = 'Reboot sent. Monitoring BuildMaster...'

        Write-BMLog -MachineName $Job.MachineName -Message 'Reboot command sent successfully. Entering monitoring.' -Level Info
    }
    catch {
        $Job.LastError = $_.Exception.Message
        $Job.Message   = "Reboot failed: $($_.Exception.Message)"
        $Job.Status    = 'Error'
        $Job.StatusIcon = [char]0x26A0
        Write-BMLog -MachineName $Job.MachineName -Message "Reboot failed: $($_.Exception.Message)" -Level Error
        Invoke-BMHandleFailure -Job $Job
    }
}

# -- Step: Monitor -------------------------------------------------------------
function Invoke-BMStep_Monitor {
    param([object]$Job)

    # -- Timeout check for current BuildStage ---------------------------------
    if ($Job.BuildStage -in $script:StageTimeouts.Keys -and $null -ne $Job.StageEnteredAt) {
        $elapsed = [datetime]::Now - $Job.StageEnteredAt
        $limit   = $script:StageTimeouts[$Job.BuildStage]
        if ($elapsed -gt $limit) {
            $msg = ("Timeout in stage '{0}' ({1:hh\:mm\:ss} elapsed, limit {2:hh\:mm\:ss})" -f
                    $Job.BuildStage, $elapsed, $limit)
            Write-BMLog -MachineName $Job.MachineName -Message $msg -Level Warning
            $Job.LastError  = $msg
            $Job.Status     = 'Error'
            $Job.StatusIcon = [char]0x23F1  # [timer]
            Invoke-BMHandleFailure -Job $Job
            return
        }
    }

    # -- Poll BuildMaster (two-step: get BuildId first, then instance) --------
    try {
        $regCred = Get-BMRegularCredential   # $null for regular sessions (SSO)

        # Step 1: If we don't have the BuildId yet, query by computer name
        if ([string]::IsNullOrEmpty($Job.BuildId)) {
            $Job.Message = 'Waiting for BuildMaster build record...'
            $buildRef = Get-BMBuildData -ComputerName $Job.MachineName -Credential $regCred

            if ($null -eq $buildRef -or [string]::IsNullOrEmpty($buildRef.BuildId)) {
                # Build record not available yet; machine may still be rebooting
                $Job.Message = 'Monitoring - build record not yet available...'
                return
            }

            $Job.BuildId    = $buildRef.BuildId
            $Job.InstanceId = if ($null -ne $buildRef.InstanceId) { $buildRef.InstanceId } else { '' }

            Write-BMLog -MachineName $Job.MachineName `
                        -Message ("BuildId acquired during monitoring: {0}" -f $Job.BuildId) -Level Debug
        }

        # Step 2: Get instance data using the BuildId
        $instanceData = Get-BMBuildInstance -BuildId $Job.BuildId -Credential $regCred
        $newStage     = Get-BMCurrentStage -BuildData $instanceData

        # -- Advance state on stage transition --------------------------------
        if ($newStage -ne $Job.BuildStage -and $newStage -ne 'Unknown') {
            $prev               = $Job.BuildStage
            $Job.BuildStage     = $newStage
            $Job.StageEnteredAt = Get-BMStageEntryTime -BuildData $instanceData -Stage $newStage

            Write-BMLog -MachineName $Job.MachineName `
                        -Message ("Stage advanced: {0} -> {1}" -f $prev, $newStage) `
                        -Level Info
        }

        # -- Terminal / display state -----------------------------------------
        switch ($newStage) {
            'Completed' {
                $Job.Status      = 'Completed'
                $Job.CompletedAt = [datetime]::Now
                $Job.StatusIcon  = [char]0x2705   # [OK]
                $Job.Message     = 'Build completed successfully!'
                Write-BMLog -MachineName $Job.MachineName -Message 'BUILD COMPLETED [+]' -Level Info
                return
            }
            'OSComplete' {
                $Job.Message    = 'OS complete - applying post-build configuration...'
                $Job.StatusIcon = [char]0x25A0    # black square
            }
            'Started' {
                $Job.Message    = 'Build started - OS installing...'
                $Job.StatusIcon = [char]0x25B2    # triangle up
            }
            'Staged' {
                $Job.Message    = 'Machine staged - waiting for build to start...'
                $Job.StatusIcon = [char]0x25CB    # white circle
            }
            default {
                $Job.Message = ("Monitoring... (last stage seen: {0})" -f $newStage)
            }
        }
    }
    catch {
        # Non-fatal: machine may still be rebooting; log and continue
        $Job.Message = ("Poll error (will retry): {0}" -f $_.Exception.Message)
        Write-BMLog -MachineName $Job.MachineName `
                    -Message ("BuildMaster poll error (non-fatal): {0}" -f $_.Exception.Message) `
                    -Level Warning
    }
}

# -----------------------------------------------------------------------------
#  RETRY / FAILURE HANDLING
# -----------------------------------------------------------------------------

function Invoke-BMHandleFailure {
    param([object]$Job)

    $Job.RetryCount++

    if ($Job.RetryCount -ge $script:MaxRetries) {
        $Job.Status     = 'Failed'
        $Job.StatusIcon = [char]0x274C  # [X]
        $Job.Message    = "FAILED after $script:MaxRetries retries. Last: $($Job.LastError)"
        Write-BMLog -MachineName $Job.MachineName `
                    -Message ("Job FAILED after {0} retries." -f $script:MaxRetries) `
                    -Level Error
    }
    else {
        Write-BMLog -MachineName $Job.MachineName `
                    -Message ("Scheduling retry {0}/{1}..." -f $Job.RetryCount, $script:MaxRetries) `
                    -Level Warning
        Invoke-BMResetForRetry -Job $Job
    }
}

function Invoke-BMResetForRetry {
    param([object]$Job)
    $Job.Status         = 'Pending'
    $Job.BuildStage     = 'N/A'
    $Job.StageEnteredAt = $null
    $Job.StagedAt       = $null
    $Job.StatusIcon     = [char]0x21BA   # (<o) counterclockwise arrow - retry
    $Job.Message        = ("Retry {0}/{1} - re-staging..." -f $Job.RetryCount, $script:MaxRetries)
}

# -----------------------------------------------------------------------------
#  EXPORTS
# -----------------------------------------------------------------------------
Export-ModuleMember -Function @(
    # Logging
    'Set-BMLogPath',
    'Register-BMUILogSink',
    'Write-BMLog',
    # Config
    'Set-BMEngineConfig',
    # Job management
    'Add-BMJob',
    'Stop-BMJob',
    'Stop-AllBMJobs',
    'Remove-BMJob',
    'Get-BMJob',
    'Get-BMJobs',
    'Get-BMJobCollection',
    'Clear-BMCompletedJobs',
    # Engine
    'Start-BMEngine',
    'Stop-BMEngine',
    'Test-BMEngineRunning',
    'Invoke-BMEngineTick',
    'Request-BMQuickTick'
)
