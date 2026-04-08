#Requires -Version 5.1
<#
.SYNOPSIS
    BM-GUI — WPF interface for BuildMaster
.DESCRIPTION
    Defines the full WPF window, wires all controls to engine/API functions,
    and drives a UI-refresh DispatcherTimer for live grid/log updates.

    Designed for PS 5.1 — no MVVM framework, no external dependencies.
    All UI work runs on the WPF Dispatcher thread.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ─────────────────────────────────────────────────────────────────────────────
#  XAML DEFINITION
# ─────────────────────────────────────────────────────────────────────────────
[string]$script:XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="BuildMaster — Machine Rebuild Manager"
    Height="820" Width="1250"
    MinHeight="600" MinWidth="900"
    WindowStartupLocation="CenterScreen"
    Background="#12121F"
    FontFamily="Segoe UI" FontSize="13">

    <Window.Resources>
        <!-- Global control styles for dark theme -->
        <Style TargetType="Button">
            <Setter Property="Background"             Value="#2E2E4A"/>
            <Setter Property="Foreground"             Value="#E0E0FF"/>
            <Setter Property="BorderBrush"            Value="#5050A0"/>
            <Setter Property="BorderThickness"        Value="1"/>
            <Setter Property="Padding"                Value="8,4"/>
            <Setter Property="Cursor"                 Value="Hand"/>
            <Setter Property="SnapsToDevicePixels"    Value="True"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#404070"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#505090"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background"      Value="#1E1E35"/>
            <Setter Property="Foreground"      Value="#D0D0F0"/>
            <Setter Property="BorderBrush"     Value="#404068"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="5,3"/>
            <Setter Property="CaretBrush"      Value="White"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Foreground"      Value="#A0A0D0"/>
            <Setter Property="BorderBrush"     Value="#303058"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding"         Value="6"/>
        </Style>
        <Style TargetType="Label">
            <Setter Property="Foreground" Value="#A0A0C8"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#A0A0C8"/>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background"                Value="#1A1A30"/>
            <Setter Property="Foreground"                Value="#D0D0F0"/>
            <Setter Property="BorderBrush"               Value="#303058"/>
            <Setter Property="GridLinesVisibility"       Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush"  Value="#2A2A48"/>
            <Setter Property="RowBackground"             Value="#1A1A30"/>
            <Setter Property="AlternatingRowBackground"  Value="#1E1E38"/>
            <Setter Property="SelectionMode"             Value="Single"/>
            <Setter Property="SelectionUnit"             Value="FullRow"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background"    Value="#252545"/>
            <Setter Property="Foreground"    Value="#C0C0F0"/>
            <Setter Property="Padding"       Value="8,6"/>
            <Setter Property="FontWeight"    Value="SemiBold"/>
            <Setter Property="BorderBrush"   Value="#404068"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>
        <Style TargetType="DataGridRow">
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Style.Triggers>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="#304080"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="DataGridCell">
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding"         Value="4,0"/>
        </Style>
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#1A1A30"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="62"/>  <!-- Header -->
            <RowDefinition Height="*"/>   <!-- Main content -->
            <RowDefinition Height="30"/>  <!-- Status bar -->
        </Grid.RowDefinitions>

        <!-- ═══════════════════  HEADER  ═══════════════════ -->
        <Border Grid.Row="0" Background="#1A1A35" BorderBrush="#303058" BorderThickness="0,0,0,1">
            <Grid Margin="12,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Left: title + env badge -->
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="⚙" FontSize="26" Foreground="#7070FF"
                               VerticalAlignment="Center" Margin="0,0,8,0"/>
                    <TextBlock Text="BuildMaster" FontSize="22" FontWeight="Bold"
                               Foreground="White" VerticalAlignment="Center"/>
                    <Border x:Name="EnvBadge" CornerRadius="4" Margin="14,0,0,0"
                            Padding="10,3" VerticalAlignment="Center">
                        <TextBlock x:Name="EnvLabel" FontWeight="Bold" FontSize="13"
                                   Foreground="White"/>
                    </Border>
                    <TextBlock x:Name="EngineStatusLabel" Margin="18,0,0,0"
                               FontSize="11" VerticalAlignment="Center" Foreground="#606080"/>
                </StackPanel>

                <!-- Right: cred status -->
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock x:Name="RegCredLabel"  Margin="0,0,16,0" FontSize="11"/>
                    <TextBlock x:Name="PrivCredLabel" Margin="0,0,8,0"  FontSize="11"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ═══════════════════  MAIN CONTENT  ═══════════════════ -->
        <Grid Grid.Row="1" Margin="10,8,10,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="265"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- ──────────────  LEFT PANEL  ────────────── -->
            <StackPanel Grid.Column="0" Margin="0,0,10,0">

                <!-- Credentials -->
                <GroupBox Header=" Credentials " Margin="0,0,0,8">
                    <StackPanel Margin="2">
                        <Button x:Name="BtnSetRegCred"  Content="🔑 Set Regular Account"
                                Height="32" Margin="0,4,0,0"/>
                        <Button x:Name="BtnSetPrivCred" Content="🔐 Set Privileged Account"
                                Height="32" Margin="0,6,0,4"/>
                        <TextBlock x:Name="CredsSummary" TextWrapping="Wrap"
                                   FontSize="10.5" Foreground="#707098" Margin="2,2,0,0"/>
                    </StackPanel>
                </GroupBox>

                <!-- Machine Input -->
                <GroupBox Header=" Machines to Rebuild " Margin="0,0,0,8">
                    <StackPanel Margin="2">
                        <TextBlock Text="One machine name per line:" FontSize="11"
                                   Margin="0,2,0,4" Foreground="#6868A0"/>
                        <TextBox x:Name="TxtMachines"
                                 Height="160"
                                 AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"
                                 TextWrapping="NoWrap"
                                 FontFamily="Consolas" FontSize="12"
                                 VerticalContentAlignment="Top"/>
                        <Button x:Name="BtnStartRebuild"
                                Content="▶  Start Rebuild"
                                Height="36" Margin="0,8,0,0"
                                Background="#1B5E20" Foreground="White"
                                FontWeight="Bold" FontSize="13"
                                BorderBrush="#388E3C"/>
                        <Button x:Name="BtnClearCompleted"
                                Content="Clear Completed / Failed"
                                Height="28" Margin="0,5,0,2"/>
                    </StackPanel>
                </GroupBox>

                <!-- Settings -->
                <GroupBox Header=" Settings ">
                    <Grid Margin="2">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                            <RowDefinition Height="30"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="60"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Grid.Row="0" Grid.Column="0" Text="Staging wait (min):"
                                   VerticalAlignment="Center"/>
                        <TextBox   Grid.Row="0" Grid.Column="1" x:Name="TxtStagingWait"
                                   Text="35" VerticalAlignment="Center" Height="24"/>

                        <TextBlock Grid.Row="1" Grid.Column="0" Text="Poll interval (sec):"
                                   VerticalAlignment="Center"/>
                        <TextBox   Grid.Row="1" Grid.Column="1" x:Name="TxtPollInterval"
                                   Text="30" VerticalAlignment="Center" Height="24"/>

                        <TextBlock Grid.Row="2" Grid.Column="0" Text="Max retries:"
                                   VerticalAlignment="Center"/>
                        <TextBox   Grid.Row="2" Grid.Column="1" x:Name="TxtMaxRetries"
                                   Text="3" VerticalAlignment="Center" Height="24"
                                   IsReadOnly="True" Foreground="#606080"/>

                        <Button    Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="2"
                                   x:Name="BtnApplySettings" Content="Apply Settings"
                                   Height="26" Margin="0,4,0,0"/>
                    </Grid>
                </GroupBox>
            </StackPanel>

            <!-- ──────────────  RIGHT PANEL  ────────────── -->
            <Grid Grid.Column="1">
                <Grid.RowDefinitions>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="160"/>
                </Grid.RowDefinitions>

                <!-- Job DataGrid -->
                <DataGrid x:Name="JobGrid"
                          Grid.Row="0"
                          AutoGenerateColumns="False"
                          CanUserAddRows="False"
                          CanUserDeleteRows="False"
                          CanUserReorderColumns="True"
                          CanUserSortColumns="True"
                          IsReadOnly="True"
                          RowHeight="28"
                          ColumnHeaderHeight="30"
                          Margin="0,0,0,6">

                    <DataGrid.Columns>
                        <!-- Status icon -->
                        <DataGridTextColumn Header=""
                                            Binding="{Binding StatusIcon}"
                                            Width="30"
                                            CanUserSort="False"
                                            CanUserResize="False">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="HorizontalAlignment" Value="Center"/>
                                    <Setter Property="FontSize"            Value="15"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>

                        <!-- Machine Name -->
                        <DataGridTextColumn Header="Machine"
                                            Binding="{Binding MachineName}"
                                            Width="130">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="FontFamily" Value="Consolas"/>
                                    <Setter Property="FontWeight" Value="SemiBold"/>
                                    <Setter Property="Foreground" Value="#E0E0FF"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>

                        <!-- Status -->
                        <DataGridTextColumn Header="Status"
                                            Binding="{Binding Status}"
                                            Width="115"/>

                        <!-- Build Stage -->
                        <DataGridTextColumn Header="Build Stage"
                                            Binding="{Binding BuildStage}"
                                            Width="110"/>

                        <!-- Machine Type -->
                        <DataGridTextColumn Header="Type"
                                            Binding="{Binding MachineType}"
                                            Width="80"/>

                        <!-- Retry count -->
                        <DataGridTextColumn Header="Retry"
                                            Binding="{Binding RetryCount}"
                                            Width="50">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="HorizontalAlignment" Value="Center"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>

                        <!-- Elapsed -->
                        <DataGridTextColumn Header="Elapsed"
                                            Binding="{Binding ElapsedDisplay}"
                                            Width="78">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="FontFamily" Value="Consolas"/>
                                    <Setter Property="FontSize"   Value="11.5"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>

                        <!-- Message -->
                        <DataGridTextColumn Header="Message"
                                            Binding="{Binding Message}"
                                            Width="*">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="FontSize" Value="11.5"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>

                        <!-- Cancel button -->
                        <DataGridTemplateColumn Header="Action" Width="72"
                                                CanUserSort="False"
                                                CanUserResize="False">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <Button Content="Cancel"
                                            Tag="{Binding MachineName}"
                                            x:Name="BtnCancelJob"
                                            Height="22" Width="62"
                                            FontSize="11"
                                            Background="#5C1A1A"
                                            Foreground="#FFB0B0"
                                            BorderBrush="#8B3030"/>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>
                    </DataGrid.Columns>
                </DataGrid>

                <!-- Activity Log -->
                <GroupBox Grid.Row="1" Header=" Activity Log ">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="26"/>
                        </Grid.RowDefinitions>
                        <ScrollViewer x:Name="LogScroller"
                                      Grid.Row="0"
                                      VerticalScrollBarVisibility="Auto"
                                      HorizontalScrollBarVisibility="Auto">
                            <TextBox x:Name="TxtLog"
                                     IsReadOnly="True"
                                     AcceptsReturn="True"
                                     TextWrapping="NoWrap"
                                     Background="#0E0E20"
                                     Foreground="#00E5A0"
                                     FontFamily="Consolas"
                                     FontSize="11"
                                     BorderThickness="0"
                                     Padding="4,2"/>
                        </ScrollViewer>
                        <Button Grid.Row="1" x:Name="BtnClearLog"
                                Content="Clear Log" HorizontalAlignment="Right"
                                Width="90" Height="22" Margin="0,2,2,0" FontSize="11"/>
                    </Grid>
                </GroupBox>

            </Grid>
        </Grid>

        <!-- ═══════════════════  STATUS BAR  ═══════════════════ -->
        <Border Grid.Row="2" Background="#151525" BorderBrush="#252545" BorderThickness="0,1,0,0">
            <Grid Margin="10,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="StatusBarLeft"  VerticalAlignment="Center" FontSize="11"/>
                <TextBlock x:Name="StatusBarRight" Grid.Column="1"
                           VerticalAlignment="Center" FontSize="11"
                           Foreground="#505070"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

# ─────────────────────────────────────────────────────────────────────────────
#  HELPER: ENV BADGE COLOR
# ─────────────────────────────────────────────────────────────────────────────
function Get-BMEnvColor {
    param([string]$Env)
    switch ($Env.ToLower()) {
        'prod' { return '#C62828' }   # Red   — production
        'uat'  { return '#6A1B9A' }   # Purple
        'qa'   { return '#E65100' }   # Orange
        'dev'  { return '#1565C0' }   # Blue
        default{ return '#37474F' }   # Grey
    }
}

# ─────────────────────────────────────────────────────────────────────────────
#  HELPER: ROW BACKGROUND COLOR
# ─────────────────────────────────────────────────────────────────────────────
function Get-BMRowBrush {
    param([string]$Status)
    switch ($Status) {
        'Completed' { return New-Object System.Windows.Media.SolidColorBrush(
                              [System.Windows.Media.Color]::FromArgb(200,10,60,20)) }
        'Failed'    { return New-Object System.Windows.Media.SolidColorBrush(
                              [System.Windows.Media.Color]::FromArgb(200,80,10,10)) }
        'Cancelled' { return New-Object System.Windows.Media.SolidColorBrush(
                              [System.Windows.Media.Color]::FromArgb(180,40,40,50)) }
        'Error'     { return New-Object System.Windows.Media.SolidColorBrush(
                              [System.Windows.Media.Color]::FromArgb(200,80,40,0))  }
        'Monitoring'{ return New-Object System.Windows.Media.SolidColorBrush(
                              [System.Windows.Media.Color]::FromArgb(100,10,30,80)) }
        default     { return [System.Windows.Media.Brushes]::Transparent }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
#  MAIN ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────
function Start-BMGui {
    <#
    .SYNOPSIS  Loads and shows the BuildMaster WPF window. Blocks until closed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Environment,
        [Parameter(Mandatory)][string]$LogPath
    )

    # ── Load XAML ─────────────────────────────────────────────────────────────
    try {
        [xml]$xamlDoc = $script:XAML
        $reader       = New-Object System.Xml.XmlNodeReader($xamlDoc)
        $window       = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        throw "Failed to load BuildMaster XAML: $($_.Exception.Message)"
    }

    # ── Get named controls ────────────────────────────────────────────────────
    $ctrl = @{}
    $controlNames = @(
        'EnvBadge','EnvLabel','EngineStatusLabel',
        'RegCredLabel','PrivCredLabel','CredsSummary',
        'BtnSetRegCred','BtnSetPrivCred',
        'TxtMachines','BtnStartRebuild','BtnClearCompleted',
        'TxtStagingWait','TxtPollInterval','TxtMaxRetries','BtnApplySettings',
        'JobGrid',
        'TxtLog','LogScroller','BtnClearLog',
        'StatusBarLeft','StatusBarRight'
    )
    foreach ($name in $controlNames) {
        $ctrl[$name] = $window.FindName($name)
        if ($null -eq $ctrl[$name]) {
            Write-Warning "Control '$name' not found in XAML"
        }
    }

    # ── Environment badge ─────────────────────────────────────────────────────
    $envColor = Get-BMEnvColor -Env $Environment
    $ctrl['EnvBadge'].Background = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.ColorConverter]::ConvertFromString($envColor))
    $ctrl['EnvLabel'].Text = $Environment.ToUpper()

    # ── Register UI log sink with engine ─────────────────────────────────────
    $logCtrl    = $ctrl['TxtLog']
    $logScroll  = $ctrl['LogScroller']
    Register-BMUILogSink -Sink {
        param([string]$line)
        $logCtrl.AppendText($line + "`n")
        $logScroll.ScrollToBottom()
    }

    # ── Bind ObservableCollection to DataGrid ─────────────────────────────────
    $ctrl['JobGrid'].ItemsSource = Get-BMJobs

    # ── Row coloring on load/refresh ──────────────────────────────────────────
    $ctrl['JobGrid'].Add_LoadingRow({
        param($s, $e)
        $item = $e.Row.DataContext
        if ($item -is [System.Management.Automation.PSCustomObject]) {
            $e.Row.Background = Get-BMRowBrush -Status $item.Status
        }
    })

    # ── Cancel button in DataGrid rows ────────────────────────────────────────
    $ctrl['JobGrid'].Add_PreviewMouseLeftButtonUp({
        param($s, $e)
        # Walk up the visual tree to find a Button
        $el = $e.OriginalSource
        $maxDepth = 10
        while ($null -ne $el -and $maxDepth-- -gt 0) {
            if ($el -is [System.Windows.Controls.Button] -and $el.Name -eq 'BtnCancelJob') {
                $machineName = $el.Tag
                if (![string]::IsNullOrWhiteSpace($machineName)) {
                    $result = [System.Windows.MessageBox]::Show(
                        "Cancel rebuild for $machineName?",
                        'Confirm Cancel',
                        [System.Windows.MessageBoxButton]::YesNo,
                        [System.Windows.MessageBoxImage]::Question)
                    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
                        Stop-BMJob -MachineName $machineName
                        Write-BMLog -MachineName $machineName -Message 'User confirmed cancel' -Level Warning
                    }
                }
                break
            }
            $el = [System.Windows.Media.VisualTreeHelper]::GetParent($el)
        }
    })

    # ─────────────────────────────────────────────────────────────────────────
    #  CREDENTIAL BUTTONS
    # ─────────────────────────────────────────────────────────────────────────

    $updateCredsDisplay = {
        # Regular account
        if (Test-BMRegularCredentialSet) {
            $regCred = Get-BMRegularCredential
            $ctrl['RegCredLabel'].Text       = "🔑 Regular: $($regCred.UserName)"
            $ctrl['RegCredLabel'].Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $ctrl['RegCredLabel'].Text       = '🔑 Regular: not set'
            $ctrl['RegCredLabel'].Foreground = [System.Windows.Media.Brushes]::Salmon
        }
        # Privileged account
        if (Test-BMPrivilegedInfoSet) {
            $pInfo = Get-BMPrivilegedInfo
            $label = if ($pInfo.Method -eq 'EPM') { '🔐 Priv: EPM (auto-elevate)' }
                     else { "🔐 Priv: $($pInfo.Credential.UserName)" }
            $ctrl['PrivCredLabel'].Text       = $label
            $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $ctrl['PrivCredLabel'].Text       = '🔐 Priv: not set'
            $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::Goldenrod
        }
        # Summary in left panel
        $regSet  = Test-BMRegularCredentialSet
        $privSet = Test-BMPrivilegedInfoSet
        $ctrl['CredsSummary'].Text = if ($regSet -and $privSet) { '✓ Both credentials ready' }
                                     elseif ($regSet)           { '⚠ Privileged not set (needed for physical machines)' }
                                     else                       { '✗ Regular credentials required to start' }
    }

    $ctrl['BtnSetRegCred'].Add_Click({
        try {
            Get-BMRegularCredential -Force | Out-Null
            & $updateCredsDisplay
            Write-BMLog -Message 'Regular credentials updated' -Level Info
        }
        catch {
            Write-BMLog -Message "Regular credential prompt failed: $($_.Exception.Message)" -Level Error
        }
    })

    $ctrl['BtnSetPrivCred'].Add_Click({
        try {
            Get-BMPrivilegedInfo -Force | Out-Null
            & $updateCredsDisplay
            Write-BMLog -Message 'Privileged credential info updated' -Level Info
        }
        catch {
            Write-BMLog -Message "Privileged credential prompt failed: $($_.Exception.Message)" -Level Error
        }
    })

    # ─────────────────────────────────────────────────────────────────────────
    #  START REBUILD BUTTON
    # ─────────────────────────────────────────────────────────────────────────

    $ctrl['BtnStartRebuild'].Add_Click({
        # Validate regular credentials
        if (-not (Test-BMRegularCredentialSet)) {
            [System.Windows.MessageBox]::Show(
                'Please set regular account credentials first.',
                'Credentials Required',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        # Parse machine names
        $raw = $ctrl['TxtMachines'].Text
        $machines = $raw -split "`n" |
                    ForEach-Object { $_.Trim().ToUpper() } |
                    Where-Object   { $_ -ne '' } |
                    Select-Object  -Unique

        if ($machines.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                'Enter at least one machine name.',
                'No Machines',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }

        # Apply settings
        $stagingWait  = 35
        $pollInterval = 30
        if ([int]::TryParse($ctrl['TxtStagingWait'].Text,  [ref]$stagingWait)  -and $stagingWait  -gt 0) {}
        if ([int]::TryParse($ctrl['TxtPollInterval'].Text, [ref]$pollInterval) -and $pollInterval -gt 0) {}
        Set-BMEngineConfig -StagingWaitMinutes $stagingWait -PollIntervalSeconds $pollInterval

        # Add jobs
        foreach ($m in $machines) {
            Add-BMJob -MachineName $m | Out-Null
        }

        # Start engine if not running
        if (-not (Test-BMEngineRunning)) {
            Start-BMEngine
        }

        $ctrl['TxtMachines'].Clear()
        Write-BMLog -Message ("Queued {0} machine(s) for rebuild" -f $machines.Count) -Level Info
    })

    # ─────────────────────────────────────────────────────────────────────────
    #  CLEAR COMPLETED BUTTON
    # ─────────────────────────────────────────────────────────────────────────

    $ctrl['BtnClearCompleted'].Add_Click({
        $before = (Get-BMJobs).Count
        Clear-BMCompletedJobs
        $removed = $before - (Get-BMJobs).Count
        Write-BMLog -Message "Cleared $removed completed/failed/cancelled job(s)" -Level Info
    })

    # ─────────────────────────────────────────────────────────────────────────
    #  APPLY SETTINGS BUTTON
    # ─────────────────────────────────────────────────────────────────────────

    $ctrl['BtnApplySettings'].Add_Click({
        $sw = 35; $pi = 30
        if (-not ([int]::TryParse($ctrl['TxtStagingWait'].Text,  [ref]$sw) -and $sw  -in 1..120)) {
            [System.Windows.MessageBox]::Show('Staging wait must be 1–120 minutes.', 'Invalid Setting') | Out-Null; return
        }
        if (-not ([int]::TryParse($ctrl['TxtPollInterval'].Text, [ref]$pi) -and $pi  -in 5..300)) {
            [System.Windows.MessageBox]::Show('Poll interval must be 5–300 seconds.', 'Invalid Setting') | Out-Null; return
        }
        Set-BMEngineConfig -StagingWaitMinutes $sw -PollIntervalSeconds $pi
        Write-BMLog -Message ("Settings applied — staging wait: {0}m  poll: {1}s" -f $sw, $pi) -Level Info
    })

    # ─────────────────────────────────────────────────────────────────────────
    #  CLEAR LOG BUTTON
    # ─────────────────────────────────────────────────────────────────────────

    $ctrl['BtnClearLog'].Add_Click({
        $ctrl['TxtLog'].Clear()
    })

    # ─────────────────────────────────────────────────────────────────────────
    #  UI REFRESH TIMER  (2 sec — updates grid rows and status bar)
    # ─────────────────────────────────────────────────────────────────────────

    $uiTimer          = New-Object System.Windows.Threading.DispatcherTimer
    $uiTimer.Interval = [timespan]::FromSeconds(2)

    $uiTimer.Add_Tick({
        # Refresh DataGrid row rendering (picks up property changes + row colors)
        try { $ctrl['JobGrid'].Items.Refresh() } catch { }

        # Status bar — job summary
        $jobs     = @(Get-BMJobs)
        $total    = $jobs.Count
        $active   = @($jobs | Where-Object { $_.Status -notin @('Completed','Failed','Cancelled') }).Count
        $done     = @($jobs | Where-Object { $_.Status -eq 'Completed' }).Count
        $failed   = @($jobs | Where-Object { $_.Status -eq 'Failed'    }).Count

        $engineState = if (Test-BMEngineRunning) { '● Engine running' } else { '○ Engine idle' }
        $ctrl['EngineStatusLabel'].Text = $engineState

        $ctrl['StatusBarLeft'].Text = if ($total -gt 0) {
            "Jobs: $total total  |  $active active  |  $done completed  |  $failed failed"
        } else {
            'No jobs queued. Enter machine names and click Start Rebuild.'
        }
        $ctrl['StatusBarRight'].Text = "Environment: $Environment  |  {0}" -f (Get-Date -Format 'HH:mm:ss')
    })

    # ─────────────────────────────────────────────────────────────────────────
    #  WINDOW EVENTS
    # ─────────────────────────────────────────────────────────────────────────

    $window.Add_Loaded({
        # Initial credential display
        & $updateCredsDisplay
        # Start UI refresh timer
        $uiTimer.Start()
        Write-BMLog -Message "BuildMaster GUI loaded — Environment: $Environment" -Level Info
        Write-BMLog -Message "Network path: \\FileServer\ITTools\BuildMaster\$Environment\BuildMaster.ps1" -Level Debug
    })

    $window.Add_Closing({
        param($s, $e)
        $activeJobs = @(Get-BMJobs | Where-Object {
            $_.Status -notin @('Completed','Failed','Cancelled')
        })
        if ($activeJobs.Count -gt 0) {
            $result = [System.Windows.MessageBox]::Show(
                "$($activeJobs.Count) active job(s) still running.`nClose anyway? This will NOT stop the rebuilds already in progress on the machines.",
                'Active Jobs',
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
            if ($result -eq [System.Windows.MessageBoxResult]::No) {
                $e.Cancel = $true
                return
            }
        }
        $uiTimer.Stop()
        Stop-BMEngine
        Clear-BMCredentials
    })

    # ─────────────────────────────────────────────────────────────────────────
    #  SHOW WINDOW  (blocks until closed)
    # ─────────────────────────────────────────────────────────────────────────

    Write-BMLog -Message 'Launching BuildMaster window...' -Level Info
    $window.ShowDialog() | Out-Null
}

# ─────────────────────────────────────────────────────────────────────────────
#  EXPORTS
# ─────────────────────────────────────────────────────────────────────────────
Export-ModuleMember -Function @('Start-BMGui')
