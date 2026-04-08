#Requires -Version 5.1
<#
.SYNOPSIS
    BM-GUI - WPF interface for BuildMaster
.DESCRIPTION
    Defines the full WPF window, wires all controls to engine/API functions,
    and drives a UI-refresh DispatcherTimer for live grid/log updates.

    Designed for PS 5.1 - no MVVM framework, no external dependencies.
    All UI work runs on the WPF Dispatcher thread.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
#  XAML DEFINITION
# -----------------------------------------------------------------------------
[string]$script:XAML = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="BuildMaster - Machine Rebuild Manager"
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

        <!-- ===================  HEADER  =================== -->
        <Border Grid.Row="0" Background="#1A1A35" BorderBrush="#303058" BorderThickness="0,0,0,1">
            <Grid Margin="12,0">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Left: title + env badge -->
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="" FontSize="26" Foreground="#7070FF"
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

        <!-- ===================  MAIN CONTENT  =================== -->
        <Grid Grid.Row="1" Margin="10,8,10,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="265"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- --------------  LEFT PANEL  -------------- -->
            <StackPanel Grid.Column="0" Margin="0,0,10,0">

                <!-- Credentials -->
                <GroupBox Header=" Credentials " Margin="0,0,0,8">
                    <StackPanel Margin="2">
                        <Button x:Name="BtnSetRegCred"  Content=" Set Regular Account"
                                Height="32" Margin="0,4,0,0"/>
                        <Button x:Name="BtnSetPrivCred" Content=" Set Privileged Account"
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
                                 Height="130"
                                 AcceptsReturn="True"
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"
                                 TextWrapping="NoWrap"
                                 FontFamily="Consolas" FontSize="12"
                                 VerticalContentAlignment="Top"/>

                        <!-- Schedule toggle -->
                        <CheckBox x:Name="ChkSchedule"
                                  Content="Schedule for later"
                                  Foreground="#A0A0D0"
                                  Margin="2,8,0,0"/>

                        <!-- Schedule date/time panel (hidden until toggle is checked) -->
                        <StackPanel x:Name="SchedulePanel" Visibility="Collapsed" Margin="2,6,0,0">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="6"/>
                                    <ColumnDefinition Width="68"/>
                                </Grid.ColumnDefinitions>
                                <DatePicker x:Name="DpScheduleDate"
                                            Grid.Column="0" Height="26"
                                            Background="#1E1E35" Foreground="#D0D0F0"
                                            BorderBrush="#404068"
                                            FontSize="12"/>
                                <TextBox x:Name="TxtScheduleTime"
                                         Grid.Column="2" Height="26"
                                         Text="08:00"
                                         FontFamily="Consolas" FontSize="12"
                                         ToolTip="24-hour time  HH:mm"/>
                            </Grid>
                            <TextBlock Text="App must stay open for scheduled builds."
                                       FontSize="10" Foreground="#505070"
                                       Margin="0,3,0,0" TextWrapping="Wrap"/>
                        </StackPanel>

                        <Button x:Name="BtnStartRebuild"
                                Content=">  Start Rebuild"
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

            <!-- --------------  RIGHT PANEL  -------------- -->
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
                                            Width="110"/>

                        <!-- Scheduled For -->
                        <DataGridTextColumn Header="Scheduled"
                                            Binding="{Binding ScheduleLabel}"
                                            Width="88">
                            <DataGridTextColumn.ElementStyle>
                                <Style TargetType="TextBlock">
                                    <Setter Property="FontFamily" Value="Consolas"/>
                                    <Setter Property="FontSize"   Value="11"/>
                                </Style>
                            </DataGridTextColumn.ElementStyle>
                        </DataGridTextColumn>

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

        <!-- ===================  STATUS BAR  =================== -->
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

# -----------------------------------------------------------------------------
#  HELPER: ENV BADGE COLOR
# -----------------------------------------------------------------------------
function Get-BMEnvColor {
    param([string]$Env)
    switch ($Env.ToLower()) {
        'prod' { return '#C62828' }   # Red   - production
        'uat'  { return '#6A1B9A' }   # Purple
        'qa'   { return '#E65100' }   # Orange
        'dev'  { return '#1565C0' }   # Blue
        default{ return '#37474F' }   # Grey
    }
}

# -----------------------------------------------------------------------------
#  HELPER: ROW BACKGROUND COLOR
# -----------------------------------------------------------------------------
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
        'Monitoring' { return New-Object System.Windows.Media.SolidColorBrush(
                               [System.Windows.Media.Color]::FromArgb(100,10,30,80)) }
        'Scheduled'  { return New-Object System.Windows.Media.SolidColorBrush(
                               [System.Windows.Media.Color]::FromArgb(120,20,50,80)) }
        default      { return [System.Windows.Media.Brushes]::Transparent }
    }
}

# -----------------------------------------------------------------------------
#  MAIN ENTRY POINT
# -----------------------------------------------------------------------------
function Start-BMGui {
    <#
    .SYNOPSIS  Loads and shows the BuildMaster WPF window. Blocks until closed.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Environment,
        [Parameter(Mandatory)][string]$LogPath
    )

    # -- Load XAML -------------------------------------------------------------
    try {
        [xml]$xamlDoc = $script:XAML
        $reader       = New-Object System.Xml.XmlNodeReader($xamlDoc)
        $window       = [Windows.Markup.XamlReader]::Load($reader)
    }
    catch {
        throw "Failed to load BuildMaster XAML: $($_.Exception.Message)"
    }

    # -- Get named controls ----------------------------------------------------
    $ctrl = @{}
    $controlNames = @(
        'EnvBadge','EnvLabel','EngineStatusLabel',
        'RegCredLabel','PrivCredLabel','CredsSummary',
        'BtnSetRegCred','BtnSetPrivCred',
        'TxtMachines','BtnStartRebuild','BtnClearCompleted',
        'ChkSchedule','SchedulePanel','DpScheduleDate','TxtScheduleTime',
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

    # -- Environment badge -----------------------------------------------------
    $envColor = Get-BMEnvColor -Env $Environment
    $ctrl['EnvBadge'].Background = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.ColorConverter]::ConvertFromString($envColor))
    $ctrl['EnvLabel'].Text = $Environment.ToUpper()

    # -- Register UI log sink with engine -------------------------------------
    $logCtrl    = $ctrl['TxtLog']
    $logScroll  = $ctrl['LogScroller']
    Register-BMUILogSink -Sink {
        param([string]$line)
        $logCtrl.AppendText($line + "`n")
        $logScroll.ScrollToBottom()
    }

    # -- Bind ObservableCollection to DataGrid ---------------------------------
    $ctrl['JobGrid'].ItemsSource = Get-BMJobs

    # -- Row coloring on load/refresh ------------------------------------------
    $ctrl['JobGrid'].Add_LoadingRow({
        param($s, $e)
        $item = $e.Row.DataContext
        if ($item -is [System.Management.Automation.PSCustomObject]) {
            $e.Row.Background = Get-BMRowBrush -Status $item.Status
        }
    })

    # -- Cancel button in DataGrid rows ----------------------------------------
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

    # -------------------------------------------------------------------------
    #  CREDENTIAL BUTTONS
    # -------------------------------------------------------------------------

    $updateCredsDisplay = {
        # Regular account
        if (Test-BMRegularCredentialSet) {
            $regCred = Get-BMRegularCredential
            $ctrl['RegCredLabel'].Text       = " Regular: $($regCred.UserName)"
            $ctrl['RegCredLabel'].Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $ctrl['RegCredLabel'].Text       = ' Regular: not set'
            $ctrl['RegCredLabel'].Foreground = [System.Windows.Media.Brushes]::Salmon
        }
        # Privileged account
        if (Test-BMPrivilegedInfoSet) {
            $pInfo = Get-BMPrivilegedInfo
            $label = if ($pInfo.Method -eq 'EPM') { ' Priv: EPM (auto-elevate)' }
                     else { " Priv: $($pInfo.Credential.UserName)" }
            $ctrl['PrivCredLabel'].Text       = $label
            $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $ctrl['PrivCredLabel'].Text       = ' Priv: not set'
            $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::Goldenrod
        }
        # Summary in left panel
        $regSet  = Test-BMRegularCredentialSet
        $privSet = Test-BMPrivilegedInfoSet
        $ctrl['CredsSummary'].Text = if ($regSet -and $privSet) { '[+] Both credentials ready' }
                                     elseif ($regSet)           { '[!] Privileged not set (needed for physical machines)' }
                                     else                       { '[X] Regular credentials required to start' }
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

    # -------------------------------------------------------------------------
    #  SCHEDULE TOGGLE  (show/hide date+time picker, relabel Start button)
    # -------------------------------------------------------------------------

    $ctrl['ChkSchedule'].Add_Checked({
        $ctrl['SchedulePanel'].Visibility    = [System.Windows.Visibility]::Visible
        $ctrl['BtnStartRebuild'].Content     = '>  Schedule Rebuild'
        $ctrl['BtnStartRebuild'].Background  = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.ColorConverter]::ConvertFromString('#1A3A6E'))
        $ctrl['BtnStartRebuild'].BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.ColorConverter]::ConvertFromString('#2E5FAE'))
        # Default the date picker to today if nothing is selected
        if ($null -eq $ctrl['DpScheduleDate'].SelectedDate) {
            $ctrl['DpScheduleDate'].SelectedDate = [datetime]::Today
        }
    })

    $ctrl['ChkSchedule'].Add_Unchecked({
        $ctrl['SchedulePanel'].Visibility    = [System.Windows.Visibility]::Collapsed
        $ctrl['BtnStartRebuild'].Content     = '>  Start Rebuild'
        $ctrl['BtnStartRebuild'].Background  = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.ColorConverter]::ConvertFromString('#1B5E20'))
        $ctrl['BtnStartRebuild'].BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.ColorConverter]::ConvertFromString('#388E3C'))
    })

    # -------------------------------------------------------------------------
    #  START / SCHEDULE REBUILD BUTTON
    # -------------------------------------------------------------------------

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

        # -- Resolve scheduled time if toggle is checked ----------------------
        $scheduleFor = $null
        if ($ctrl['ChkSchedule'].IsChecked -eq $true) {
            $selectedDate = $ctrl['DpScheduleDate'].SelectedDate
            $timeText     = $ctrl['TxtScheduleTime'].Text.Trim()

            if ($null -eq $selectedDate) {
                [System.Windows.MessageBox]::Show(
                    'Please select a date for the scheduled rebuild.',
                    'Date Required',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            $parsedTime = $null
            if (-not [datetime]::TryParseExact($timeText, 'HH:mm',
                    [System.Globalization.CultureInfo]::InvariantCulture,
                    [System.Globalization.DateTimeStyles]::None, [ref]$parsedTime)) {
                [System.Windows.MessageBox]::Show(
                    "Invalid time format: '$timeText'`nUse 24-hour format HH:mm (e.g. 14:30).",
                    'Invalid Time',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }

            $scheduleFor = $selectedDate.Date.Add($parsedTime.TimeOfDay)

            if ($scheduleFor -le [datetime]::Now) {
                [System.Windows.MessageBox]::Show(
                    ("Scheduled time {0} is in the past.`nPlease choose a future date/time." -f $scheduleFor.ToString('MM/dd/yyyy HH:mm')),
                    'Time in Past',
                    [System.Windows.MessageBoxButton]::OK,
                    [System.Windows.MessageBoxImage]::Warning) | Out-Null
                return
            }
        }

        # Apply settings
        $stagingWait  = 35
        $pollInterval = 30
        if ([int]::TryParse($ctrl['TxtStagingWait'].Text,  [ref]$stagingWait)  -and $stagingWait  -gt 0) {}
        if ([int]::TryParse($ctrl['TxtPollInterval'].Text, [ref]$pollInterval) -and $pollInterval -gt 0) {}
        Set-BMEngineConfig -StagingWaitMinutes $stagingWait -PollIntervalSeconds $pollInterval

        # Add jobs
        foreach ($m in $machines) {
            Add-BMJob -MachineName $m -ScheduledFor $scheduleFor | Out-Null
        }

        # Start engine if not running (engine must run to watch for scheduled times)
        if (-not (Test-BMEngineRunning)) {
            Start-BMEngine
        }

        $ctrl['TxtMachines'].Clear()

        if ($null -ne $scheduleFor) {
            Write-BMLog -Message ("Scheduled {0} machine(s) for {1}" -f $machines.Count, $scheduleFor.ToString('MM/dd/yyyy HH:mm')) -Level Info
        } else {
            Write-BMLog -Message ("Queued {0} machine(s) for immediate rebuild" -f $machines.Count) -Level Info
        }
    })

    # -------------------------------------------------------------------------
    #  CLEAR COMPLETED BUTTON
    # -------------------------------------------------------------------------

    $ctrl['BtnClearCompleted'].Add_Click({
        $before = (Get-BMJobs).Count
        Clear-BMCompletedJobs
        $removed = $before - (Get-BMJobs).Count
        Write-BMLog -Message "Cleared $removed completed/failed/cancelled job(s)" -Level Info
    })

    # -------------------------------------------------------------------------
    #  APPLY SETTINGS BUTTON
    # -------------------------------------------------------------------------

    $ctrl['BtnApplySettings'].Add_Click({
        $sw = 35; $pi = 30
        if (-not ([int]::TryParse($ctrl['TxtStagingWait'].Text,  [ref]$sw) -and $sw  -in 1..120)) {
            [System.Windows.MessageBox]::Show('Staging wait must be 1-120 minutes.', 'Invalid Setting') | Out-Null; return
        }
        if (-not ([int]::TryParse($ctrl['TxtPollInterval'].Text, [ref]$pi) -and $pi  -in 5..300)) {
            [System.Windows.MessageBox]::Show('Poll interval must be 5-300 seconds.', 'Invalid Setting') | Out-Null; return
        }
        Set-BMEngineConfig -StagingWaitMinutes $sw -PollIntervalSeconds $pi
        Write-BMLog -Message ("Settings applied - staging wait: {0}m  poll: {1}s" -f $sw, $pi) -Level Info
    })

    # -------------------------------------------------------------------------
    #  CLEAR LOG BUTTON
    # -------------------------------------------------------------------------

    $ctrl['BtnClearLog'].Add_Click({
        $ctrl['TxtLog'].Clear()
    })

    # -------------------------------------------------------------------------
    #  UI REFRESH TIMER  (2 sec - updates grid rows and status bar)
    # -------------------------------------------------------------------------

    $uiTimer          = New-Object System.Windows.Threading.DispatcherTimer
    $uiTimer.Interval = [timespan]::FromSeconds(2)

    $uiTimer.Add_Tick({
        # Refresh DataGrid row rendering (picks up property changes + row colors)
        try { $ctrl['JobGrid'].Items.Refresh() } catch { }

        # Status bar - job summary
        $jobs      = @(Get-BMJobs)
        $total     = $jobs.Count
        $scheduled = @($jobs | Where-Object { $_.Status -eq 'Scheduled'  }).Count
        $active    = @($jobs | Where-Object { $_.Status -notin @('Completed','Failed','Cancelled','Scheduled') }).Count
        $done      = @($jobs | Where-Object { $_.Status -eq 'Completed'  }).Count
        $failed    = @($jobs | Where-Object { $_.Status -eq 'Failed'     }).Count

        $engineState = if (Test-BMEngineRunning) { '* Engine running' } else { 'o Engine idle' }
        $ctrl['EngineStatusLabel'].Text = $engineState

        $ctrl['StatusBarLeft'].Text = if ($total -gt 0) {
            $sb = "Jobs: $total total  |  $active active  |  $done completed  |  $failed failed"
            if ($scheduled -gt 0) { $sb += "  |  $scheduled scheduled" }
            $sb
        } else {
            'No jobs queued. Enter machine names and click Start Rebuild.'
        }
        $ctrl['StatusBarRight'].Text = "Environment: $Environment  |  {0}" -f (Get-Date -Format 'HH:mm:ss')
    })

    # -------------------------------------------------------------------------
    #  WINDOW EVENTS
    # -------------------------------------------------------------------------

    $window.Add_Loaded({
        # Initial credential display
        & $updateCredsDisplay
        # Start UI refresh timer
        $uiTimer.Start()
        Write-BMLog -Message "BuildMaster GUI loaded - Environment: $Environment" -Level Info
        Write-BMLog -Message "Network path: \\FileServer\ITTools\BuildMaster\$Environment\BuildMaster.ps1" -Level Debug
    })

    $window.Add_Closing({
        param($s, $e)
        $activeJobs = @(Get-BMJobs | Where-Object {
            $_.Status -notin @('Completed','Failed','Cancelled')
        })
        if ($activeJobs.Count -gt 0) {
            $runningCount   = @($activeJobs | Where-Object { $_.Status -ne 'Scheduled' }).Count
            $scheduledCount = @($activeJobs | Where-Object { $_.Status -eq 'Scheduled' }).Count
            $detail = ''
            if ($runningCount   -gt 0) { $detail += "$runningCount active"    }
            if ($scheduledCount -gt 0) { $detail += "  $scheduledCount scheduled" }
            $result = [System.Windows.MessageBox]::Show(
                "$($activeJobs.Count) job(s) still pending ($($detail.Trim())).`nClose anyway? Active rebuilds already sent to machines will continue; scheduled jobs will be lost.",
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

    # -------------------------------------------------------------------------
    #  SHOW WINDOW  (blocks until closed)
    # -------------------------------------------------------------------------

    Write-BMLog -Message 'Launching BuildMaster window...' -Level Info
    $window.ShowDialog() | Out-Null
}

# -----------------------------------------------------------------------------
#  EXPORTS
# -----------------------------------------------------------------------------
Export-ModuleMember -Function @('Start-BMGui')
