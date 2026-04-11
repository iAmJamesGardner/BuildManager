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
        <!-- Dark-theme ProgressBar template - used in the job grid Progress column.
             PART_Track measures available width; PART_Indicator is sized by the control. -->
        <ControlTemplate x:Key="DarkProgressBarTemplate" TargetType="ProgressBar">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="1" CornerRadius="2">
                <Grid Margin="1">
                    <Rectangle x:Name="PART_Track"     Fill="Transparent"/>
                    <Rectangle x:Name="PART_Indicator" HorizontalAlignment="Left"
                               Fill="{TemplateBinding Foreground}"
                               RadiusX="1" RadiusY="1"/>
                </Grid>
            </Border>
        </ControlTemplate>
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

            <!-- ==================  LEFT PANEL  ================== -->
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
                        <!-- Shown only when session is blocked from admin actions -->
                        <Border x:Name="BlockedBanner" Visibility="Collapsed"
                                Background="#3A1010" BorderBrush="#8B2020"
                                BorderThickness="1" CornerRadius="3"
                                Margin="0,6,0,0" Padding="6,5">
                            <TextBlock x:Name="BlockedBannerText"
                                       TextWrapping="Wrap" FontSize="10.5"
                                       Foreground="#FF7070" FontWeight="SemiBold"/>
                        </Border>
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
                            <RowDefinition Height="Auto"/>
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

                        <!-- What If mode: DEV environment only, shown programmatically -->
                        <CheckBox  Grid.Row="4" Grid.Column="0" Grid.ColumnSpan="2"
                                   x:Name="ChkWhatIf"
                                   Content="What If Mode  (no commands sent to machines)"
                                   Foreground="#FF9800" FontSize="10.5"
                                   Margin="0,10,0,2"
                                   Visibility="Collapsed"/>
                    </Grid>
                </GroupBox>
            </StackPanel>

            <!-- ==================  RIGHT PANEL  ================= -->
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

                        <!-- Progress bar -->
                        <DataGridTemplateColumn Header="Progress" Width="100"
                                                CanUserSort="False" CanUserResize="True">
                            <DataGridTemplateColumn.CellTemplate>
                                <DataTemplate>
                                    <ProgressBar Value="{Binding ProgressValue}" Maximum="100"
                                                 Height="14" Margin="4,2">
                                        <ProgressBar.Style>
                                            <Style TargetType="ProgressBar">
                                                <Setter Property="Template"    Value="{StaticResource DarkProgressBarTemplate}"/>
                                                <Setter Property="Background"  Value="#1A1A40"/>
                                                <Setter Property="Foreground"  Value="#4488FF"/>
                                                <Setter Property="BorderBrush" Value="#404068"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Status}" Value="Completed">
                                                        <Setter Property="Foreground" Value="#00C853"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Monitoring">
                                                        <Setter Property="Foreground" Value="#00BCD4"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Error">
                                                        <Setter Property="Foreground" Value="#FF9800"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Failed">
                                                        <Setter Property="Foreground" Value="#FF1744"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding Status}" Value="Cancelled">
                                                        <Setter Property="Foreground" Value="#607D8B"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </ProgressBar.Style>
                                    </ProgressBar>
                                </DataTemplate>
                            </DataGridTemplateColumn.CellTemplate>
                        </DataGridTemplateColumn>

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
        'BlockedBanner','BlockedBannerText',
        'BtnSetRegCred','BtnSetPrivCred',
        'TxtMachines','BtnStartRebuild','BtnClearCompleted',
        'ChkSchedule','SchedulePanel','DpScheduleDate','TxtScheduleTime',
        'TxtStagingWait','TxtPollInterval','TxtMaxRetries','BtnApplySettings',
        'ChkWhatIf',
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
    }.GetNewClosure()

    # -- Tell the engine where BM-API.psm1 lives (needed for background runspace)
    # LogPath is <root>\Logs; modules are one level up in <root>\Modules.
    $modulesRoot = Join-Path (Split-Path -Parent $LogPath) 'Modules'
    Set-BMEngineConfig -ModulePath (Join-Path $modulesRoot 'BM-API.psm1')

    # -- Bind ObservableCollection to DataGrid ---------------------------------
    # Must use Get-BMJobCollection, NOT Get-BMJobs.
    # Get-BMJobs returns enumerated items (or $null for empty); Get-BMJobCollection
    # returns the live ObservableCollection reference so the DataGrid receives
    # CollectionChanged events when jobs are added/removed.
    $ctrl['JobGrid'].ItemsSource = Get-BMJobCollection

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
        $machineName = $null   # pre-initialise for Set-StrictMode -Version Latest
        $el = $e.OriginalSource
        $maxDepth = 10
        while ($null -ne $el -and $maxDepth-- -gt 0) {
            if ($el -is [System.Windows.Controls.Button] -and
                $el.Content -eq 'Cancel' -and
                -not [string]::IsNullOrWhiteSpace($el.Tag)) {
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
    }.GetNewClosure())

    # -------------------------------------------------------------------------
    #  CREDENTIAL BUTTONS
    # -------------------------------------------------------------------------

    $updateCredsDisplay = {
        $ctx = Get-BMSessionContext

        # -- Header: Regular credential label --------------------------------
        if ($ctx.IsPrivilegedSession) {
            if (Test-BMRegularCredentialSet) {
                $rc = Get-BMRegularCredential
                $ctrl['RegCredLabel'].Text       = (' Reg: {0}' -f $rc.UserName)
                $ctrl['RegCredLabel'].Foreground = [System.Windows.Media.Brushes]::LightGreen
            } else {
                $ctrl['RegCredLabel'].Text       = ' Reg: REQUIRED (not set)'
                $ctrl['RegCredLabel'].Foreground = [System.Windows.Media.Brushes]::Salmon
            }
        } else {
            # Regular session - current Windows token IS the regular account
            $ctrl['RegCredLabel'].Text       = (' Reg: {0} (session)' -f $ctx.FullUser)
            $ctrl['RegCredLabel'].Foreground = [System.Windows.Media.Brushes]::LightGreen
        }

        # -- Header: Elevation / privileged label ----------------------------
        switch ($ctx.ElevationMethod) {
            'DirectSession' {
                $ctrl['PrivCredLabel'].Text       = (' Priv: {0} (current session)' -f $ctx.FullUser)
                $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::LightGreen
            }
            'EPM' {
                $ctrl['PrivCredLabel'].Text       = ' Priv: CyberArk EPM'
                $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::LightGreen
            }
            'PrivAccount' {
                if ($ctx.Blocked) {
                    $ctrl['PrivCredLabel'].Text       = ' Priv: REQUIRED - relaunch as priv account'
                    $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::Salmon
                } else {
                    $wif = Get-BMWhatIfMode
                    $ctrl['PrivCredLabel'].Text       = if ($wif) { ' Priv: WHAT-IF mode active' }
                                                        else      { ' Priv: not available (DEV)' }
                    $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::Goldenrod
                }
            }
            default {
                $ctrl['PrivCredLabel'].Text       = ' Priv: Detecting...'
                $ctrl['PrivCredLabel'].Foreground = [System.Windows.Media.Brushes]::Gray
            }
        }

        # -- Left panel: summary text ----------------------------------------
        $regReady  = Test-BMRegularCredentialSet
        $wifMode   = Get-BMWhatIfMode
        $elevReady = $ctx.ElevationMethod -in @('DirectSession','EPM') -or $wifMode

        if ($ctx.Blocked) {
            $ctrl['CredsSummary'].Text       = '[!] No elevation - VMs OK, physicals will fail'
            $ctrl['CredsSummary'].Foreground = [System.Windows.Media.Brushes]::Goldenrod
        } elseif ($ctx.IsPrivilegedSession -and -not $regReady) {
            $ctrl['CredsSummary'].Text       = '[!] Set regular credentials to enable API calls'
            $ctrl['CredsSummary'].Foreground = [System.Windows.Media.Brushes]::Goldenrod
        } elseif ($regReady -and $elevReady) {
            $sessionType = switch ($ctx.ElevationMethod) {
                'DirectSession' { 'Privileged session' }
                'EPM'           { 'Regular + CyberArk EPM' }
                default         { if ($wifMode) { 'DEV - WhatIf mode' } else { 'DEV - limited' } }
            }
            $ctrl['CredsSummary'].Text       = ('[+] Ready  ({0})' -f $sessionType)
            $ctrl['CredsSummary'].Foreground = [System.Windows.Media.Brushes]::LightGreen
        } else {
            $ctrl['CredsSummary'].Text       = '[?] Check configuration'
            $ctrl['CredsSummary'].Foreground = [System.Windows.Media.Brushes]::Goldenrod
        }
    }.GetNewClosure()

    # -- Set Regular Account button (only meaningful in privileged sessions) --
    $ctrl['BtnSetRegCred'].Add_Click({
        try {
            Get-BMRegularCredential -Force | Out-Null
            & $updateCredsDisplay
            Write-BMLog -Message 'Regular credentials updated' -Level Info
        }
        catch {
            [System.Windows.MessageBox]::Show(
                $_.Exception.Message, 'Credential Error',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
            Write-BMLog -Message ("Regular credential prompt failed: {0}" -f $_.Exception.Message) -Level Error
        }
    }.GetNewClosure())

    # -- Set Privileged Account button is obsolete in the new model -----------
    # Hidden at load time (see Add_Loaded below); keep stub in case of future use.
    $ctrl['BtnSetPrivCred'].Add_Click({
        Write-BMLog -Message 'Privileged account is managed via session context - no manual entry required.' -Level Info
    }.GetNewClosure())

    # -------------------------------------------------------------------------
    #  START / STOP BUTTON MODE HELPERS
    #  Call  & $setStartMode  or  & $setStopMode  from any handler.
    # -------------------------------------------------------------------------

    $setStartMode = {
        $ctrl['BtnStartRebuild'].Tag = 'start'
        if ($ctrl['ChkSchedule'].IsChecked -eq $true) {
            $ctrl['BtnStartRebuild'].Content     = '>  Schedule Rebuild'
            $ctrl['BtnStartRebuild'].Background  = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#1A3A6E'))
            $ctrl['BtnStartRebuild'].BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#2E5FAE'))
        } else {
            $ctrl['BtnStartRebuild'].Content     = '>  Start Rebuild'
            $ctrl['BtnStartRebuild'].Background  = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#1B5E20'))
            $ctrl['BtnStartRebuild'].BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#388E3C'))
        }
    }.GetNewClosure()

    $setStopMode = {
        $ctrl['BtnStartRebuild'].Tag         = 'stop'
        $ctrl['BtnStartRebuild'].Content     = ([char]0x25A0 + '  Stop All Builds')
        $ctrl['BtnStartRebuild'].Background  = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.ColorConverter]::ConvertFromString('#7F0000'))
        $ctrl['BtnStartRebuild'].BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.ColorConverter]::ConvertFromString('#B71C1C'))
    }.GetNewClosure()

    # -------------------------------------------------------------------------
    #  SCHEDULE TOGGLE  (show/hide date+time picker, relabel Start button)
    # -------------------------------------------------------------------------

    $ctrl['ChkSchedule'].Add_Checked({
        $ctrl['SchedulePanel'].Visibility = [System.Windows.Visibility]::Visible
        # Only relabel the button when not in stop mode
        if ($ctrl['BtnStartRebuild'].Tag -ne 'stop') {
            $ctrl['BtnStartRebuild'].Content     = '>  Schedule Rebuild'
            $ctrl['BtnStartRebuild'].Background  = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#1A3A6E'))
            $ctrl['BtnStartRebuild'].BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#2E5FAE'))
        }
        # Default the date picker to today if nothing is selected
        if ($null -eq $ctrl['DpScheduleDate'].SelectedDate) {
            $ctrl['DpScheduleDate'].SelectedDate = [datetime]::Today
        }
    }.GetNewClosure())

    $ctrl['ChkSchedule'].Add_Unchecked({
        $ctrl['SchedulePanel'].Visibility = [System.Windows.Visibility]::Collapsed
        # Only relabel the button when not in stop mode
        if ($ctrl['BtnStartRebuild'].Tag -ne 'stop') {
            $ctrl['BtnStartRebuild'].Content     = '>  Start Rebuild'
            $ctrl['BtnStartRebuild'].Background  = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#1B5E20'))
            $ctrl['BtnStartRebuild'].BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#388E3C'))
        }
    }.GetNewClosure())

    # -------------------------------------------------------------------------
    #  START / SCHEDULE REBUILD BUTTON
    # -------------------------------------------------------------------------

    $ctrl['BtnStartRebuild'].Add_Click({

        # ==========  STOP MODE  ==============================================
        # Button is currently "Stop All Builds" - cancel everything.
        if ($ctrl['BtnStartRebuild'].Tag -eq 'stop') {
            $active = @(Get-BMJobs | Where-Object {
                $_.Status -notin @('Completed','Failed','Cancelled')
            })
            if ($active.Count -gt 0) {
                $ans = [System.Windows.MessageBox]::Show(
                    ("Stop all $($active.Count) active / pending rebuild(s)?`n`n" +
                     "Jobs already staged will continue running on the machines;`n" +
                     "only this tool's monitoring will be cancelled."),
                    'Stop All Builds',
                    [System.Windows.MessageBoxButton]::YesNo,
                    [System.Windows.MessageBoxImage]::Warning)
                if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }
            }

            Stop-AllBMJobs | Out-Null
            Stop-BMEngine
            & $setStartMode
            try { $ctrl['JobGrid'].Items.Refresh() } catch { }
            Write-BMLog -Message 'All active builds stopped by user.' -Level Warning
            return
        }

        # ==========  START MODE  =============================================

        # -- No elevation available: warn but allow VMs to proceed -------------
        $sessCtx = Get-BMSessionContext
        if ($sessCtx.Blocked) {
            $answer = [System.Windows.MessageBox]::Show(
                ("No elevation available for this session.`n`n" +
                 "VM / DiDC machines will rebuild normally via the VirtualWorks API.`n" +
                 "Physical / laptop machines will FAIL at the reboot step.`n`n" +
                 "If your list contains only VMs, click Yes to proceed.`n" +
                 "If it contains physical machines, click No and relaunch as a $($sessCtx.Domain) privileged account."),
                'Elevation Warning',
                [System.Windows.MessageBoxButton]::YesNo,
                [System.Windows.MessageBoxImage]::Warning)
            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }

        # -- Regular credentials required in privileged sessions ---------------
        if (-not (Test-BMRegularCredentialSet)) {
            [System.Windows.MessageBox]::Show(
                'Please set your regular account credentials first.' + "`n`n" +
                'These are needed for BuildMaster and VirtualWorks API calls.',
                'Credentials Required',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning) | Out-Null
            return
        }

        # Parse machine names (@() forces array so .Count is always valid)
        $raw      = $ctrl['TxtMachines'].Text
        $machines = @($raw -split "`n" |
                      ForEach-Object { $_.Trim().ToUpper() } |
                      Where-Object   { $_ -ne '' } |
                      Select-Object  -Unique)

        if ($machines.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                'Enter at least one machine name.',
                'No Machines',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information) | Out-Null
            return
        }

        # -- Resolve scheduled time if toggle is checked -----------------------
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
                    ("Scheduled time {0} is in the past.`nPlease choose a future date/time." -f
                     $scheduleFor.ToString('MM/dd/yyyy HH:mm')),
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

        # Add jobs - ObservableCollection fires CollectionChanged immediately;
        # rows appear in the DataGrid as soon as each Add-BMJob returns.
        foreach ($m in $machines) {
            Add-BMJob -MachineName $m -ScheduledFor $scheduleFor | Out-Null
        }

        # Start engine if not already running
        if (-not (Test-BMEngineRunning)) {
            Start-BMEngine
        }

        # Request a fast tick (500 ms) so staging begins shortly after the
        # click handler returns and the UI has rendered the new rows.
        # Staging runs in a background runspace so the UI thread stays free.
        Request-BMQuickTick

        $ctrl['TxtMachines'].Clear()

        if ($null -ne $scheduleFor) {
            Write-BMLog -Message ("Scheduled {0} machine(s) for {1}" -f
                                  $machines.Count, $scheduleFor.ToString('MM/dd/yyyy HH:mm')) -Level Info
        } else {
            Write-BMLog -Message ("Queued {0} machine(s) for rebuild" -f $machines.Count) -Level Info
        }

        # Switch button to Stop mode so user can halt everything mid-run
        & $setStopMode

    }.GetNewClosure())

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
    }.GetNewClosure())

    # -------------------------------------------------------------------------
    #  WHAT-IF MODE CHECKBOX  (DEV only - shown/enabled in Add_Loaded)
    # -------------------------------------------------------------------------

    $ctrl['ChkWhatIf'].Add_Checked({
        Set-BMWhatIfMode -Enabled $true
        & $updateCredsDisplay
        Write-BMLog -Message '[DEV] What If mode ENABLED - no commands will be sent to machines.' -Level Warning
    }.GetNewClosure())

    $ctrl['ChkWhatIf'].Add_Unchecked({
        Set-BMWhatIfMode -Enabled $false
        & $updateCredsDisplay
        Write-BMLog -Message '[DEV] What If mode disabled.' -Level Info
    }.GetNewClosure())

    # -------------------------------------------------------------------------
    #  CLEAR LOG BUTTON
    # -------------------------------------------------------------------------

    $ctrl['BtnClearLog'].Add_Click({
        $ctrl['TxtLog'].Clear()
    }.GetNewClosure())

    # -------------------------------------------------------------------------
    #  UI REFRESH TIMER  (2 sec - updates grid rows and status bar)
    # -------------------------------------------------------------------------

    $uiTimer          = New-Object System.Windows.Threading.DispatcherTimer
    $uiTimer.Interval = [timespan]::FromSeconds(2)

    $uiTimer.Add_Tick({
        try {
            # Refresh DataGrid row rendering (picks up property changes + row colours)
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
                $sb = ("Jobs: {0} total  |  {1} active  |  {2} completed  |  {3} failed" -f
                       $total, $active, $done, $failed)
                if ($scheduled -gt 0) { $sb += ("  |  {0} scheduled" -f $scheduled) }
                $sb
            } else {
                'No jobs queued. Enter machine names and click Start Rebuild.'
            }
            $ctrl['StatusBarRight'].Text = ("Environment: {0}  |  {1}" -f $Environment, (Get-Date -Format 'HH:mm:ss'))

            # Auto-reset Start/Stop button once all jobs reach a terminal state.
            # Only fires when in stop mode (meaning a batch was actually started).
            if ($ctrl['BtnStartRebuild'].Tag -eq 'stop') {
                $nonTerminal = @($jobs | Where-Object {
                    $_.Status -notin @('Completed','Failed','Cancelled')
                }).Count
                if ($nonTerminal -eq 0 -and $total -gt 0) {
                    & $setStartMode
                    try { Stop-BMEngine } catch { }
                    Write-BMLog -Message ('All builds finished - engine stopped.') -Level Info
                }
            }
        }
        catch {
            # Non-fatal UI refresh error - swallow to prevent crashing ShowDialog
            try { Write-BMLog -Message ("UI refresh error: {0}" -f $_.Exception.Message) -Level Warning } catch { }
        }
    }.GetNewClosure())

    # -------------------------------------------------------------------------
    #  WINDOW EVENTS
    # -------------------------------------------------------------------------

    $window.Add_Loaded({
        $ctx = Get-BMSessionContext

        # -- Configure session-specific UI -----------------------------------

        # "Set Privileged Account" button is obsolete in the new auth model
        $ctrl['BtnSetPrivCred'].Visibility = [System.Windows.Visibility]::Collapsed

        # "Set Regular Account" only needed for privileged sessions
        if (-not $ctx.IsPrivilegedSession) {
            $ctrl['BtnSetRegCred'].Visibility = [System.Windows.Visibility]::Collapsed
        } else {
            $ctrl['BtnSetRegCred'].Content    = ' Set Regular Account (required)'
            $ctrl['BtnSetRegCred'].Background = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#3A1A00'))
            $ctrl['BtnSetRegCred'].BorderBrush = New-Object System.Windows.Media.SolidColorBrush(
                [System.Windows.Media.ColorConverter]::ConvertFromString('#7A4A00'))
        }

        # Blocked state - show warning banner but keep Start button enabled.
        # VM/DiDC machines work fine with a regular account via the VirtualWorks API.
        # Only physical machine reboots require elevation; those jobs will fail fast
        # with a clear message if the list contains physicals.
        if ($ctx.Blocked) {
            $ctrl['BlockedBannerText'].Text   = $ctx.BlockReason
            $ctrl['BlockedBanner'].Visibility = [System.Windows.Visibility]::Visible
        }

        # What If mode checkbox - DEV only; also show when regular session has
        # no EPM in any environment (since the checkbox is harmless in DEV)
        if ($Environment -eq 'dev' -and $ctx.ElevationMethod -eq 'PrivAccount') {
            $ctrl['ChkWhatIf'].Visibility = [System.Windows.Visibility]::Visible
            Write-BMLog -Message '[DEV] WhatIf mode checkbox enabled - no EPM or privileged session detected.' -Level Warning
        }

        # -- Initialise Start button to start mode ----------------------------
        $ctrl['BtnStartRebuild'].Tag = 'start'

        # -- Initial credential display & startup log --------------------------
        & $updateCredsDisplay
        $uiTimer.Start()
        Write-BMLog -Message ("BuildMaster GUI loaded - Environment: {0}" -f $Environment) -Level Info
        Write-BMLog -Message ("Session: {0}  Elevation: {1}  Blocked: {2}" -f `
                              $ctx.FullUser, $ctx.ElevationMethod, $ctx.Blocked) -Level Info
        # Derive the script root from the log path ($LogPath = <root>\Logs)
        $scriptDir = Split-Path -Parent $LogPath
        Write-BMLog -Message ("Script root: {0}" -f $scriptDir) -Level Debug
    }.GetNewClosure())

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
    }.GetNewClosure())

    # -------------------------------------------------------------------------
    #  SHOW WINDOW  (blocks until closed)
    # -------------------------------------------------------------------------

    Write-BMLog -Message 'Launching BuildMaster window...' -Level Info
    try {
        $window.ShowDialog() | Out-Null
    }
    catch {
        # A fatal error escaped the WPF dispatcher - show it before the window
        # can turn into a zombie (dispatcher crash leaves the window visible
        # with no powershell backing it).
        $errMsg = $_.Exception.Message
        try {
            $uiTimer.Stop()
            Stop-BMEngine
        }
        catch { }
        try {
            [System.Windows.MessageBox]::Show(
                ("A fatal error occurred inside the BuildMaster window:`n`n{0}`n`n" +
                 "The application will now close.  Check the console for the full stack trace.") -f $errMsg,
                'BuildMaster - Fatal Error',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error) | Out-Null
            $window.Close()
        }
        catch { }
        throw   # re-throw so BuildMaster.ps1 can log/exit cleanly
    }
}

# -----------------------------------------------------------------------------
#  EXPORTS
# -----------------------------------------------------------------------------
Export-ModuleMember -Function @('Start-BMGui')
