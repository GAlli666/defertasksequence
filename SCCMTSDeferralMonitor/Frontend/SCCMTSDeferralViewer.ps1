<#
.SYNOPSIS
    SCCM Task Sequence Deferral Monitor - WPF Viewer

.DESCRIPTION
    WPF application that displays SCCM TS deferral monitoring data
    collected by the backend collector script.

.NOTES
    Date: 2025-11-28
    Requires: PowerShell 5.1 with WPF support
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#region Configuration

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load configuration from XML
function Load-ViewerConfig {
    $configPath = Join-Path $scriptDir "ViewerConfig.xml"

    if (Test-Path $configPath) {
        try {
            [xml]$config = Get-Content $configPath
            $script:DataPath = $config.ViewerConfiguration.Paths.DataPath
            $script:LogsPath = $config.ViewerConfiguration.Paths.LogsPath
            Write-Host "Configuration loaded from: $configPath"
        }
        catch {
            Write-Host "Error loading config, using defaults: $_"
            $script:DataPath = "C:\SCCMDeferralMonitor\Data"
            $script:LogsPath = "C:\SCCMDeferralMonitor\Logs"
        }
    }
    else {
        $script:DataPath = "C:\SCCMDeferralMonitor\Data"
        $script:LogsPath = "C:\SCCMDeferralMonitor\Logs"
    }
}

function Save-ViewerConfig {
    $configPath = Join-Path $scriptDir "ViewerConfig.xml"

    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<ViewerConfiguration>
    <Paths>
        <DataPath>$($script:DataPath)</DataPath>
        <LogsPath>$($script:LogsPath)</LogsPath>
    </Paths>
</ViewerConfiguration>
"@

    $xml | Out-File -FilePath $configPath -Encoding utf8 -Force
}

Load-ViewerConfig

$script:deviceData = @()
$script:metadata = @{}

#endregion

#region XAML Definition

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SCCM TS Deferral Monitor"
        Height="800" Width="1600"
        WindowStartupLocation="CenterScreen"
        Background="#FF2E3440">

    <Window.Resources>
        <SolidColorBrush x:Key="PrimaryBrush" Color="#FF5E81AC"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#FF88C0D0"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#FFA3BE8C"/>
        <SolidColorBrush x:Key="ErrorBrush" Color="#FFBF616A"/>
        <SolidColorBrush x:Key="BackgroundBrush" Color="#FF2E3440"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="#FF3B4252"/>
        <SolidColorBrush x:Key="TextBrush" Color="#FFECEFF4"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#FFD8DEE9"/>

        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource PrimaryBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="15,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBox">
            <Setter Property="Background" Value="#FF434C5E"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#FF4C566A"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton"
                                          Background="{TemplateBinding Background}"
                                          BorderBrush="{TemplateBinding BorderBrush}"
                                          BorderThickness="{TemplateBinding BorderThickness}"
                                          IsChecked="{Binding Path=IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}"
                                          ClickMode="Press">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="20"/>
                                    </Grid.ColumnDefinitions>
                                    <ContentPresenter Grid.Column="0"
                                                      Content="{TemplateBinding SelectionBoxItem}"
                                                      ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                                      ContentStringFormat="{TemplateBinding SelectionBoxItemStringFormat}"
                                                      HorizontalAlignment="Left"
                                                      VerticalAlignment="Center"
                                                      Margin="8,0,0,0"
                                                      TextElement.Foreground="{StaticResource TextBrush}"/>
                                    <Path Grid.Column="1"
                                          Data="M 0 0 L 4 4 L 8 0 Z"
                                          Fill="{StaticResource TextBrush}"
                                          HorizontalAlignment="Center"
                                          VerticalAlignment="Center"/>
                                </Grid>
                            </ToggleButton>
                            <Popup IsOpen="{TemplateBinding IsDropDownOpen}"
                                   Placement="Bottom"
                                   AllowsTransparency="True"
                                   PopupAnimation="Slide">
                                <Border Background="#FF434C5E"
                                        BorderBrush="#FF4C566A"
                                        BorderThickness="1"
                                        MaxHeight="200">
                                    <ScrollViewer VerticalScrollBarVisibility="Auto">
                                        <ItemsPresenter/>
                                    </ScrollViewer>
                                </Border>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="#FF434C5E"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Padding" Value="8,4"/>
            <Style.Triggers>
                <Trigger Property="IsHighlighted" Value="True">
                    <Setter Property="Background" Value="{StaticResource PrimaryBrush}"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
                <Trigger Property="IsSelected" Value="True">
                    <Setter Property="Background" Value="{StaticResource PrimaryBrush}"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{StaticResource PrimaryBrush}" Padding="20,15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0">
                    <TextBlock Text="SCCM TS Deferral Monitor" FontSize="22" FontWeight="Bold" Foreground="White"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <StackPanel Margin="0,0,30,0">
                        <TextBlock Text="Collection" FontSize="11" Foreground="White" Opacity="0.8"/>
                        <TextBlock x:Name="txtCollectionName" Text="Loading..." FontSize="13" FontWeight="Bold" Foreground="White"/>
                    </StackPanel>
                    <StackPanel>
                        <TextBlock Text="Task Sequence" FontSize="11" Foreground="White" Opacity="0.8"/>
                        <TextBlock x:Name="txtTSName" Text="Loading..." FontSize="13" FontWeight="Bold" Foreground="White"/>
                    </StackPanel>
                </StackPanel>
            </Grid>
        </Border>

        <Border Grid.Row="1" Background="{StaticResource SurfaceBrush}" Padding="20,10" BorderBrush="#FF4C566A" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Margin="0,0,30,0">
                    <TextBlock Text="LAST UPDATED" FontSize="10" Foreground="{StaticResource TextSecondaryBrush}" Opacity="0.7"/>
                    <TextBlock x:Name="txtLastUpdate" Text="-" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}"/>
                </StackPanel>

                <StackPanel Grid.Column="1" Margin="0,0,30,0">
                    <TextBlock Text="TOTAL DEVICES" FontSize="10" Foreground="{StaticResource TextSecondaryBrush}" Opacity="0.7"/>
                    <TextBlock x:Name="txtTotalDevices" Text="-" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}"/>
                </StackPanel>

                <StackPanel Grid.Column="2" Margin="0,0,30,0">
                    <TextBlock Text="ONLINE" FontSize="10" Foreground="{StaticResource TextSecondaryBrush}" Opacity="0.7"/>
                    <TextBlock x:Name="txtOnlineDevices" Text="-" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource SuccessBrush}"/>
                </StackPanel>

                <StackPanel Grid.Column="3" Margin="0,0,30,0">
                    <TextBlock Text="OFFLINE" FontSize="10" Foreground="{StaticResource TextSecondaryBrush}" Opacity="0.7"/>
                    <TextBlock x:Name="txtOfflineDevices" Text="-" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource ErrorBrush}"/>
                </StackPanel>

                <StackPanel Grid.Column="5" Orientation="Horizontal">
                    <Button x:Name="btnSettings" Content="⚙ Settings" Style="{StaticResource ModernButton}" Margin="5,0"/>
                    <Button x:Name="btnRefresh" Content="↻ Refresh" Style="{StaticResource ModernButton}" Background="{StaticResource AccentBrush}" Margin="5,0"/>
                </StackPanel>
            </Grid>
        </Border>

        <Border Grid.Row="2" Background="{StaticResource SurfaceBrush}" Padding="20,10" BorderBrush="#FF4C566A" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBox x:Name="txtSearch" Grid.Column="0" Height="32" Background="#FF434C5E" Foreground="{StaticResource TextBrush}"
                         BorderThickness="0" Padding="10,0" FontSize="13" VerticalContentAlignment="Center"/>

                <ComboBox x:Name="cmbStatusFilter" Grid.Column="1" Width="150" Height="32" Margin="10,0,0,0" FontSize="13"/>

                <ComboBox x:Name="cmbTSFilter" Grid.Column="2" Width="150" Height="32" Margin="10,0,0,0" FontSize="13"/>
            </Grid>
        </Border>

        <DataGrid x:Name="dgDevices" Grid.Row="3"
                  AutoGenerateColumns="False"
                  IsReadOnly="True"
                  GridLinesVisibility="Horizontal"
                  HeadersVisibility="Column"
                  SelectionMode="Single"
                  CanUserSortColumns="True"
                  Background="{StaticResource BackgroundBrush}"
                  Foreground="{StaticResource TextBrush}"
                  RowBackground="{StaticResource SurfaceBrush}"
                  AlternatingRowBackground="#FF373E4C"
                  BorderThickness="0"
                  HorizontalGridLinesBrush="#FF4C566A"
                  FontSize="13">

            <DataGrid.ColumnHeaderStyle>
                <Style TargetType="DataGridColumnHeader">
                    <Setter Property="Background" Value="#FF434C5E"/>
                    <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
                    <Setter Property="FontWeight" Value="SemiBold"/>
                    <Setter Property="Padding" Value="12,10"/>
                    <Setter Property="BorderThickness" Value="0,0,0,2"/>
                    <Setter Property="BorderBrush" Value="{StaticResource PrimaryBrush}"/>
                </Style>
            </DataGrid.ColumnHeaderStyle>

            <DataGrid.CellStyle>
                <Style TargetType="DataGridCell">
                    <Setter Property="BorderThickness" Value="0"/>
                    <Setter Property="Padding" Value="12,8"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="DataGridCell">
                                <Border Padding="{TemplateBinding Padding}" Background="{TemplateBinding Background}">
                                    <ContentPresenter VerticalAlignment="Center"/>
                                </Border>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </DataGrid.CellStyle>

            <DataGrid.Columns>
                <DataGridTextColumn Header="Device Name" Binding="{Binding DeviceName}" Width="180"/>
                <DataGridTextColumn Header="Primary User" Binding="{Binding PrimaryUser}" Width="140"/>

                <DataGridTemplateColumn Header="Status" Width="100">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <StackPanel Orientation="Horizontal">
                                <Ellipse Width="12" Height="12" Margin="0,0,8,0">
                                    <Ellipse.Style>
                                        <Style TargetType="Ellipse">
                                            <Setter Property="Fill" Value="#FFBF616A"/>
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding IsOnline}" Value="True">
                                                    <Setter Property="Fill" Value="#FFA3BE8C"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </Ellipse.Style>
                                </Ellipse>
                                <TextBlock Text="{Binding OnlineStatus}" VerticalAlignment="Center"/>
                            </StackPanel>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>

                <DataGridTextColumn Header="OS Name" Binding="{Binding OSName}" Width="100"/>
                <DataGridTextColumn Header="Build" Binding="{Binding OSBuildNumber}" Width="80"/>
                <DataGridTextColumn Header="TS Status" Binding="{Binding TSStatus}" Width="120"/>
                <DataGridTextColumn Header="Deferral Count" Binding="{Binding DeferralCount}" Width="110"/>
                <DataGridTextColumn Header="TS Trigger" Binding="{Binding TSTriggerAttempted}" Width="90"/>
                <DataGridTextColumn Header="Trigger Success" Binding="{Binding TSTriggerSuccess}" Width="110"/>

                <DataGridTemplateColumn Header="Actions" Width="200">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="btnTSStatus" Content="TS Status" Padding="10,5" Margin="0,0,5,0" FontSize="11" Tag="{Binding}">
                                    <Button.Style>
                                        <Style TargetType="Button" BasedOn="{StaticResource ModernButton}">
                                            <Setter Property="IsEnabled" Value="{Binding TSStatusMessagesAvailable}"/>
                                        </Style>
                                    </Button.Style>
                                </Button>
                                <Button x:Name="btnDeferralLog" Content="Deferral Log" Padding="10,5" FontSize="11" Tag="{Binding}">
                                    <Button.Style>
                                        <Style TargetType="Button" BasedOn="{StaticResource ModernButton}">
                                            <Setter Property="Background" Value="#FF6C757D"/>
                                            <Setter Property="IsEnabled" Value="{Binding LogAvailable}"/>
                                        </Style>
                                    </Button.Style>
                                </Button>
                            </StackPanel>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
            </DataGrid.Columns>
        </DataGrid>

        <Border Grid.Row="4" Background="{StaticResource SurfaceBrush}" Padding="20,8" BorderBrush="#FF4C566A" BorderThickness="0,1,0,0">
            <TextBlock x:Name="txtStatusBar" Text="Ready" FontSize="12" Foreground="{StaticResource TextSecondaryBrush}"/>
        </Border>
    </Grid>
</Window>
'@

#endregion

#region Functions

function Load-Data {
    try {
        $script:txtStatusBar.Text = "Loading data..."

        # Load metadata
        $metadataPath = Join-Path $script:DataPath "metadata.json"
        if (Test-Path $metadataPath) {
            $script:metadata = Get-Content $metadataPath | ConvertFrom-Json

            $script:txtLastUpdate.Text = $script:metadata.LastUpdate
            $script:txtTotalDevices.Text = $script:metadata.TotalDevices
            $script:txtOnlineDevices.Text = $script:metadata.OnlineDevices
            $script:txtOfflineDevices.Text = $script:metadata.OfflineDevices
            $script:txtCollectionName.Text = $script:metadata.CollectionName
            $script:txtTSName.Text = $script:metadata.TaskSequenceName
        }

        # Load device data
        $deviceDataPath = Join-Path $script:DataPath "devicedata.json"
        if (Test-Path $deviceDataPath) {
            $script:deviceData = Get-Content $deviceDataPath | ConvertFrom-Json

            Apply-Filters

            $script:txtStatusBar.Text = "Data loaded - $($script:deviceData.Count) devices"
        }
        else {
            $script:txtStatusBar.Text = "Error: devicedata.json not found at $deviceDataPath"
            [System.Windows.MessageBox]::Show("Device data file not found: $deviceDataPath", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
    catch {
        $script:txtStatusBar.Text = "Error loading data: $_"
        [System.Windows.MessageBox]::Show("Error loading data: $_", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
}

function Apply-Filters {
    $filtered = $script:deviceData

    # Search filter
    $searchText = $script:txtSearch.Text
    if (-not [string]::IsNullOrWhiteSpace($searchText)) {
        $filtered = $filtered | Where-Object {
            $_.DeviceName -like "*$searchText*" -or $_.PrimaryUser -like "*$searchText*"
        }
    }

    # Status filter
    $statusFilter = $script:cmbStatusFilter.SelectedItem
    if ($statusFilter -and $statusFilter.ToString() -ne "All Status") {
        if ($statusFilter -eq "Online Only") {
            $filtered = $filtered | Where-Object { $_.IsOnline -eq $true }
        }
        elseif ($statusFilter -eq "Offline Only") {
            $filtered = $filtered | Where-Object { $_.IsOnline -eq $false }
        }
    }

    # TS Status filter
    $tsFilter = $script:cmbTSFilter.SelectedItem
    if ($tsFilter -and $tsFilter.ToString() -ne "All TS Status") {
        $filtered = $filtered | Where-Object { $_.TSStatus -eq $tsFilter.ToString() }
    }

    $script:dgDevices.ItemsSource = $filtered
    $script:txtStatusBar.Text = "Showing $($filtered.Count) of $($script:deviceData.Count) devices"
}

function Show-SettingsDialog {
    $browser = New-Object System.Windows.Forms.FolderBrowserDialog
    $browser.Description = "Select Data Directory"
    $browser.SelectedPath = $script:DataPath

    if ($browser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:DataPath = $browser.SelectedPath
        $script:LogsPath = Join-Path (Split-Path $script:DataPath -Parent) "Logs"
        Save-ViewerConfig
        Load-Data
        [System.Windows.MessageBox]::Show("Configuration saved. Data Path: $($script:DataPath)`nLogs Path: $($script:LogsPath)", "Settings Updated", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
}

function Open-LogFile {
    param([string]$FilePath)

    if (Test-Path $FilePath) {
        Start-Process notepad.exe -ArgumentList $FilePath
    }
    else {
        [System.Windows.MessageBox]::Show("Log file not found: $FilePath", "File Not Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
    }
}

#endregion

#region Main Window

# Load XAML
$reader = New-Object System.IO.StringReader($xaml)
$xmlReader = [System.Xml.XmlReader]::Create($reader)
$window = [Windows.Markup.XamlReader]::Load($xmlReader)

# Get controls
$script:txtCollectionName = $window.FindName("txtCollectionName")
$script:txtTSName = $window.FindName("txtTSName")
$script:txtLastUpdate = $window.FindName("txtLastUpdate")
$script:txtTotalDevices = $window.FindName("txtTotalDevices")
$script:txtOnlineDevices = $window.FindName("txtOnlineDevices")
$script:txtOfflineDevices = $window.FindName("txtOfflineDevices")
$script:btnSettings = $window.FindName("btnSettings")
$script:btnRefresh = $window.FindName("btnRefresh")
$script:txtSearch = $window.FindName("txtSearch")
$script:cmbStatusFilter = $window.FindName("cmbStatusFilter")
$script:cmbTSFilter = $window.FindName("cmbTSFilter")
$script:dgDevices = $window.FindName("dgDevices")
$script:txtStatusBar = $window.FindName("txtStatusBar")

# Initialize combo boxes
$script:cmbStatusFilter.ItemsSource = @("All Status", "Online Only", "Offline Only")
$script:cmbStatusFilter.SelectedIndex = 0

$script:cmbTSFilter.ItemsSource = @("All TS Status", "Success", "Failed", "In Progress", "Not Started")
$script:cmbTSFilter.SelectedIndex = 0

# Event handlers
$script:btnRefresh.Add_Click({ Load-Data })
$script:btnSettings.Add_Click({ Show-SettingsDialog })
$script:txtSearch.Add_TextChanged({ Apply-Filters })
$script:cmbStatusFilter.Add_SelectionChanged({ Apply-Filters })
$script:cmbTSFilter.Add_SelectionChanged({ Apply-Filters })

# DataGrid button click handler using PreviewMouseLeftButtonDown
$script:dgDevices.Add_PreviewMouseLeftButtonDown({
    param($sender, $e)

    # Find if we clicked on a button
    $source = $e.OriginalSource
    $button = $null

    # Walk up the visual tree to find a button
    $current = $source
    while ($current -ne $null) {
        if ($current -is [System.Windows.Controls.Button]) {
            $button = $current
            break
        }
        if ($current -is [System.Windows.Media.Visual]) {
            $current = [System.Windows.Media.VisualTreeHelper]::GetParent($current)
        }
        else {
            break
        }
    }

    if ($button -and $button.Tag) {
        $device = $button.Tag

        if ($button.Content -eq "TS Status") {
            # Try HTML file first, then JSON
            $htmlPath = Join-Path $script:LogsPath ($device.DeviceName + "_TSExecutionLog.html")
            $jsonPath = Join-Path $script:LogsPath $device.TSStatusMessagesPath

            if (Test-Path $htmlPath) {
                Start-Process $htmlPath
            }
            elseif (Test-Path $jsonPath) {
                Open-LogFile $jsonPath
            }
            else {
                [System.Windows.MessageBox]::Show("TS execution log not found for $($device.DeviceName)", "File Not Found", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
            }
        }
        elseif ($button.Content -eq "Deferral Log") {
            $logPath = Join-Path $script:LogsPath $device.DeferralLogPath
            Open-LogFile $logPath
        }
    }
})

# Load initial data
Load-Data

# Show window
$window.ShowDialog() | Out-Null

#endregion
