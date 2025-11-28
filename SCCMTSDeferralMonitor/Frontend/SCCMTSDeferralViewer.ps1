<#
.SYNOPSIS
    SCCM Task Sequence Deferral Monitor - WPF Viewer

.DESCRIPTION
    WPF application that displays SCCM TS deferral monitoring data
    collected by the backend collector script.

.NOTES
    Date: 2025-11-28
    Requires: PowerShell 5.1 with WPF support

    Features:
    - Real-time data refresh from JSON files
    - Sortable columns
    - Filter by online/offline status
    - Filter by TS status
    - Search by device name or user
    - View TS execution logs
    - View deferral logs
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

#region Configuration

# Default data path (can be changed in UI)
$script:DataPath = "C:\SCCMDeferralMonitor\Data"
$script:LogsPath = "C:\SCCMDeferralMonitor\Logs"
$script:deviceData = @()
$script:metadata = @{}

#endregion

#region XAML Definition

$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SCCM TS Deferral Monitor"
        Height="800" Width="1600"
        WindowStartupLocation="CenterScreen"
        Background="#FF2E3440">

    <Window.Resources>
        <!-- Modern Color Scheme -->
        <SolidColorBrush x:Key="PrimaryBrush" Color="#FF5E81AC"/>
        <SolidColorBrush x:Key="SecondaryBrush" Color="#FF81A1C1"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#FF88C0D0"/>
        <SolidColorBrush x:Key="SuccessBrush" Color="#FFA3BE8C"/>
        <SolidColorBrush x:Key="WarningBrush" Color="#FFEBCB8B"/>
        <SolidColorBrush x:Key="ErrorBrush" Color="#FFBF616A"/>
        <SolidColorBrush x:Key="BackgroundBrush" Color="#FF2E3440"/>
        <SolidColorBrush x:Key="SurfaceBrush" Color="#FF3B4252"/>
        <SolidColorBrush x:Key="TextBrush" Color="#FFECEFF4"/>
        <SolidColorBrush x:Key="TextSecondaryBrush" Color="#FFD8DEE9"/>

        <!-- Button Style -->
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
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="{StaticResource SecondaryBrush}"/>
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

        <!-- Header -->
        <Border Grid.Row="0" Background="{StaticResource PrimaryBrush}" Padding="20,15">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0">
                    <TextBlock Text="SCCM TS Deferral Monitor"
                               FontSize="22" FontWeight="Bold"
                               Foreground="White"/>
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

        <!-- Metadata Bar -->
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

        <!-- Filters -->
        <Border Grid.Row="2" Background="{StaticResource SurfaceBrush}" Padding="20,10" BorderBrush="#FF4C566A" BorderThickness="0,0,0,1">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <TextBox x:Name="txtSearch" Grid.Column="0"
                         Height="32"
                         Background="#FF434C5E"
                         Foreground="{StaticResource TextBrush}"
                         BorderThickness="0"
                         Padding="10,0"
                         FontSize="13"
                         VerticalContentAlignment="Center">
                    <TextBox.Resources>
                        <Style TargetType="Border">
                            <Setter Property="CornerRadius" Value="4"/>
                        </Style>
                    </TextBox.Resources>
                </TextBox>
                <TextBlock IsHitTestVisible="False"
                           Text="Search by device name or user..."
                           VerticalAlignment="Center"
                           HorizontalAlignment="Left"
                           Margin="15,0,0,0"
                           Foreground="{StaticResource TextSecondaryBrush}"
                           Opacity="0.6"
                           FontSize="13">
                    <TextBlock.Style>
                        <Style TargetType="TextBlock">
                            <Setter Property="Visibility" Value="Collapsed"/>
                            <Style.Triggers>
                                <DataTrigger Binding="{Binding Text, ElementName=txtSearch}" Value="">
                                    <Setter Property="Visibility" Value="Visible"/>
                                </DataTrigger>
                            </Style.Triggers>
                        </Style>
                    </TextBlock.Style>
                </TextBlock>

                <ComboBox x:Name="cmbStatusFilter" Grid.Column="1"
                          Width="150" Height="32" Margin="10,0,0,0"
                          Background="#FF434C5E"
                          Foreground="{StaticResource TextBrush}"
                          BorderThickness="0"
                          FontSize="13"/>

                <ComboBox x:Name="cmbTSFilter" Grid.Column="2"
                          Width="150" Height="32" Margin="10,0,0,0"
                          Background="#FF434C5E"
                          Foreground="{StaticResource TextBrush}"
                          BorderThickness="0"
                          FontSize="13"/>
            </Grid>
        </Border>

        <!-- DataGrid -->
        <DataGrid x:Name="dgDevices" Grid.Row="3"
                  AutoGenerateColumns="False"
                  IsReadOnly="True"
                  GridLinesVisibility="Horizontal"
                  HeadersVisibility="Column"
                  SelectionMode="Single"
                  CanUserSortColumns="True"
                  CanUserResizeColumns="True"
                  CanUserReorderColumns="False"
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
                    <Setter Property="HorizontalContentAlignment" Value="Left"/>
                </Style>
            </DataGrid.ColumnHeaderStyle>

            <DataGrid.CellStyle>
                <Style TargetType="DataGridCell">
                    <Setter Property="BorderThickness" Value="0"/>
                    <Setter Property="Padding" Value="12,8"/>
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="DataGridCell">
                                <Border Padding="{TemplateBinding Padding}"
                                        Background="{TemplateBinding Background}">
                                    <ContentPresenter VerticalAlignment="Center"/>
                                </Border>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </DataGrid.CellStyle>

            <DataGrid.Columns>
                <DataGridTextColumn Header="Device Name" Binding="{Binding DeviceName}" Width="200"/>
                <DataGridTextColumn Header="Primary User" Binding="{Binding PrimaryUser}" Width="150"/>
                <DataGridTextColumn Header="Online Status" Binding="{Binding OnlineStatus}" Width="100"/>
                <DataGridTextColumn Header="TS Status" Binding="{Binding TSStatus}" Width="120"/>
                <DataGridTextColumn Header="Deferral Count" Binding="{Binding DeferralCount}" Width="120"/>
                <DataGridTextColumn Header="TS Trigger" Binding="{Binding TSTriggerAttempted}" Width="100"/>
                <DataGridTextColumn Header="Trigger Success" Binding="{Binding TSTriggerSuccess}" Width="120"/>
                <DataGridTemplateColumn Header="Actions" Width="200">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <StackPanel Orientation="Horizontal">
                                <Button x:Name="btnViewTSStatus" Content="TS Status"
                                        Style="{StaticResource ModernButton}"
                                        Background="{StaticResource PrimaryBrush}"
                                        Padding="10,5" Margin="0,0,5,0"
                                        FontSize="11"
                                        Tag="{Binding}"/>
                                <Button x:Name="btnViewDeferralLog" Content="Deferral Log"
                                        Style="{StaticResource ModernButton}"
                                        Background="#FF6C757D"
                                        Padding="10,5"
                                        FontSize="11"
                                        Tag="{Binding}"/>
                            </StackPanel>
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
            </DataGrid.Columns>
        </DataGrid>

        <!-- Status Bar -->
        <Border Grid.Row="4" Background="{StaticResource SurfaceBrush}" Padding="20,8" BorderBrush="#FF4C566A" BorderThickness="0,1,0,0">
            <TextBlock x:Name="txtStatusBar"
                       Text="Ready"
                       FontSize="12"
                       Foreground="{StaticResource TextSecondaryBrush}"/>
        </Border>
    </Grid>
</Window>
"@

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

            # Apply filters and update grid
            Apply-Filters

            $script:txtStatusBar.Text = "Data loaded successfully - $($script:deviceData.Count) devices"
        }
        else {
            $script:txtStatusBar.Text = "Error: devicedata.json not found"
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
    if ($statusFilter -and $statusFilter -ne "All Status") {
        if ($statusFilter -eq "Online Only") {
            $filtered = $filtered | Where-Object { $_.IsOnline -eq $true }
        }
        elseif ($statusFilter -eq "Offline Only") {
            $filtered = $filtered | Where-Object { $_.IsOnline -eq $false }
        }
    }

    # TS Status filter
    $tsFilter = $script:cmbTSFilter.SelectedItem
    if ($tsFilter -and $tsFilter -ne "All TS Status") {
        $filtered = $filtered | Where-Object { $_.TSStatus -eq $tsFilter }
    }

    $script:dgDevices.ItemsSource = $filtered
    $script:txtStatusBar.Text = "Showing $($filtered.Count) of $($script:deviceData.Count) devices"
}

function Show-SettingsDialog {
    $settingsXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Settings" Height="250" Width="500"
        WindowStartupLocation="CenterOwner"
        Background="#FF2E3440"
        ResizeMode="NoResize">
    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Data Path" Foreground="#FFECEFF4" FontSize="13" FontWeight="SemiBold" Margin="0,0,0,5"/>
        <TextBox x:Name="txtDataPath" Grid.Row="1" Height="32" Background="#FF434C5E" Foreground="#FFECEFF4" BorderThickness="0" Padding="10,0" FontSize="13"/>

        <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right">
            <Button x:Name="btnOK" Content="OK" Width="80" Height="32" Margin="0,0,10,0"/>
            <Button x:Name="btnCancel" Content="Cancel" Width="80" Height="32"/>
        </StackPanel>
    </Grid>
</Window>
"@

    $settingsReader = [System.Xml.XmlNodeReader]::new([xml]$settingsXaml)
    $settingsWindow = [Windows.Markup.XamlReader]::Load($settingsReader)

    $txtDataPath = $settingsWindow.FindName("txtDataPath")
    $btnOK = $settingsWindow.FindName("btnOK")
    $btnCancel = $settingsWindow.FindName("btnCancel")

    $txtDataPath.Text = $script:DataPath

    $btnOK.Add_Click({
        $script:DataPath = $txtDataPath.Text
        $settingsWindow.Close()
        Load-Data
    })

    $btnCancel.Add_Click({
        $settingsWindow.Close()
    })

    $settingsWindow.ShowDialog()
}

function Open-LogFile {
    param([string]$FilePath)

    if (Test-Path $FilePath) {
        Start-Process notepad.exe -ArgumentList $FilePath
    }
    else {
        [System.Windows.MessageBox]::Show("Log file not found: $FilePath", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
    }
}

#endregion

#region Main Window

# Load XAML
$reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

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

# DataGrid button click handling
$script:dgDevices.Add_LoadingRow({
    param($sender, $e)

    $row = $e.Row
    $device = $row.Item

    # Find buttons in the row
    $btnTSStatus = [System.Windows.LogicalTreeHelper]::FindLogicalNode($row, "btnViewTSStatus")
    $btnDeferralLog = [System.Windows.LogicalTreeHelper]::FindLogicalNode($row, "btnViewDeferralLog")

    if ($btnTSStatus) {
        $btnTSStatus.Add_Click({
            param($btnSender, $btnE)
            $deviceData = $btnSender.Tag
            $logPath = Join-Path $script:LogsPath $deviceData.TSStatusMessagesPath
            if ($deviceData.TSStatusMessagesAvailable) {
                Open-LogFile $logPath
            }
            else {
                [System.Windows.MessageBox]::Show("No TS status messages available for this device.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        })
    }

    if ($btnDeferralLog) {
        $btnDeferralLog.Add_Click({
            param($btnSender, $btnE)
            $deviceData = $btnSender.Tag
            $logPath = Join-Path $script:LogsPath $deviceData.DeferralLogPath
            if ($deviceData.LogAvailable) {
                Open-LogFile $logPath
            }
            else {
                [System.Windows.MessageBox]::Show("No deferral log available for this device.", "Information", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
            }
        })
    }
})

# Load initial data
Load-Data

# Show window
$window.ShowDialog() | Out-Null

#endregion
