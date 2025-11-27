<#
.SYNOPSIS
    SCCM Task Sequence Deferral Log Reader Diagnostic Tool

.DESCRIPTION
    Connects to SCCM, retrieves collection members, and reads TaskSequenceDeferral.log files
    from each member via UNC path. Displays deferral status, trigger attempts, and success/failure
    in a modern WPF UI with traffic light indicators.

.NOTES
    Date: 2025-11-27
    Requires: PowerShell 5.1, ConfigurationManager PowerShell Module, SCCM Admin Rights
#>

[CmdletBinding()]
param()

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

#region Data Classes

class ComputerStatus {
    [string]$ComputerName
    [bool]$IsOnline
    [string]$Status
    [int]$DeferralCount
    [string]$DeferralStatus
    [bool]$TSTriggerAttempted
    [bool]$TSTriggerSuccess
    [string]$TSPackageID
    [string]$LastDeferralDate
    [string]$LastTriggerDate
    [string]$ErrorMessage
    [System.Collections.ArrayList]$LogEntries

    ComputerStatus() {
        $this.LogEntries = New-Object System.Collections.ArrayList
        $this.IsOnline = $false
        $this.Status = "Pending"
        $this.DeferralCount = 0
        $this.TSTriggerAttempted = $false
        $this.TSTriggerSuccess = $false
        $this.TSPackageID = "N/A"
        $this.DeferralStatus = "Unknown"
        $this.LastDeferralDate = "N/A"
        $this.LastTriggerDate = "N/A"
        $this.ErrorMessage = ""
    }
}

class LogEntry {
    [datetime]$Timestamp
    [string]$Level
    [string]$Message

    LogEntry([datetime]$timestamp, [string]$level, [string]$message) {
        $this.Timestamp = $timestamp
        $this.Level = $level
        $this.Message = $message
    }
}

#endregion

#region Log Parsing Functions

function Test-ComputerOnline {
    param(
        [string]$ComputerName,
        [int]$TimeoutMs = 1000
    )

    try {
        $ping = New-Object System.Net.NetworkInformation.Ping
        $result = $ping.Send($ComputerName, $TimeoutMs)
        return ($result.Status -eq 'Success')
    }
    catch {
        return $false
    }
}

function Parse-LogFile {
    param(
        [string]$LogPath,
        [string]$ComputerName
    )

    $status = [ComputerStatus]::new()
    $status.ComputerName = $ComputerName

    # First check if computer is online
    Write-Host "Checking if $ComputerName is online..." -ForegroundColor Cyan
    $status.IsOnline = Test-ComputerOnline -ComputerName $ComputerName

    if (-not $status.IsOnline) {
        $status.Status = "Offline"
        $status.ErrorMessage = "Computer is offline or unreachable"
        Write-Host "$ComputerName is OFFLINE" -ForegroundColor Red
        return $status
    }

    Write-Host "$ComputerName is ONLINE" -ForegroundColor Green

    try {
        # Build UNC path
        $uncPath = "\\$ComputerName\$($LogPath.Replace(':', '$'))"

        Write-Host "Reading log from: $uncPath" -ForegroundColor Cyan

        if (-not (Test-Path $uncPath)) {
            $status.Status = "No Log File"
            $status.ErrorMessage = "Log file not found"
            return $status
        }

        # Read log file
        $logContent = Get-Content -Path $uncPath -ErrorAction Stop

        if ($logContent.Count -eq 0) {
            $status.Status = "Empty Log"
            $status.ErrorMessage = "Log file is empty"
            return $status
        }

        # Parse each line
        $deferrals = 0
        $triggerAttempted = $false
        $triggerSuccess = $false
        $lastDeferralDate = $null
        $lastTriggerDate = $null
        $packageID = "N/A"
        $resetDetected = $false

        foreach ($line in $logContent) {
            # Parse log line format: [2025-11-27 10:13:45] [Info] Message
            if ($line -match '^\[([\d-]+\s+[\d:]+)\]\s+\[(\w+)\]\s+(.+)$') {
                $timestamp = [datetime]::ParseExact($matches[1], "yyyy-MM-dd HH:mm:ss", $null)
                $level = $matches[2]
                $message = $matches[3]

                $entry = [LogEntry]::new($timestamp, $level, $message)
                $status.LogEntries.Add($entry) | Out-Null

                # Check for deferral increment
                if ($message -match 'Deferral count incremented immediately:\s+(\d+)\s+/\s+\d+') {
                    $deferrals = [int]$matches[1]
                    $lastDeferralDate = $timestamp
                }

                # Check for deferral reset
                if ($message -match 'Deferral count reset to 0') {
                    $resetDetected = $true
                    $deferrals = 0
                }

                # Check for user choice to defer
                if ($message -match 'User chose to defer') {
                    $lastDeferralDate = $timestamp
                }

                # Check for TS trigger attempt
                if ($message -match 'Attempting to start Task Sequence:\s+(\w+)') {
                    $triggerAttempted = $true
                    $packageID = $matches[1]
                    $lastTriggerDate = $timestamp
                }

                # Check for TS trigger success
                if ($message -match 'Task Sequence triggered successfully!') {
                    $triggerSuccess = $true
                }

                # Check for TS trigger failure
                if ($message -match 'Failed to (start Task Sequence|trigger schedule)') {
                    $triggerSuccess = $false
                }

                # Check for TS started successfully
                if ($message -match 'Task Sequence started successfully') {
                    $triggerSuccess = $true
                }
            }
        }

        # Set final status
        $status.DeferralCount = $deferrals
        $status.TSTriggerAttempted = $triggerAttempted
        $status.TSTriggerSuccess = $triggerSuccess
        $status.TSPackageID = $packageID

        if ($lastDeferralDate) {
            $status.LastDeferralDate = $lastDeferralDate.ToString("yyyy-MM-dd HH:mm:ss")
        }

        if ($lastTriggerDate) {
            $status.LastTriggerDate = $lastTriggerDate.ToString("yyyy-MM-dd HH:mm:ss")
        }

        # Determine deferral status
        if ($triggerAttempted) {
            if ($triggerSuccess) {
                $status.DeferralStatus = "TS Started"
                $status.Status = "Success"
            } else {
                $status.DeferralStatus = "TS Failed"
                $status.Status = "Failed"
            }
        } elseif ($deferrals -gt 0) {
            $status.DeferralStatus = "Deferred ($deferrals)"
            $status.Status = "Deferred"
        } else {
            $status.DeferralStatus = "No Deferrals"
            $status.Status = "Ready"
        }

    }
    catch {
        $status.Status = "Error"
        $status.ErrorMessage = $_.Exception.Message
        Write-Host "Error reading log for $ComputerName : $_" -ForegroundColor Red
    }

    return $status
}

#endregion

#region SCCM Functions

function Connect-ToSCCM {
    param(
        [string]$SiteCode,
        [string]$SiteServer
    )

    try {
        Write-Host "Connecting to SCCM Site: $SiteCode on $SiteServer" -ForegroundColor Cyan

        # Import ConfigurationManager module
        if (-not (Get-Module -Name ConfigurationManager)) {
            Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" -ErrorAction Stop
        }

        # Connect to site
        $siteCodePath = "$SiteCode" + ":"
        if (-not (Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue)) {
            New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop | Out-Null
        }

        # Store site code in script variable for later use
        $script:sccmSiteCode = $SiteCode

        Write-Host "Successfully connected to SCCM site: $SiteCode" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect to SCCM: $_" -ForegroundColor Red
        return $false
    }
}

function Get-CollectionMembers {
    param(
        [string]$CollectionID
    )

    try {
        Write-Host "Retrieving members of collection: $CollectionID" -ForegroundColor Cyan

        # Store current location and switch to SCCM drive
        $currentLocation = Get-Location
        $siteCodePath = "$($script:sccmSiteCode):"

        Set-Location $siteCodePath -ErrorAction Stop
        Write-Host "Switched to SCCM drive: $siteCodePath" -ForegroundColor Cyan

        # Get collection members
        $members = Get-CMCollectionMember -CollectionId $CollectionID -ErrorAction Stop

        # Restore original location
        Set-Location $currentLocation

        if ($members) {
            Write-Host "Found $($members.Count) members in collection" -ForegroundColor Green
            return $members
        } else {
            Write-Host "No members found in collection" -ForegroundColor Yellow
            return @()
        }
    }
    catch {
        Write-Host "Failed to retrieve collection members: $_" -ForegroundColor Red
        # Try to restore location even on error
        try {
            if ($currentLocation) {
                Set-Location $currentLocation
            }
        } catch {
            # Silently ignore restore errors
        }
        return $null
    }
}

#endregion

#region UI Functions

function Show-MainWindow {
    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="SCCM Task Sequence Deferral Log Reader"
    Height="900" Width="1500"
    WindowStartupLocation="CenterScreen"
    Background="#F5F7FA">

    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#5DADE2"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="20,10"/>
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
                    <Setter Property="Background" Value="#3498DB"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#2980B9"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#BDC3C7"/>
                    <Setter Property="Cursor" Value="Arrow"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="#BDC3C7"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="White"/>
        </Style>

        <Style x:Key="ModernLabel" TargetType="Label">
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="#2C3E50"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header -->
        <Border Grid.Row="0" Background="#5DADE2" CornerRadius="8" Padding="20" Margin="0,0,0,20">
            <StackPanel>
                <TextBlock Text="SCCM Task Sequence Deferral Log Reader"
                           FontSize="28"
                           FontWeight="Bold"
                           Foreground="White"
                           HorizontalAlignment="Center"/>
                <TextBlock Text="Diagnostic Tool for Collection Member Status"
                           FontSize="14"
                           Foreground="White"
                           Opacity="0.9"
                           HorizontalAlignment="Center"
                           Margin="0,5,0,0"/>
            </StackPanel>
        </Border>

        <!-- Connection Settings -->
        <Border Grid.Row="1" Background="White" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="8" Padding="20" Margin="0,0,0,20">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="150"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="150"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Row 1: Site Code and Site Server -->
                <Label Grid.Row="0" Grid.Column="0" Content="Site Code:" Style="{StaticResource ModernLabel}" VerticalAlignment="Center"/>
                <TextBox Grid.Row="0" Grid.Column="1" Name="txtSiteCode" Style="{StaticResource ModernTextBox}" Margin="0,5,20,5"/>

                <Label Grid.Row="0" Grid.Column="2" Content="Site Server:" Style="{StaticResource ModernLabel}" VerticalAlignment="Center"/>
                <TextBox Grid.Row="0" Grid.Column="3" Name="txtSiteServer" Style="{StaticResource ModernTextBox}" Margin="0,5,20,5"/>

                <Button Grid.Row="0" Grid.Column="4" Name="btnConnect" Content="Connect" Style="{StaticResource ModernButton}" Width="120" Margin="0,5,0,5"/>

                <!-- Row 2: Collection ID and Log Path -->
                <Label Grid.Row="1" Grid.Column="0" Content="Collection ID:" Style="{StaticResource ModernLabel}" VerticalAlignment="Center"/>
                <TextBox Grid.Row="1" Grid.Column="1" Name="txtCollectionID" Style="{StaticResource ModernTextBox}" Margin="0,5,20,5"/>

                <Label Grid.Row="1" Grid.Column="2" Content="Log File Path:" Style="{StaticResource ModernLabel}" VerticalAlignment="Center"/>
                <TextBox Grid.Row="1" Grid.Column="3" Name="txtLogPath" Style="{StaticResource ModernTextBox}" Text="C:\Windows\ccm\logs\TaskSequenceDeferral.log" Margin="0,5,20,5"/>

                <Button Grid.Row="1" Grid.Column="4" Name="btnScan" Content="Scan Collection" Style="{StaticResource ModernButton}" Width="120" Margin="0,5,0,5" IsEnabled="False"/>

                <!-- Row 3: Status and Progress -->
                <Label Grid.Row="2" Grid.Column="0" Content="Status:" Style="{StaticResource ModernLabel}" VerticalAlignment="Center"/>
                <TextBlock Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="3" Name="txtStatus"
                           Text="Not connected"
                           FontSize="14"
                           Foreground="#7F8C8D"
                           VerticalAlignment="Center"
                           Margin="0,5,0,5"/>
                <ProgressBar Grid.Row="2" Grid.Column="4" Name="progressBar"
                             Height="20"
                             Width="120"
                             Margin="0,5,0,5"
                             Visibility="Collapsed"/>
            </Grid>
        </Border>

        <!-- Data Grid -->
        <Border Grid.Row="2" Background="White" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="8" Padding="10">
            <DataGrid Name="dataGrid"
                      AutoGenerateColumns="False"
                      IsReadOnly="True"
                      GridLinesVisibility="Horizontal"
                      HeadersVisibility="Column"
                      RowHeight="35"
                      FontSize="12"
                      AlternatingRowBackground="#F8F9FA"
                      BorderThickness="0">
                <DataGrid.RowStyle>
                    <Style TargetType="DataGridRow">
                        <Style.Triggers>
                            <DataTrigger Binding="{Binding IsOnline}" Value="False">
                                <Setter Property="BorderBrush" Value="#E74C3C"/>
                                <Setter Property="BorderThickness" Value="2"/>
                                <Setter Property="Background" Value="#FADBD8"/>
                            </DataTrigger>
                        </Style.Triggers>
                    </Style>
                </DataGrid.RowStyle>
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Computer Name" Binding="{Binding ComputerName}" Width="150"/>
                    <DataGridTemplateColumn Header="Online" Width="80">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <Ellipse Width="16" Height="16" HorizontalAlignment="Center">
                                    <Ellipse.Style>
                                        <Style TargetType="Ellipse">
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding IsOnline}" Value="True">
                                                    <Setter Property="Fill" Value="#2ECC71"/>
                                                </DataTrigger>
                                                <DataTrigger Binding="{Binding IsOnline}" Value="False">
                                                    <Setter Property="Fill" Value="#E74C3C"/>
                                                </DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </Ellipse.Style>
                                </Ellipse>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                    <DataGridTemplateColumn Header="Deferral Count" Width="120">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                    <Ellipse Width="16" Height="16" Margin="0,0,8,0">
                                        <Ellipse.Style>
                                            <Style TargetType="Ellipse">
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding DeferralCount}" Value="0">
                                                        <Setter Property="Fill" Value="#2ECC71"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding DeferralCount}" Value="1">
                                                        <Setter Property="Fill" Value="#F39C12"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding DeferralCount}" Value="2">
                                                        <Setter Property="Fill" Value="#F39C12"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding DeferralCount}" Value="3">
                                                        <Setter Property="Fill" Value="#E74C3C"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                                <Setter Property="Fill" Value="#E74C3C"/>
                                            </Style>
                                        </Ellipse.Style>
                                    </Ellipse>
                                    <TextBlock Text="{Binding DeferralCount}" VerticalAlignment="Center"/>
                                </StackPanel>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="Deferral Status" Binding="{Binding DeferralStatus}" Width="120"/>
                    <DataGridTemplateColumn Header="TS Trigger Attempted" Width="140">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                    <Ellipse Width="16" Height="16">
                                        <Ellipse.Style>
                                            <Style TargetType="Ellipse">
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding TSTriggerAttempted}" Value="True">
                                                        <Setter Property="Fill" Value="#2ECC71"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding TSTriggerAttempted}" Value="False">
                                                        <Setter Property="Fill" Value="#95A5A6"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </Ellipse.Style>
                                    </Ellipse>
                                    <TextBlock Text="{Binding TSTriggerAttempted}" VerticalAlignment="Center" Margin="8,0,0,0"/>
                                </StackPanel>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTemplateColumn Header="TS Trigger Success" Width="130">
                        <DataGridTemplateColumn.CellTemplate>
                            <DataTemplate>
                                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                                    <Ellipse Width="16" Height="16">
                                        <Ellipse.Style>
                                            <Style TargetType="Ellipse">
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding TSTriggerSuccess}" Value="True">
                                                        <Setter Property="Fill" Value="#2ECC71"/>
                                                    </DataTrigger>
                                                    <DataTrigger Binding="{Binding TSTriggerSuccess}" Value="False">
                                                        <Setter Property="Fill" Value="#E74C3C"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </Ellipse.Style>
                                    </Ellipse>
                                    <TextBlock Text="{Binding TSTriggerSuccess}" VerticalAlignment="Center" Margin="8,0,0,0"/>
                                </StackPanel>
                            </DataTemplate>
                        </DataGridTemplateColumn.CellTemplate>
                    </DataGridTemplateColumn>
                    <DataGridTextColumn Header="TS Package ID" Binding="{Binding TSPackageID}" Width="120"/>
                    <DataGridTextColumn Header="Last Deferral" Binding="{Binding LastDeferralDate}" Width="150"/>
                    <DataGridTextColumn Header="Last TS Trigger" Binding="{Binding LastTriggerDate}" Width="150"/>
                    <DataGridTextColumn Header="Error" Binding="{Binding ErrorMessage}" Width="*"/>
                </DataGrid.Columns>
            </DataGrid>
        </Border>

        <!-- Footer with Legend and Export -->
        <Border Grid.Row="3" Background="White" BorderBrush="#BDC3C7" BorderThickness="1" CornerRadius="8" Padding="15" Margin="0,20,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <!-- Legend -->
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <TextBlock Text="Legend:" FontWeight="SemiBold" Margin="0,0,20,0" VerticalAlignment="Center"/>
                    <Ellipse Width="16" Height="16" Fill="#2ECC71" Margin="0,0,5,0"/>
                    <TextBlock Text="Good/Success" Margin="0,0,20,0" VerticalAlignment="Center"/>
                    <Ellipse Width="16" Height="16" Fill="#F39C12" Margin="0,0,5,0"/>
                    <TextBlock Text="Warning (1-2 deferrals)" Margin="0,0,20,0" VerticalAlignment="Center"/>
                    <Ellipse Width="16" Height="16" Fill="#E74C3C" Margin="0,0,5,0"/>
                    <TextBlock Text="Critical (3+ deferrals or Failed)" Margin="0,0,20,0" VerticalAlignment="Center"/>
                    <Ellipse Width="16" Height="16" Fill="#95A5A6" Margin="0,0,5,0"/>
                    <TextBlock Text="Not Attempted" Margin="0,0,20,0" VerticalAlignment="Center"/>
                    <Border BorderBrush="#E74C3C" BorderThickness="2" Width="16" Height="16" Background="#FADBD8" Margin="0,0,5,0"/>
                    <TextBlock Text="Offline" VerticalAlignment="Center"/>
                </StackPanel>

                <!-- Export Button -->
                <Button Grid.Column="1" Name="btnExport" Content="Export to CSV" Style="{StaticResource ModernButton}" Width="140"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    # Load XAML
    $reader = New-Object System.Xml.XmlNodeReader($xaml)
    $script:window = [Windows.Markup.XamlReader]::Load($reader)

    # Get controls
    $script:txtSiteCode = $window.FindName("txtSiteCode")
    $script:txtSiteServer = $window.FindName("txtSiteServer")
    $script:txtCollectionID = $window.FindName("txtCollectionID")
    $script:txtLogPath = $window.FindName("txtLogPath")
    $script:txtStatus = $window.FindName("txtStatus")
    $script:btnConnect = $window.FindName("btnConnect")
    $script:btnScan = $window.FindName("btnScan")
    $script:btnExport = $window.FindName("btnExport")
    $script:dataGrid = $window.FindName("dataGrid")
    $script:progressBar = $window.FindName("progressBar")

    # Initialize data collection
    $script:computerStatuses = New-Object System.Collections.ObjectModel.ObservableCollection[Object]
    $dataGrid.ItemsSource = $script:computerStatuses

    # Connect button click
    $btnConnect.Add_Click({
        $siteCode = $txtSiteCode.Text.Trim()
        $siteServer = $txtSiteServer.Text.Trim()

        if ([string]::IsNullOrEmpty($siteCode) -or [string]::IsNullOrEmpty($siteServer)) {
            [System.Windows.MessageBox]::Show("Please enter both Site Code and Site Server", "Input Required", "OK", "Warning")
            return
        }

        $txtStatus.Text = "Connecting to SCCM..."
        $btnConnect.IsEnabled = $false

        if (Connect-ToSCCM -SiteCode $siteCode -SiteServer $siteServer) {
            $txtStatus.Text = "Connected to SCCM site: $siteCode"
            $txtStatus.Foreground = "#2ECC71"
            $btnScan.IsEnabled = $true
        } else {
            $txtStatus.Text = "Failed to connect to SCCM"
            $txtStatus.Foreground = "#E74C3C"
            $btnConnect.IsEnabled = $true
        }
    })

    # Scan button click
    $btnScan.Add_Click({
        $collectionID = $txtCollectionID.Text.Trim()
        $logPath = $txtLogPath.Text.Trim()

        if ([string]::IsNullOrEmpty($collectionID)) {
            [System.Windows.MessageBox]::Show("Please enter a Collection ID", "Input Required", "OK", "Warning")
            return
        }

        if ([string]::IsNullOrEmpty($logPath)) {
            [System.Windows.MessageBox]::Show("Please enter a Log File Path", "Input Required", "OK", "Warning")
            return
        }

        # Clear previous results
        $script:computerStatuses.Clear()

        $txtStatus.Text = "Retrieving collection members..."
        $btnScan.IsEnabled = $false
        $progressBar.Visibility = "Visible"
        $progressBar.IsIndeterminate = $true

        # Get collection members
        $members = Get-CollectionMembers -CollectionID $collectionID

        if ($null -eq $members) {
            $txtStatus.Text = "Failed to retrieve collection members"
            $txtStatus.Foreground = "#E74C3C"
            $btnScan.IsEnabled = $true
            $progressBar.Visibility = "Collapsed"
            return
        }

        if ($members.Count -eq 0) {
            $txtStatus.Text = "No members found in collection"
            $txtStatus.Foreground = "#F39C12"
            $btnScan.IsEnabled = $true
            $progressBar.Visibility = "Collapsed"
            return
        }

        $txtStatus.Text = "Scanning $($members.Count) computers..."
        $progressBar.IsIndeterminate = $false
        $progressBar.Maximum = $members.Count
        $progressBar.Value = 0

        # Process each member
        $count = 0
        foreach ($member in $members) {
            $count++
            $computerName = $member.Name

            Write-Host "`nProcessing $count/$($members.Count): $computerName" -ForegroundColor Yellow

            $status = Parse-LogFile -LogPath $logPath -ComputerName $computerName
            $script:computerStatuses.Add($status)

            $progressBar.Value = $count
            $txtStatus.Text = "Scanned $count / $($members.Count) computers..."

            # Allow UI to update
            [System.Windows.Forms.Application]::DoEvents()
        }

        $txtStatus.Text = "Scan complete. $($members.Count) computers processed."
        $txtStatus.Foreground = "#2ECC71"
        $btnScan.IsEnabled = $true
        $progressBar.Visibility = "Collapsed"

        Write-Host "`nScan complete!" -ForegroundColor Green
    })

    # Export button click
    $btnExport.Add_Click({
        if ($script:computerStatuses.Count -eq 0) {
            [System.Windows.MessageBox]::Show("No data to export. Please scan a collection first.", "No Data", "OK", "Information")
            return
        }

        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "SCCMLogReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                $script:computerStatuses | Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Force
                [System.Windows.MessageBox]::Show("Report exported successfully to:`n$($saveDialog.FileName)", "Export Successful", "OK", "Information")
            }
            catch {
                [System.Windows.MessageBox]::Show("Failed to export report:`n$_", "Export Failed", "OK", "Error")
            }
        }
    })

    # Show window
    $window.ShowDialog() | Out-Null
}

#endregion

#region Main

# Save current location to restore later
$originalLocation = Get-Location

try {
    Show-MainWindow
}
finally {
    # Restore original location
    Set-Location $originalLocation
}

#endregion
