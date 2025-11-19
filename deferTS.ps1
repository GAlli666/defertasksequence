<#
.SYNOPSIS
    Task Sequence Deferral Tool with WPF UI for SCCM/ConfigMgr Application Deployment

.DESCRIPTION
    Provides users with deferral options for a Task Sequence previously deployed as Available.
    Designed to be deployed via Application Model with recurring schedule and mandatory installation.

    Features:
    - Tracks deferral count in registry
    - Modern WPF UI with company branding
    - Configurable via XML file
    - Launches Task Sequence when user accepts or deferral limit reached
    - Exits with error code 1 when deferred to trigger retry

.NOTES
    Author: Claude
    Date: 2025-11-19
    Requires: PowerShell 5.1, SCCM Client
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = ""
)

#region Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"

    # Write to console
    switch ($Level) {
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        default   { Write-Host $logMessage -ForegroundColor Gray }
    }

    # Write to log file if path specified in config
    if ($script:Config.Configuration.Settings.LogPath) {
        try {
            Add-Content -Path $script:Config.Configuration.Settings.LogPath -Value $logMessage -ErrorAction SilentlyContinue
        } catch {
            # Silently continue if logging fails
        }
    }
}

function Load-Configuration {
    param([string]$Path)

    try {
        if (-not (Test-Path $Path)) {
            throw "Configuration file not found: $Path"
        }

        # Create XmlDocument object and load the file properly
        $configXml = New-Object System.Xml.XmlDocument
        $configXml.Load($Path)

        # Use Write-Host instead of Write-Log to avoid circular dependency
        # (Write-Log tries to access $script:Config which isn't set yet)
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [Info] Configuration loaded successfully from: $Path" -ForegroundColor Gray

        return $configXml
    }
    catch {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [Error] Failed to load configuration: $_" -ForegroundColor Red
        throw
    }
}

function Get-DeferralCount {
    param(
        [string]$RegistryPath,
        [string]$ValueName
    )

    try {
        if (Test-Path $RegistryPath) {
            $count = Get-ItemProperty -Path $RegistryPath -Name $ValueName -ErrorAction SilentlyContinue
            if ($count) {
                return [int]$count.$ValueName
            }
        }
        return 0
    }
    catch {
        Write-Log "Error reading deferral count: $_" -Level Warning
        return 0
    }
}

function Set-DeferralCount {
    param(
        [string]$RegistryPath,
        [string]$ValueName,
        [int]$Count,
        [hashtable]$Metadata = $null
    )

    try {
        # Ensure registry path exists
        if (-not (Test-Path $RegistryPath)) {
            New-Item -Path $RegistryPath -Force | Out-Null
            Write-Log "Created registry path: $RegistryPath"
        }

        # Set deferral count
        Set-ItemProperty -Path $RegistryPath -Name $ValueName -Value $Count -Type DWord -Force
        Write-Log "Set deferral count to: $Count"

        # If metadata provided and FirstRunDate doesn't exist, set all metadata
        if ($Metadata) {
            $existingFirstRun = Get-ItemProperty -Path $RegistryPath -Name "FirstRunDate" -ErrorAction SilentlyContinue

            if (-not $existingFirstRun) {
                # First time - set all metadata
                Set-ItemProperty -Path $RegistryPath -Name "Vendor" -Value $Metadata.Vendor -Type String -Force
                Set-ItemProperty -Path $RegistryPath -Name "Product" -Value $Metadata.Product -Type String -Force
                Set-ItemProperty -Path $RegistryPath -Name "Version" -Value $Metadata.Version -Type String -Force
                Set-ItemProperty -Path $RegistryPath -Name "FirstRunDate" -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss") -Type String -Force
                Write-Log "Set application metadata: $($Metadata.Vendor) - $($Metadata.Product) v$($Metadata.Version)"
            }
            else {
                # Subsequent runs - update version if changed
                $existingVersion = Get-ItemProperty -Path $RegistryPath -Name "Version" -ErrorAction SilentlyContinue
                if ($existingVersion -and $existingVersion.Version -ne $Metadata.Version) {
                    Set-ItemProperty -Path $RegistryPath -Name "Version" -Value $Metadata.Version -Type String -Force
                    Write-Log "Updated version to: $($Metadata.Version)"
                }
            }
        }

        return $true
    }
    catch {
        Write-Log "Error setting deferral count: $_" -Level Error
        return $false
    }
}

function Start-TaskSequence {
    param([string]$PackageID)

    Write-Log "Attempting to start Task Sequence: $PackageID"

    try {
        # Get SCCM Client COM Object
        $sccmClient = New-Object -ComObject UIResource.UIResourceMgr

        # Get available deployments
        $deployments = $sccmClient.GetAvailableApplications()

        # Find the Task Sequence by Package ID
        $ts = $deployments | Where-Object { $_.PackageID -eq $PackageID }

        if ($ts) {
            Write-Log "Found Task Sequence: $($ts.PackageName)"

            # Execute the Task Sequence
            $sccmClient.ExecuteProgram($ts.ProgramID, $ts.PackageID, $true)

            Write-Log "Task Sequence started successfully"
            return $true
        }
        else {
            Write-Log "Task Sequence not found with Package ID: $PackageID" -Level Error

            # Alternative method: Use WMI to trigger TS
            try {
                Write-Log "Attempting alternative method using scheduled message..."

                $namespace = "ROOT\ccm\clientsdk"
                $class = [wmiclass]"\\localhost\${namespace}:CCM_ProgramsManager"

                # Try to find and execute the TS
                $programs = Get-WmiObject -Namespace "ROOT\ccm\ClientSDK" -Class CCM_Program -Filter "PackageID='$PackageID'"

                if ($programs) {
                    foreach ($program in $programs) {
                        $result = Invoke-WmiMethod -Namespace "ROOT\ccm\ClientSDK" -Class CCM_ProgramsManager -Name ExecuteProgram `
                            -ArgumentList @($program.ProgramID, $PackageID)

                        if ($result.ReturnValue -eq 0) {
                            Write-Log "Task Sequence started via WMI method"
                            return $true
                        }
                    }
                }
            }
            catch {
                Write-Log "Alternative method failed: $_" -Level Error
            }

            return $false
        }
    }
    catch {
        Write-Log "Error starting Task Sequence: $_" -Level Error
        return $false
    }
}

function Show-DeferralUI {
    param(
        [object]$Config,
        [int]$CurrentDeferrals,
        [int]$MaxDeferrals,
        [string]$ScriptDirectory
    )

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    $script:UserChoice = $null

    # Convert RGB strings to Color objects
    function ConvertTo-Color {
        param([string]$rgb)
        $parts = $rgb -split ','
        return [System.Windows.Media.Color]::FromRgb([byte]$parts[0].Trim(), [byte]$parts[1].Trim(), [byte]$parts[2].Trim())
    }

    $accentColor = ConvertTo-Color $Config.Configuration.Settings.UI.AccentColor
    $fontColor = ConvertTo-Color $Config.Configuration.Settings.UI.FontColor
    $backgroundColor = ConvertTo-Color $Config.Configuration.Settings.UI.BackgroundColor

    $accentBrush = New-Object System.Windows.Media.SolidColorBrush($accentColor)
    $fontBrush = New-Object System.Windows.Media.SolidColorBrush($fontColor)
    $backgroundBrush = New-Object System.Windows.Media.SolidColorBrush($backgroundColor)

    # Create XAML
    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="$($Config.Configuration.Settings.UI.WindowTitle)"
    Height="700" Width="900"
    WindowStartupLocation="CenterScreen"
    ResizeMode="NoResize"
    WindowStyle="None"
    AllowsTransparency="True"
    Background="Transparent">

    <Window.Resources>
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#FF1E3A5F"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="16"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="30,12"/>
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
                    <Setter Property="Background" Value="#FF2A5080"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                    <Setter Property="Background" Value="#FF152840"/>
                </Trigger>
            </Style.Triggers>
        </Style>
    </Window.Resources>

    <Border Background="{DynamicResource BackgroundBrush}" CornerRadius="10">
        <Border.Effect>
            <DropShadowEffect BlurRadius="20" ShadowDepth="0" Opacity="0.5"/>
        </Border.Effect>

        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="125"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>

            <!-- Banner -->
            <Border Grid.Row="0" Background="{DynamicResource AccentBrush}" CornerRadius="10,10,0,0">
                <Grid>
                    <Image Name="BannerImage" Stretch="UniformToFill" VerticalAlignment="Center" HorizontalAlignment="Center"/>
                    <TextBlock Name="BannerText"
                               Text="$($Config.Configuration.Settings.UI.BannerText)"
                               FontSize="32"
                               FontWeight="Bold"
                               Foreground="White"
                               HorizontalAlignment="Center"
                               VerticalAlignment="Center"
                               TextWrapping="Wrap"
                               Margin="20"/>
                </Grid>
            </Border>

            <!-- Main Content -->
            <Grid Grid.Row="1" Margin="40,30,40,40">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Main Message -->
                <StackPanel Grid.Row="0" Name="MainPanel">
                    <TextBlock Name="MainMessage"
                               Text="$($Config.Configuration.Settings.UI.MainMessage)"
                               FontSize="18"
                               Foreground="{DynamicResource FontBrush}"
                               TextWrapping="Wrap"
                               LineHeight="28"
                               Margin="0,0,0,20"/>

                    <TextBlock Name="DeferralInfo"
                               FontSize="14"
                               Foreground="{DynamicResource FontBrush}"
                               Opacity="0.7"
                               Margin="0,0,0,30"/>

                    <!-- Timeout Warning -->
                    <StackPanel Name="TimeoutWarning" HorizontalAlignment="Center" Margin="0,20,0,0">
                        <TextBlock Name="TimeoutWarningText"
                                   FontSize="22"
                                   FontWeight="Bold"
                                   Foreground="Red"
                                   HorizontalAlignment="Center"
                                   TextAlignment="Center"
                                   TextWrapping="Wrap"/>
                        <TextBlock Name="TimeoutDateText"
                                   FontSize="18"
                                   FontWeight="SemiBold"
                                   Foreground="Red"
                                   HorizontalAlignment="Center"
                                   TextAlignment="Center"
                                   Margin="0,5,0,0"/>
                        <TextBlock Name="TimeoutNoInputText"
                                   FontSize="16"
                                   Foreground="Red"
                                   HorizontalAlignment="Center"
                                   TextAlignment="Center"
                                   Margin="0,5,0,0"/>
                        <TextBlock Name="TimeoutCountdown"
                                   FontSize="20"
                                   FontWeight="Bold"
                                   Foreground="Red"
                                   HorizontalAlignment="Center"
                                   TextAlignment="Center"
                                   Margin="0,10,0,0"/>
                    </StackPanel>
                </StackPanel>

                <!-- Secondary Panel (Hidden by default) -->
                <StackPanel Grid.Row="1" Name="SecondaryPanel" Visibility="Collapsed">
                    <TextBlock Name="SecondaryMessage"
                               Text="$($Config.Configuration.Settings.UI.SecondaryMessage)"
                               FontSize="18"
                               Foreground="{DynamicResource FontBrush}"
                               TextWrapping="Wrap"
                               LineHeight="28"
                               HorizontalAlignment="Center"
                               TextAlignment="Center"
                               Margin="0,40,0,0"/>
                </StackPanel>

                <!-- Final Panel (Hidden by default) -->
                <StackPanel Grid.Row="1" Name="FinalPanel" Visibility="Collapsed" VerticalAlignment="Center">
                    <TextBlock Name="FinalMessage"
                               Text="$($Config.Configuration.Settings.UI.FinalMessage)"
                               FontSize="20"
                               FontWeight="SemiBold"
                               Foreground="{DynamicResource FontBrush}"
                               TextWrapping="Wrap"
                               LineHeight="32"
                               HorizontalAlignment="Center"
                               TextAlignment="Center"/>

                    <TextBlock Name="CountdownText"
                               FontSize="48"
                               FontWeight="Bold"
                               Foreground="{DynamicResource AccentBrush}"
                               HorizontalAlignment="Center"
                               Margin="0,20,0,0"/>
                </StackPanel>

                <!-- Buttons -->
                <StackPanel Grid.Row="2" Name="ButtonPanel" Orientation="Horizontal" HorizontalAlignment="Center">
                    <Button Name="DeferButton"
                            Content="Defer"
                            Style="{StaticResource ModernButton}"
                            Width="180"
                            Margin="0,0,20,0"/>
                    <Button Name="InstallButton"
                            Content="Install Now"
                            Style="{StaticResource ModernButton}"
                            Width="180"/>
                </StackPanel>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

    # Load XAML
    $reader = New-Object System.Xml.XmlNodeReader($xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Apply dynamic brushes
    $window.Resources.Add("AccentBrush", $accentBrush)
    $window.Resources.Add("FontBrush", $fontBrush)
    $window.Resources.Add("BackgroundBrush", $backgroundBrush)

    # Get controls
    $bannerImage = $window.FindName("BannerImage")
    $bannerText = $window.FindName("BannerText")
    $deferralInfo = $window.FindName("DeferralInfo")
    $mainPanel = $window.FindName("MainPanel")
    $secondaryPanel = $window.FindName("SecondaryPanel")
    $finalPanel = $window.FindName("FinalPanel")
    $buttonPanel = $window.FindName("ButtonPanel")
    $deferButton = $window.FindName("DeferButton")
    $installButton = $window.FindName("InstallButton")
    $countdownText = $window.FindName("CountdownText")
    $timeoutWarning = $window.FindName("TimeoutWarning")
    $timeoutWarningText = $window.FindName("TimeoutWarningText")
    $timeoutDateText = $window.FindName("TimeoutDateText")
    $timeoutNoInputText = $window.FindName("TimeoutNoInputText")
    $timeoutCountdown = $window.FindName("TimeoutCountdown")

    # Load banner image if exists
    $bannerPath = Join-Path $ScriptDirectory $Config.Configuration.Settings.UI.BannerImagePath
    if (Test-Path $bannerPath) {
        try {
            $bannerImage.Source = New-Object System.Windows.Media.Imaging.BitmapImage(New-Object Uri($bannerPath))
            $bannerText.Visibility = [System.Windows.Visibility]::Collapsed
        }
        catch {
            Write-Log "Could not load banner image: $_" -Level Warning
        }
    }

    # Set deferral info
    $deferralsRemaining = $MaxDeferrals - $CurrentDeferrals
    if ($deferralsRemaining -gt 0) {
        if ($deferralsRemaining -eq 1) {
            $deferralInfo.Text = "You can defer this installation 1 more time."
        }
        else {
            $deferralInfo.Text = "You can defer this installation $deferralsRemaining more times."
        }
    }
    else {
        $deferralInfo.Text = "This is a required installation and cannot be deferred."
        $deferButton.IsEnabled = $false
        $deferButton.Opacity = 0.5
    }

    # Defer button click
    $deferButton.Add_Click({
        if ($deferralsRemaining -gt 0) {
            $script:UserChoice = 'Defer'
            $window.Close()
        }
    })

    # Install button click - First click
    $installButton.Add_Click({
        if ($secondaryPanel.Visibility -eq [System.Windows.Visibility]::Collapsed) {
            # Show secondary confirmation
            $mainPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $secondaryPanel.Visibility = [System.Windows.Visibility]::Visible
            $installButton.Content = "Yes, Install Now"
            $deferButton.Content = "No, Go Back"

            # Update defer button to go back
            $deferButton.Tag = "GoBack"
        }
        else {
            # User confirmed installation
            $script:UserChoice = 'Install'
            $window.Close()
        }
    })

    # Handle defer button in secondary state
    $deferButton.Add_Click({
        if ($deferButton.Tag -eq "GoBack") {
            # Go back to main screen
            $mainPanel.Visibility = [System.Windows.Visibility]::Visible
            $secondaryPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $installButton.Content = "Install Now"
            $deferButton.Content = "Defer"
            $deferButton.Tag = $null
        }
    })

    # Prevent window from being closed via Alt+F4, X button, or other means (except programmatic close)
    $window.Add_Closing({
        param($sender, $e)
        # Only allow closing if UserChoice has been set (programmatic close)
        if ($null -eq $script:UserChoice) {
            $e.Cancel = $true
            Write-Log "User attempted to close window via system method - prevented" -Level Warning
        }
    })

    # Set up main window timeout (if deferrals still available)
    if ($deferralsRemaining -gt 0) {
        # Get timeout in minutes from config
        $timeoutMinutes = [int]$Config.Configuration.Settings.MainWindowTimeoutMinutes

        # Set timeout warning text
        $timeoutWarningText.Text = $Config.Configuration.Settings.UI.TimeoutWarningText
        $timeoutNoInputText.Text = $Config.Configuration.Settings.UI.TimeoutNoInputText

        # Calculate deadline date/time (without seconds)
        $deadlineTime = (Get-Date).AddMinutes($timeoutMinutes)
        $timeoutDateText.Text = $deadlineTime.ToString("yyyy-MM-dd HH:mm")

        # Start timeout timer
        $script:timeoutSecondsRemaining = $timeoutMinutes * 60
        $mainTimeoutTimer = New-Object System.Windows.Threading.DispatcherTimer
        $mainTimeoutTimer.Interval = [TimeSpan]::FromSeconds(1)

        $mainTimeoutTimer.Add_Tick({
            $script:timeoutSecondsRemaining--

            # Update countdown display
            $minutes = [Math]::Floor($script:timeoutSecondsRemaining / 60)
            $seconds = $script:timeoutSecondsRemaining % 60
            $timeoutCountdown.Text = "Time remaining: $minutes minutes, $seconds seconds"

            if ($script:timeoutSecondsRemaining -le 0) {
                $mainTimeoutTimer.Stop()
                Write-Log "Main window timeout reached - auto-starting installation" -Level Warning
                $script:UserChoice = 'Install'
                $window.Close()
            }
        })

        # Initial display
        $minutes = [Math]::Floor($script:timeoutSecondsRemaining / 60)
        $seconds = $script:timeoutSecondsRemaining % 60
        $timeoutCountdown.Text = "Time remaining: $minutes minutes, $seconds seconds"

        # Start timer when window loads
        $window.Add_Loaded({
            $mainTimeoutTimer.Start()
            Write-Log "Main window timeout started: $timeoutMinutes minutes"
        })
    }

    # If deferral limit reached, skip main dialog and show countdown immediately
    if ($deferralsRemaining -le 0) {
        # Hide main content, buttons, and timeout warning immediately
        $mainPanel.Visibility = [System.Windows.Visibility]::Collapsed
        $buttonPanel.Visibility = [System.Windows.Visibility]::Collapsed
        $finalPanel.Visibility = [System.Windows.Visibility]::Visible

        # Hide timeout warning panel (only relevant when deferrals available)
        if ($null -ne $timeoutWarning) {
            $timeoutWarning.Visibility = [System.Windows.Visibility]::Collapsed
        }

        $window.Add_Loaded({
            # Verify countdown control exists
            if ($null -eq $countdownText) {
                Write-Log "CountdownText control not found - cannot display countdown" -Level Error
                $script:UserChoice = 'Install'
                $window.Close()
                return
            }

            # Start countdown
            $script:secondsRemaining = [int]$Config.Configuration.Settings.UI.FinalMessageDuration
            $countdownTimer = New-Object System.Windows.Threading.DispatcherTimer
            $countdownTimer.Interval = [TimeSpan]::FromSeconds(1)

            $countdownTimer.Add_Tick({
                $script:secondsRemaining--
                $countdownText.Text = $script:secondsRemaining

                if ($script:secondsRemaining -le 0) {
                    $countdownTimer.Stop()
                    $script:UserChoice = 'Install'
                    $window.Close()
                }
            })

            $countdownText.Text = $script:secondsRemaining
            $countdownTimer.Start()
        })
    }

    # Show window
    $window.ShowDialog() | Out-Null

    return $script:UserChoice
}

#endregion

#region Main Script

try {
    # Resolve script directory for config file and banner image
    # Using Split-Path instead of $PSScriptRoot for compatibility with SCCM and other contexts
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrEmpty($scriptPath)) {
        # Fallback to current directory if script path cannot be determined
        $scriptDirectory = Get-Location | Select-Object -ExpandProperty Path
    }
    else {
        $scriptDirectory = Split-Path -Parent $scriptPath
    }

    # Resolve config file path if not provided
    if ([string]::IsNullOrEmpty($ConfigFile)) {
        $ConfigFile = Join-Path $scriptDirectory "DeferTSConfig.xml"
    }
    elseif (-not [System.IO.Path]::IsPathRooted($ConfigFile)) {
        # If relative path provided, make it absolute relative to script directory
        $ConfigFile = Join-Path $scriptDirectory $ConfigFile
    }

    # Load configuration
    $script:Config = Load-Configuration -Path $ConfigFile

    Write-Log "=== Task Sequence Deferral Tool Started ==="

    # Get configuration values
    $maxDeferrals = [int]$Config.Configuration.Settings.MaxDeferrals
    $registryBasePath = $Config.Configuration.Settings.RegistryPath
    $registryValue = $Config.Configuration.Settings.RegistryValueName
    $packageID = $Config.Configuration.Settings.TaskSequence.PackageID

    # Build registry path with Package ID as subfolder
    # This allows reuse of the solution for multiple Task Sequences
    $registryPath = Join-Path $registryBasePath $packageID

    # Create metadata hashtable from config
    $metadata = @{
        Vendor  = $Config.Configuration.Settings.Application.Vendor
        Product = $Config.Configuration.Settings.Application.Product
        Version = $Config.Configuration.Settings.Application.Version
    }

    Write-Log "Max Deferrals: $maxDeferrals"
    Write-Log "Registry Path: $registryPath"
    Write-Log "Task Sequence Package ID: $packageID"
    Write-Log "Application: $($metadata.Vendor) - $($metadata.Product) v$($metadata.Version)"

    # Get current deferral count
    $currentDeferrals = Get-DeferralCount -RegistryPath $registryPath -ValueName $registryValue
    Write-Log "Current Deferral Count: $currentDeferrals"

    # Determine if we've reached the deferral limit
    $deferralsRemaining = $maxDeferrals - $currentDeferrals

    # INCREMENT DEFERRAL COUNT IMMEDIATELY (before showing UI)
    # This ensures that force-closing the window counts as a deferral
    if ($deferralsRemaining -gt 0) {
        $newCount = $currentDeferrals + 1
        if (Set-DeferralCount -RegistryPath $registryPath -ValueName $registryValue -Count $newCount -Metadata $metadata) {
            Write-Log "Deferral count incremented immediately: $newCount / $maxDeferrals"
            Write-Log "This ensures force-close counts as a deferral"
            # Update current count to reflect the increment
            $currentDeferrals = $newCount
        }
        else {
            Write-Log "Failed to increment deferral count. Proceeding with installation." -Level Warning
            # If we can't track deferrals, force installation
            $currentDeferrals = $maxDeferrals
        }
    }

    # Show UI with the already-incremented count
    Write-Log "Displaying deferral UI..."
    $userChoice = Show-DeferralUI -Config $Config -CurrentDeferrals $currentDeferrals -MaxDeferrals $maxDeferrals -ScriptDirectory $scriptDirectory

    Write-Log "User choice: $userChoice"

    # Process user choice
    if ($userChoice -eq 'Defer') {
        # Deferral count was already incremented above
        # Just exit with error code 1 to trigger retry by SCCM
        Write-Log "User chose to defer. Count already incremented to: $currentDeferrals / $maxDeferrals"
        Write-Log "Exiting with code 1 to trigger retry"
        exit 1
    }
    elseif ($userChoice -eq 'Install') {
        # User chose to install - start Task Sequence
        Write-Log "User chose to install. Starting Task Sequence..."

        if (Start-TaskSequence -PackageID $packageID) {
            Write-Log "Task Sequence started successfully"

            # Reset deferral count after successful start
            Set-DeferralCount -RegistryPath $registryPath -ValueName $registryValue -Count 0 -Metadata $metadata
            Write-Log "Deferral count reset to 0"

            Write-Log "=== Task Sequence Deferral Tool Completed Successfully ==="
            exit 0
        }
        else {
            Write-Log "Failed to start Task Sequence" -Level Error
            Write-Log "=== Task Sequence Deferral Tool Completed with Errors ==="
            exit 1
        }
    }
    else {
        # User closed window without choosing (or null choice)
        # Deferral count was already incremented, so this counts as a deferral
        Write-Log "User closed window without making a choice. Count already incremented." -Level Warning
        Write-Log "Exiting with code 1 to trigger retry"
        exit 1
    }
}
catch {
    Write-Log "Critical error: $_" -Level Error
    Write-Log $_.ScriptStackTrace -Level Error
    exit 1
}

#endregion
