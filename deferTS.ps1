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
    Requires: PowerShell 3.0+, SCCM Client
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = "$PSScriptRoot\DeferTSConfig.xml"
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
    if ($script:Config.Settings.LogPath) {
        try {
            Add-Content -Path $script:Config.Settings.LogPath -Value $logMessage -ErrorAction SilentlyContinue
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

        [xml]$configXml = Get-Content -Path $Path -ErrorAction Stop
        Write-Log "Configuration loaded successfully from: $Path"
        return $configXml
    }
    catch {
        Write-Log "Failed to load configuration: $_" -Level Error
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
        [int]$Count
    )

    try {
        # Ensure registry path exists
        if (-not (Test-Path $RegistryPath)) {
            New-Item -Path $RegistryPath -Force | Out-Null
            Write-Log "Created registry path: $RegistryPath"
        }

        Set-ItemProperty -Path $RegistryPath -Name $ValueName -Value $Count -Type DWord -Force
        Write-Log "Set deferral count to: $Count"
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
        [int]$MaxDeferrals
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

    $accentColor = ConvertTo-Color $Config.Settings.UI.AccentColor
    $fontColor = ConvertTo-Color $Config.Settings.UI.FontColor
    $backgroundColor = ConvertTo-Color $Config.Settings.UI.BackgroundColor

    $accentBrush = New-Object System.Windows.Media.SolidColorBrush($accentColor)
    $fontBrush = New-Object System.Windows.Media.SolidColorBrush($fontColor)
    $backgroundBrush = New-Object System.Windows.Media.SolidColorBrush($backgroundColor)

    # Create XAML
    [xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="$($Config.Settings.UI.WindowTitle)"
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
                               Text="$($Config.Settings.UI.BannerText)"
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
                               Text="$($Config.Settings.UI.MainMessage)"
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
                </StackPanel>

                <!-- Secondary Panel (Hidden by default) -->
                <StackPanel Grid.Row="1" Name="SecondaryPanel" Visibility="Collapsed">
                    <TextBlock Name="SecondaryMessage"
                               Text="$($Config.Settings.UI.SecondaryMessage)"
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
                               Text="$($Config.Settings.UI.FinalMessage)"
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

    # Load banner image if exists
    $bannerPath = Join-Path $PSScriptRoot $Config.Settings.UI.BannerImagePath
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
        $deferralInfo.Text = "You have $deferralsRemaining deferral(s) remaining."
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

    # If deferral limit reached, auto-start countdown
    if ($deferralsRemaining -le 0) {
        $window.Add_Loaded({
            # Hide main content and buttons
            $mainPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $buttonPanel.Visibility = [System.Windows.Visibility]::Collapsed
            $finalPanel.Visibility = [System.Windows.Visibility]::Visible

            # Start countdown
            $seconds = [int]$Config.Settings.UI.FinalMessageDuration
            $timer = New-Object System.Windows.Threading.DispatcherTimer
            $timer.Interval = [TimeSpan]::FromSeconds(1)

            $timer.Add_Tick({
                $script:seconds--
                $countdownText.Text = $script:seconds

                if ($script:seconds -le 0) {
                    $timer.Stop()
                    $script:UserChoice = 'Install'
                    $window.Close()
                }
            })

            $countdownText.Text = $seconds
            $timer.Start()
        })
    }

    # Show window
    $window.ShowDialog() | Out-Null

    return $script:UserChoice
}

#endregion

#region Main Script

try {
    # Load configuration
    $script:Config = Load-Configuration -Path $ConfigFile

    Write-Log "=== Task Sequence Deferral Tool Started ==="

    # Get configuration values
    $maxDeferrals = [int]$Config.Settings.MaxDeferrals
    $registryPath = $Config.Settings.RegistryPath
    $registryValue = $Config.Settings.RegistryValueName
    $packageID = $Config.Settings.TaskSequence.PackageID

    Write-Log "Max Deferrals: $maxDeferrals"
    Write-Log "Registry Path: $registryPath"
    Write-Log "Task Sequence Package ID: $packageID"

    # Get current deferral count
    $currentDeferrals = Get-DeferralCount -RegistryPath $registryPath -ValueName $registryValue
    Write-Log "Current Deferral Count: $currentDeferrals"

    # Show UI
    Write-Log "Displaying deferral UI..."
    $userChoice = Show-DeferralUI -Config $Config -CurrentDeferrals $currentDeferrals -MaxDeferrals $maxDeferrals

    Write-Log "User choice: $userChoice"

    # Process user choice
    if ($userChoice -eq 'Defer' -and $currentDeferrals -lt $maxDeferrals) {
        # Increment deferral count
        $newCount = $currentDeferrals + 1

        if (Set-DeferralCount -RegistryPath $registryPath -ValueName $registryValue -Count $newCount) {
            Write-Log "Installation deferred. Count: $newCount / $maxDeferrals"

            # Exit with error code 1 to trigger retry by SCCM
            Write-Log "Exiting with code 1 to trigger retry"
            exit 1
        }
        else {
            Write-Log "Failed to update deferral count. Proceeding with installation." -Level Warning

            # Start TS since we couldn't save the deferral
            if (Start-TaskSequence -PackageID $packageID) {
                Write-Log "Task Sequence started successfully"
                exit 0
            }
            else {
                Write-Log "Failed to start Task Sequence" -Level Error
                exit 1
            }
        }
    }
    elseif ($userChoice -eq 'Install' -or $currentDeferrals -ge $maxDeferrals) {
        # Start Task Sequence
        Write-Log "Starting Task Sequence installation..."

        if (Start-TaskSequence -PackageID $packageID) {
            Write-Log "Task Sequence started successfully"

            # Reset deferral count after successful start
            Set-DeferralCount -RegistryPath $registryPath -ValueName $registryValue -Count 0

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
        # User closed window without choosing
        Write-Log "User closed window without making a choice. Exiting..." -Level Warning
        exit 1
    }
}
catch {
    Write-Log "Critical error: $_" -Level Error
    Write-Log $_.ScriptStackTrace -Level Error
    exit 1
}

#endregion
