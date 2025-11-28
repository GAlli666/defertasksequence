<#
.SYNOPSIS
    SCCM Task Sequence Deferral Monitor - Backend Data Collection Script

.DESCRIPTION
    Connects to SCCM, collects collection member data, retrieves deferral logs,
    and generates JSON data files for the web frontend to consume.

.NOTES
    Date: 2025-11-28
    Requires: PowerShell 5.1, ConfigurationManager PowerShell Module, SCCM Admin Rights
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = ""
)

#region Configuration

function Load-Configuration {
    param([string]$Path)

    try {
        if (-not (Test-Path $Path)) {
            throw "Configuration file not found: $Path"
        }

        $configXml = New-Object System.Xml.XmlDocument
        $configXml.Load($Path)

        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Configuration loaded from: $Path" -ForegroundColor Green

        return $configXml
    }
    catch {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Failed to load configuration: $_" -ForegroundColor Red
        throw
    }
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

        if ([string]::IsNullOrEmpty($SiteCode) -or [string]::IsNullOrEmpty($SiteServer)) {
            throw "Site Code or Site Server is null or empty"
        }

        # Import ConfigurationManager module
        if (-not (Get-Module -Name ConfigurationManager)) {
            if (-not $ENV:SMS_ADMIN_UI_PATH) {
                throw "SMS_ADMIN_UI_PATH environment variable not found. Please ensure SCCM Console is installed."
            }

            $modulePath = "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1"
            Import-Module $modulePath -ErrorAction Stop
            Write-Host "ConfigurationManager module imported" -ForegroundColor Green
        }

        # Check if PSDrive already exists
        $siteCodePath = "$SiteCode" + ":"
        $existingDrive = Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue

        if (-not $existingDrive) {
            $newDrive = New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -ErrorAction Stop
            Write-Host "PSDrive created: $($newDrive.Name)" -ForegroundColor Green
        }

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

        $currentLocation = Get-Location
        $siteCodePath = "$($script:sccmSiteCode):"

        Set-Location $siteCodePath -ErrorAction Stop

        $members = Get-CMCollectionMember -CollectionId $CollectionID -ErrorAction Stop

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
        if ($currentLocation) {
            Set-Location $currentLocation -ErrorAction SilentlyContinue
        }
        return $null
    }
}

function Get-TaskSequenceDeploymentStatus {
    param(
        [string]$ComputerName,
        [string]$TaskSequenceID
    )

    try {
        $currentLocation = Get-Location
        $siteCodePath = "$($script:sccmSiteCode):"
        Set-Location $siteCodePath -ErrorAction Stop

        # Query deployment status for the specific computer and task sequence
        $query = "SELECT * FROM SMS_DeploymentSummary WHERE PackageID='$TaskSequenceID'"
        $deployment = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" -ComputerName $script:sccmSiteServer -Query $query -ErrorAction SilentlyContinue

        Set-Location $currentLocation

        if ($deployment) {
            # Try to get specific status for this computer
            $statusQuery = "SELECT * FROM SMS_ClientAdvertisementStatus WHERE AdvertisementID='$($deployment.DeploymentID)' AND ResourceID IN (SELECT ResourceID FROM SMS_R_System WHERE Name='$ComputerName')"
            $status = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" -ComputerName $script:sccmSiteServer -Query $statusQuery -ErrorAction SilentlyContinue

            if ($status) {
                switch ($status.LastState) {
                    0 { return "Not Started" }
                    1 { return "Success" }
                    2 { return "In Progress" }
                    3 { return "Requirements Not Met" }
                    4 { return "Failed" }
                    default { return "Unknown" }
                }
            }
        }

        return "Not Started"
    }
    catch {
        Write-Host "Error getting TS status for ${ComputerName}: $_" -ForegroundColor Yellow
        return "Unknown"
    }
    finally {
        if ($currentLocation) {
            Set-Location $currentLocation -ErrorAction SilentlyContinue
        }
    }
}

function Get-DevicePrimaryUser {
    param(
        [string]$ComputerName
    )

    try {
        $currentLocation = Get-Location
        $siteCodePath = "$($script:sccmSiteCode):"
        Set-Location $siteCodePath -ErrorAction Stop

        $device = Get-CMDevice -Name $ComputerName -ErrorAction SilentlyContinue

        Set-Location $currentLocation

        if ($device) {
            $primaryUser = $device.PrimaryUser
            if ($primaryUser) {
                # Extract username from domain\user format
                if ($primaryUser -match '\\(.+)$') {
                    return $matches[1]
                }
                return $primaryUser
            }
        }

        return "N/A"
    }
    catch {
        if ($currentLocation) {
            Set-Location $currentLocation -ErrorAction SilentlyContinue
        }
        return "N/A"
    }
}

function Get-DeviceOSVersion {
    param(
        [string]$ComputerName
    )

    try {
        $currentLocation = Get-Location
        $siteCodePath = "$($script:sccmSiteCode):"
        Set-Location $siteCodePath -ErrorAction Stop

        $device = Get-CMDevice -Name $ComputerName -ErrorAction SilentlyContinue

        Set-Location $currentLocation

        if ($device) {
            $os = $device.OSVersion
            if ($os -and $os -match '10\.0\.22') {
                return "Windows 11"
            }
        }

        return "Unknown"
    }
    catch {
        if ($currentLocation) {
            Set-Location $currentLocation -ErrorAction SilentlyContinue
        }
        return "Unknown"
    }
}

#endregion

#region Log File Functions

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

function Confirm-ComputerIdentity {
    <#
    .SYNOPSIS
        Verifies that the computer we're connecting to is the one we expect
    .DESCRIPTION
        Uses WMI to query the actual hostname from the remote computer and compares it
        to the expected name. Critical for VPN scenarios where DNS may be outdated.
    .OUTPUTS
        Hashtable with IsValid (bool) and ActualName (string)
    #>
    param(
        [string]$ExpectedName,
        [int]$TimeoutSeconds = 3
    )

    $result = @{
        IsValid = $false
        ActualName = "Unknown"
        ErrorMessage = ""
    }

    try {
        # Set WMI query timeout
        $connectionOptions = New-Object System.Management.ConnectionOptions
        $connectionOptions.Timeout = New-TimeSpan -Seconds $TimeoutSeconds

        $scope = New-Object System.Management.ManagementScope("\\$ExpectedName\root\cimv2", $connectionOptions)
        $scope.Connect()

        $query = New-Object System.Management.ObjectQuery("SELECT Name FROM Win32_ComputerSystem")
        $searcher = New-Object System.Management.ManagementObjectSearcher($scope, $query)

        # Execute with timeout
        $computers = $searcher.Get()

        foreach ($computer in $computers) {
            $result.ActualName = $computer["Name"]

            # Compare names (case-insensitive)
            if ($result.ActualName -eq $ExpectedName) {
                $result.IsValid = $true
            }
            else {
                $result.ErrorMessage = "Hostname mismatch: Expected '$ExpectedName', got '$($result.ActualName)'"
            }

            break
        }

        if ([string]::IsNullOrEmpty($result.ActualName)) {
            $result.ErrorMessage = "Could not retrieve hostname via WMI"
        }
    }
    catch [System.Management.ManagementException] {
        $result.ErrorMessage = "WMI error: $($_.Exception.Message)"
    }
    catch [System.UnauthorizedAccessException] {
        $result.ErrorMessage = "Access denied"
    }
    catch {
        $result.ErrorMessage = "Verification failed: $($_.Exception.Message)"
    }

    return $result
}

function Get-DeferralLogData {
    param(
        [string]$LogPath,
        [string]$ComputerName
    )

    $result = @{
        DeferralCount = 0
        TSTriggerAttempted = $false
        TSTriggerSuccess = $false
        LastDeferralDate = "N/A"
        LastTriggerDate = "N/A"
        LogAvailable = $false
        ErrorMessage = ""
    }

    try {
        $uncPath = "\\$ComputerName\$($LogPath.Replace(':', '$'))"

        if (-not (Test-Path $uncPath)) {
            $result.ErrorMessage = "Log file not found"
            return $result
        }

        $logContent = Get-Content -Path $uncPath -ErrorAction Stop

        if ($logContent.Count -eq 0) {
            $result.ErrorMessage = "Log file is empty"
            return $result
        }

        $result.LogAvailable = $true

        # Parse log entries
        $deferrals = 0
        $triggerAttempted = $false
        $triggerSuccess = $false
        $lastDeferralDate = $null
        $lastTriggerDate = $null

        foreach ($line in $logContent) {
            if ($line -match '^\[([\d-]+\s+[\d:]+)\]\s+\[(\w+)\]\s+(.+)$') {
                $timestamp = [datetime]::ParseExact($matches[1], "yyyy-MM-dd HH:mm:ss", $null)
                $message = $matches[3]

                if ($message -match 'Deferral count incremented immediately:\s+(\d+)\s+/\s+\d+') {
                    $deferrals = [int]$matches[1]
                    $lastDeferralDate = $timestamp
                }

                if ($message -match 'Deferral count reset to 0') {
                    $deferrals = 0
                }

                if ($message -match 'User chose to defer') {
                    $lastDeferralDate = $timestamp
                }

                if ($message -match 'Attempting to start Task Sequence') {
                    $triggerAttempted = $true
                    $lastTriggerDate = $timestamp
                }

                if ($message -match 'Task Sequence (triggered|started) successfully') {
                    $triggerSuccess = $true
                }

                if ($message -match 'Failed to (start Task Sequence|trigger schedule)') {
                    $triggerSuccess = $false
                }
            }
        }

        $result.DeferralCount = $deferrals
        $result.TSTriggerAttempted = $triggerAttempted
        $result.TSTriggerSuccess = $triggerSuccess

        if ($lastDeferralDate) {
            $result.LastDeferralDate = $lastDeferralDate.ToString("yyyy-MM-dd HH:mm:ss")
        }

        if ($lastTriggerDate) {
            $result.LastTriggerDate = $lastTriggerDate.ToString("yyyy-MM-dd HH:mm:ss")
        }

    }
    catch {
        $result.ErrorMessage = $_.Exception.Message
    }

    return $result
}

function Copy-DeferralLog {
    param(
        [string]$ComputerName,
        [string]$SourcePath,
        [string]$DestinationDirectory
    )

    try {
        $uncPath = "\\$ComputerName\$($SourcePath.Replace(':', '$'))"

        if (-not (Test-Path $uncPath)) {
            return $false
        }

        $sourceFile = Get-Item $uncPath -ErrorAction Stop
        $destFile = Join-Path $DestinationDirectory "$ComputerName`_TaskSequenceDeferral.log"

        # Only copy if source is newer or destination doesn't exist
        if (-not (Test-Path $destFile)) {
            Copy-Item $uncPath $destFile -Force -ErrorAction Stop
            Write-Host "Copied deferral log from $ComputerName" -ForegroundColor Green
            return $true
        }

        $destFileInfo = Get-Item $destFile
        if ($sourceFile.LastWriteTime -gt $destFileInfo.LastWriteTime) {
            Copy-Item $uncPath $destFile -Force -ErrorAction Stop
            Write-Host "Updated deferral log from $ComputerName (newer version available)" -ForegroundColor Green
            return $true
        }

        return $true
    }
    catch {
        Write-Host "Failed to copy deferral log from ${ComputerName}: $_" -ForegroundColor Yellow
        return $false
    }
}

function Get-TSLogsFromMachine {
    param(
        [string]$ComputerName,
        [string]$DestinationDirectory,
        [string]$DateMode = "DaysBack",
        [int]$LookbackDays = 7,
        [string]$StartDate = ""
    )

    try {
        $tsLogPath = "\\$ComputerName\C$\Windows\CCM\Logs\SMSTSLog"

        if (-not (Test-Path $tsLogPath)) {
            Write-Host "TS log directory not found on $ComputerName" -ForegroundColor Yellow
            return $false
        }

        # Determine cutoff date based on mode
        $cutoffDate = $null
        if ($DateMode -eq "FromDate" -and -not [string]::IsNullOrEmpty($StartDate)) {
            try {
                $cutoffDate = [datetime]::ParseExact($StartDate, "yyyy-MM-dd", $null)
                Write-Host "  Using start date: $($cutoffDate.ToString('yyyy-MM-dd'))" -ForegroundColor Gray
            }
            catch {
                Write-Host "  Invalid start date format, using DaysBack mode instead" -ForegroundColor Yellow
                $cutoffDate = (Get-Date).AddDays(-$LookbackDays)
            }
        }
        else {
            # DaysBack mode (default)
            $cutoffDate = (Get-Date).AddDays(-$LookbackDays)
            Write-Host "  Looking back $LookbackDays days from today" -ForegroundColor Gray
        }

        # Get log files newer than cutoff date
        $logFiles = Get-ChildItem -Path $tsLogPath -Filter "*.log" -ErrorAction Stop | Where-Object { $_.LastWriteTime -gt $cutoffDate }

        if ($logFiles.Count -eq 0) {
            Write-Host "No TS logs found on $ComputerName since $($cutoffDate.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow
            return $false
        }

        # Create destination directory for this machine's TS logs
        $destDir = Join-Path $DestinationDirectory "$ComputerName`_TSLogs"
        if (-not (Test-Path $destDir)) {
            New-Item -Path $destDir -ItemType Directory -Force | Out-Null
        }

        # Copy log files (overwrite if newer, keeps old files forever)
        $copiedCount = 0
        foreach ($logFile in $logFiles) {
            $destFile = Join-Path $destDir $logFile.Name

            # Only copy if destination doesn't exist or source is newer
            if (-not (Test-Path $destFile)) {
                Copy-Item $logFile.FullName $destFile -Force -ErrorAction SilentlyContinue
                $copiedCount++
            }
            else {
                $destFileInfo = Get-Item $destFile
                if ($logFile.LastWriteTime -gt $destFileInfo.LastWriteTime) {
                    Copy-Item $logFile.FullName $destFile -Force -ErrorAction SilentlyContinue
                    $copiedCount++
                }
            }
        }

        if ($copiedCount -gt 0) {
            Write-Host "Copied/updated $copiedCount TS log files from $ComputerName" -ForegroundColor Green
        }
        else {
            Write-Host "All TS logs already up to date for $ComputerName" -ForegroundColor Gray
        }

        return $true
    }
    catch {
        Write-Host "Failed to copy TS logs from ${ComputerName}: $_" -ForegroundColor Yellow
        return $false
    }
}

function Remove-OrphanedLogs {
    param(
        [string]$LogDirectory,
        [array]$CurrentMembers
    )

    try {
        $memberNames = $CurrentMembers | ForEach-Object { $_.Name }

        # Clean up deferral logs
        $deferralLogs = Get-ChildItem -Path $LogDirectory -Filter "*_TaskSequenceDeferral.log" -ErrorAction SilentlyContinue
        foreach ($log in $deferralLogs) {
            $computerName = $log.Name -replace '_TaskSequenceDeferral\.log$', ''
            if ($computerName -notin $memberNames) {
                Remove-Item $log.FullName -Force -ErrorAction SilentlyContinue
                Write-Host "Removed orphaned deferral log: $($log.Name)" -ForegroundColor Yellow
            }
        }

        # Clean up TS log directories
        $tsLogDirs = Get-ChildItem -Path $LogDirectory -Filter "*_TSLogs" -Directory -ErrorAction SilentlyContinue
        foreach ($dir in $tsLogDirs) {
            $computerName = $dir.Name -replace '_TSLogs$', ''
            if ($computerName -notin $memberNames) {
                Remove-Item $dir.FullName -Recurse -Force -ErrorAction SilentlyContinue
                Write-Host "Removed orphaned TS log directory: $($dir.Name)" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "Error during orphaned log cleanup: $_" -ForegroundColor Red
    }
}

#endregion

#region Data Export Functions

function Export-DataToJSON {
    param(
        [array]$DeviceData,
        [string]$OutputPath
    )

    try {
        $json = $DeviceData | ConvertTo-Json -Depth 10
        $json | Out-File -FilePath $OutputPath -Encoding utf8 -Force

        Write-Host "Data exported to: $OutputPath" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to export data to JSON: $_" -ForegroundColor Red
        return $false
    }
}

#endregion

#region Main Script

try {
    # Resolve script directory
    $scriptPath = $MyInvocation.MyCommand.Path
    if ([string]::IsNullOrEmpty($scriptPath)) {
        $scriptDirectory = Get-Location | Select-Object -ExpandProperty Path
    }
    else {
        $scriptDirectory = Split-Path -Parent $scriptPath
    }

    # Resolve config file path
    if ([string]::IsNullOrEmpty($ConfigFile)) {
        $ConfigFile = Join-Path $scriptDirectory "SCCMDeferralMonitorConfig.xml"
    }
    elseif (-not [System.IO.Path]::IsPathRooted($ConfigFile)) {
        $ConfigFile = Join-Path $scriptDirectory $ConfigFile
    }

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "SCCM Deferral Monitor - Data Collection" -ForegroundColor Cyan
    Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Load configuration
    $config = Load-Configuration -Path $ConfigFile

    $siteCode = $config.Configuration.Settings.SCCM.SiteCode
    $siteServer = $config.Configuration.Settings.SCCM.SiteServer
    $collectionID = $config.Configuration.Settings.SCCM.CollectionID
    $taskSequenceID = $config.Configuration.Settings.SCCM.TaskSequenceID
    $deferralLogPath = $config.Configuration.Settings.Logs.DeferralLogPath
    $tsLogDateMode = $config.Configuration.Settings.Logs.TSLogDateMode
    $tsLogLookbackDays = [int]$config.Configuration.Settings.Logs.TSLogLookbackDays
    $tsLogStartDate = $config.Configuration.Settings.Logs.TSLogStartDate
    $webRootPath = $config.Configuration.Settings.WebServer.WebRootPath
    $dataDirectoryConfig = $config.Configuration.Settings.WebServer.DataDirectory
    $logsDirectoryConfig = $config.Configuration.Settings.WebServer.LogsDirectory
    $win11Override = [System.Convert]::ToBoolean($config.Configuration.Settings.Monitoring.Windows11OverrideSuccess)
    $hostnameVerificationTimeout = [int]$config.Configuration.Settings.Monitoring.HostnameVerificationTimeoutSeconds
    $pingTimeout = [int]$config.Configuration.Settings.Monitoring.PingTimeoutMs

    # Store site server in script scope
    $script:sccmSiteServer = $siteServer

    # Create web root path if it doesn't exist
    if (-not (Test-Path $webRootPath)) {
        New-Item -Path $webRootPath -ItemType Directory -Force | Out-Null
        Write-Host "Created web root directory: $webRootPath" -ForegroundColor Green
    }

    # Resolve data directory (relative to web root or absolute)
    if ([System.IO.Path]::IsPathRooted($dataDirectoryConfig)) {
        $dataDirectory = $dataDirectoryConfig
    } else {
        $dataDirectory = Join-Path $webRootPath $dataDirectoryConfig
    }

    # Resolve logs directory (relative to web root or absolute)
    if ([System.IO.Path]::IsPathRooted($logsDirectoryConfig)) {
        $logsDirectory = $logsDirectoryConfig
    } else {
        $logsDirectory = Join-Path $webRootPath $logsDirectoryConfig
    }

    # Ensure directories exist
    if (-not (Test-Path $dataDirectory)) {
        New-Item -Path $dataDirectory -ItemType Directory -Force | Out-Null
        Write-Host "Created data directory: $dataDirectory" -ForegroundColor Green
    }

    if (-not (Test-Path $logsDirectory)) {
        New-Item -Path $logsDirectory -ItemType Directory -Force | Out-Null
        Write-Host "Created logs directory: $logsDirectory" -ForegroundColor Green
    }

    # Copy HTML file to web root if it doesn't exist or is outdated
    $sourceHtmlPath = Join-Path $scriptDirectory "SCCMDeferralMonitor.html"
    $destHtmlPath = Join-Path $webRootPath "index.html"

    if (Test-Path $sourceHtmlPath) {
        $shouldCopy = $false

        if (-not (Test-Path $destHtmlPath)) {
            $shouldCopy = $true
            Write-Host "HTML file not found in web root, copying..." -ForegroundColor Yellow
        } else {
            $sourceFile = Get-Item $sourceHtmlPath
            $destFile = Get-Item $destHtmlPath

            if ($sourceFile.LastWriteTime -gt $destFile.LastWriteTime) {
                $shouldCopy = $true
                Write-Host "HTML file in web root is outdated, updating..." -ForegroundColor Yellow
            }
        }

        if ($shouldCopy) {
            Copy-Item $sourceHtmlPath $destHtmlPath -Force
            Write-Host "HTML file copied to: $destHtmlPath" -ForegroundColor Green
        } else {
            Write-Host "HTML file in web root is up to date" -ForegroundColor Gray
        }
    } else {
        Write-Host "WARNING: Source HTML file not found: $sourceHtmlPath" -ForegroundColor Yellow
    }

    # Connect to SCCM
    if (-not (Connect-ToSCCM -SiteCode $siteCode -SiteServer $siteServer)) {
        throw "Failed to connect to SCCM"
    }

    # Get collection members
    $members = Get-CollectionMembers -CollectionID $collectionID

    if ($null -eq $members -or $members.Count -eq 0) {
        throw "No members found in collection $collectionID"
    }

    Write-Host "`nProcessing $($members.Count) collection members...`n" -ForegroundColor Cyan

    # Process each member
    $deviceDataList = @()
    $count = 0

    foreach ($member in $members) {
        $count++
        $computerName = $member.Name

        Write-Host "[$count/$($members.Count)] Processing: $computerName" -ForegroundColor Yellow

        # Check if online
        $isOnline = Test-ComputerOnline -ComputerName $computerName -TimeoutMs $pingTimeout
        Write-Host "  Online: $isOnline" -ForegroundColor $(if ($isOnline) { "Green" } else { "Red" })

        # Initialize variables
        $hostnameVerified = $false
        $actualHostname = $computerName
        $verificationError = ""

        # Verify hostname if online (critical for VPN scenarios)
        if ($isOnline) {
            Write-Host "  Verifying hostname..." -ForegroundColor Cyan
            $verification = Confirm-ComputerIdentity -ExpectedName $computerName -TimeoutSeconds $hostnameVerificationTimeout

            $hostnameVerified = $verification.IsValid
            $actualHostname = $verification.ActualName

            if ($hostnameVerified) {
                Write-Host "  Hostname verified: $actualHostname" -ForegroundColor Green
            }
            else {
                Write-Host "  Hostname verification failed: $($verification.ErrorMessage)" -ForegroundColor Red
                $verificationError = $verification.ErrorMessage

                # If hostname doesn't match, treat as offline for log collection
                if ($actualHostname -ne "Unknown" -and $actualHostname -ne $computerName) {
                    Write-Host "  WARNING: DNS mismatch detected! Expected '$computerName' but found '$actualHostname'" -ForegroundColor Yellow
                    Write-Host "  Skipping log collection to prevent pulling from wrong machine" -ForegroundColor Yellow
                }
            }
        }

        # Get primary user
        $primaryUser = Get-DevicePrimaryUser -ComputerName $computerName
        Write-Host "  Primary User: $primaryUser" -ForegroundColor Gray

        # Get OS version
        $osVersion = Get-DeviceOSVersion -ComputerName $computerName
        $isWindows11 = ($osVersion -eq "Windows 11")
        Write-Host "  OS: $osVersion" -ForegroundColor Gray

        # Get TS status
        $tsStatus = Get-TaskSequenceDeploymentStatus -ComputerName $computerName -TaskSequenceID $taskSequenceID

        # Apply Windows 11 override if enabled
        if ($win11Override -and $isWindows11 -and $tsStatus -ne "Success") {
            Write-Host "  TS Status: $tsStatus -> Success (Win11 Override)" -ForegroundColor Green
            $tsStatus = "Success"
        }
        else {
            Write-Host "  TS Status: $tsStatus" -ForegroundColor Gray
        }

        # Get deferral log data
        $deferralData = @{
            DeferralCount = "N/A"
            TSTriggerAttempted = "N/A"
            TSTriggerSuccess = "N/A"
            LogAvailable = $false
        }

        # Only collect logs if online AND hostname verified
        if ($isOnline -and $hostnameVerified) {
            $deferralData = Get-DeferralLogData -LogPath $deferralLogPath -ComputerName $computerName

            if ($deferralData.LogAvailable) {
                Write-Host "  Deferral Count: $($deferralData.DeferralCount)" -ForegroundColor Gray
                Write-Host "  TS Trigger Attempted: $($deferralData.TSTriggerAttempted)" -ForegroundColor Gray
                Write-Host "  TS Trigger Success: $($deferralData.TSTriggerSuccess)" -ForegroundColor Gray

                # Copy deferral log
                Copy-DeferralLog -ComputerName $computerName -SourcePath $deferralLogPath -DestinationDirectory $logsDirectory

                # Copy TS logs
                Get-TSLogsFromMachine -ComputerName $computerName -DestinationDirectory $logsDirectory -DateMode $tsLogDateMode -LookbackDays $tsLogLookbackDays -StartDate $tsLogStartDate
            }
            else {
                Write-Host "  Deferral log not available: $($deferralData.ErrorMessage)" -ForegroundColor Yellow
            }
        }
        elseif ($isOnline -and -not $hostnameVerified) {
            Write-Host "  Skipping log collection: Hostname not verified (DNS issue?)" -ForegroundColor Yellow
            $deferralData = @{
                DeferralCount = "N/A"
                TSTriggerAttempted = "N/A"
                TSTriggerSuccess = "N/A"
                LogAvailable = $false
                ErrorMessage = "Hostname verification failed: $verificationError"
            }
        }

        # Build device data object
        $deviceData = [PSCustomObject]@{
            DeviceName = $computerName
            PrimaryUser = $primaryUser
            IsOnline = $isOnline
            OnlineStatus = if ($isOnline) { "Online" } else { "Offline" }
            HostnameVerified = $hostnameVerified
            ActualHostname = $actualHostname
            VerificationError = $verificationError
            TSStatus = $tsStatus
            OSVersion = $osVersion
            IsWindows11 = $isWindows11
            DeferralCount = if ($deferralData.LogAvailable) { $deferralData.DeferralCount } else { "N/A" }
            TSTriggerAttempted = if ($deferralData.LogAvailable) { $deferralData.TSTriggerAttempted } else { "N/A" }
            TSTriggerSuccess = if ($deferralData.LogAvailable) { $deferralData.TSTriggerSuccess } else { "N/A" }
            LogAvailable = $deferralData.LogAvailable
            LastDeferralDate = if ($deferralData.LogAvailable) { $deferralData.LastDeferralDate } else { "N/A" }
            LastTriggerDate = if ($deferralData.LogAvailable) { $deferralData.LastTriggerDate } else { "N/A" }
            DeferralLogPath = "$computerName`_TaskSequenceDeferral.log"
            TSLogsPath = "$computerName`_TSLogs"
            LastUpdated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }

        $deviceDataList += $deviceData
        Write-Host ""
    }

    # Clean up orphaned logs
    Write-Host "Cleaning up orphaned log files..." -ForegroundColor Cyan
    Remove-OrphanedLogs -LogDirectory $logsDirectory -CurrentMembers $members

    # Export data to JSON
    $jsonOutputPath = Join-Path $dataDirectory "devicedata.json"
    Export-DataToJSON -DeviceData $deviceDataList -OutputPath $jsonOutputPath

    # Export metadata
    $metadata = [PSCustomObject]@{
        LastUpdate = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        TotalDevices = $deviceDataList.Count
        OnlineDevices = ($deviceDataList | Where-Object { $_.IsOnline }).Count
        OfflineDevices = ($deviceDataList | Where-Object { -not $_.IsOnline }).Count
        CollectionID = $collectionID
        TaskSequenceID = $taskSequenceID
        Windows11Override = $win11Override
    }

    $metadataPath = Join-Path $dataDirectory "metadata.json"
    $metadata | ConvertTo-Json | Out-File -FilePath $metadataPath -Encoding utf8 -Force

    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "Data collection completed successfully!" -ForegroundColor Green
    Write-Host "Total devices: $($deviceDataList.Count)" -ForegroundColor Green
    Write-Host "Online: $($metadata.OnlineDevices) | Offline: $($metadata.OfflineDevices)" -ForegroundColor Green
    Write-Host "Completed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Green
    Write-Host "========================================`n" -ForegroundColor Green

    exit 0
}
catch {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host "========================================`n" -ForegroundColor Red
    exit 1
}

#endregion
