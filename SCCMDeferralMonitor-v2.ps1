<#
.SYNOPSIS
    SCCM Task Sequence Deferral Monitor - Backend Data Collection Script (V2)

.DESCRIPTION
    Connects to SCCM, collects collection member data via jumpbox if needed,
    retrieves TS status from SCCM status messages, and generates JSON data files.

.NOTES
    Date: 2025-11-28
    Requires: PowerShell 5.1, ConfigurationManager PowerShell Module, SCCM Admin Rights

    V2 Changes:
    - Uses SCCM status messages for TS deployment status
    - Uses SCCM central log repository for TS logs
    - Supports jumpbox PSSession for VPN client access
    - Gets collection and TS names from SCCM
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
            $newDrive = New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $SiteServer -Scope Script -ErrorAction Stop
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

function Get-CollectionInfo {
    param([string]$CollectionID)

    try {
        $currentLocation = Get-Location
        Set-Location "$($script:sccmSiteCode):" -ErrorAction Stop

        $collection = Get-CMCollection -CollectionId $CollectionID -ErrorAction Stop

        Set-Location $currentLocation

        if ($collection) {
            return @{
                Name = $collection.Name
                ID = $collection.CollectionID
                MemberCount = $collection.MemberCount
            }
        }

        return $null
    }
    catch {
        Write-Host "Failed to get collection info: $_" -ForegroundColor Red
        if ($currentLocation) {
            Set-Location $currentLocation -ErrorAction SilentlyContinue
        }
        return $null
    }
}

function Get-TaskSequenceInfo {
    param([string]$TaskSequenceID)

    try {
        $currentLocation = Get-Location
        Set-Location "$($script:sccmSiteCode):" -ErrorAction Stop

        $ts = Get-CMTaskSequence -TaskSequencePackageId $TaskSequenceID -ErrorAction Stop

        Set-Location $currentLocation

        if ($ts) {
            return @{
                Name = $ts.Name
                PackageID = $ts.PackageID
                Description = $ts.Description
            }
        }

        return $null
    }
    catch {
        Write-Host "Failed to get task sequence info: $_" -ForegroundColor Red
        if ($currentLocation) {
            Set-Location $currentLocation -ErrorAction SilentlyContinue
        }
        return $null
    }
}

function Get-CollectionMembers {
    param([string]$CollectionID)

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

function Get-TaskSequenceDeploymentStatusFromSCCM {
    param(
        [string]$ResourceID,
        [string]$TaskSequenceID
    )

    try {
        $currentLocation = Get-Location
        Set-Location "$($script:sccmSiteCode):" -ErrorAction Stop

        # Query status messages for this computer and TS
        $query = @"
SELECT stat.RecordID, stat.MachineName, stat.MessageID, stat.Severity, stat.Time,
       stat.MessageState, stat.TopLevelSiteCode
FROM v_StatMsg stat
WHERE stat.RecordID IN (
    SELECT RecordID FROM v_TaskExecutionStatus
    WHERE ResourceID = '$ResourceID'
    AND PackageID = '$TaskSequenceID'
)
ORDER BY stat.Time DESC
"@

        $statusMessages = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" `
            -ComputerName $script:sccmSiteServer `
            -Query $query `
            -ErrorAction SilentlyContinue

        Set-Location $currentLocation

        if ($statusMessages -and $statusMessages.Count -gt 0) {
            # Get most recent message
            $latestMessage = $statusMessages[0]

            # Determine status based on message
            switch ($latestMessage.MessageState) {
                1 { return "Success" }
                2 { return "In Progress" }
                3 { return "Failed" }
                default { return "Unknown" }
            }
        }

        return "Not Started"
    }
    catch {
        Write-Host "Error getting TS status from SCCM: $_" -ForegroundColor Yellow
        if ($currentLocation) {
            Set-Location $currentLocation -ErrorAction SilentlyContinue
        }
        return "Unknown"
    }
}

function Get-AndSaveTaskSequenceStatusMessages {
    <#
    .SYNOPSIS
        Downloads ALL TS status messages from SCCM for a computer and saves to file
    #>
    param(
        [string]$ComputerName,
        [string]$ResourceID,
        [string]$TaskSequenceID,
        [string]$DestinationDirectory
    )

    try {
        Write-Host "  Downloading TS status messages from SCCM..." -ForegroundColor Cyan

        # Query ALL status messages for this resource and TS
        $query = @"
SELECT stat.Time, stat.MachineName, stat.MessageID, stat.MessageType, stat.Severity,
       stat.MessageState, stat.Component, stat.InsString1, stat.InsString2, stat.InsString3,
       stat.InsString4, stat.InsString5, stat.InsString6, stat.InsString7, stat.InsString8,
       stat.InsString9, stat.InsString10
FROM v_StatusMessage stat
INNER JOIN v_TaskExecutionStatus tes ON stat.RecordID = tes.StatusMessageID
WHERE tes.ResourceID = '$ResourceID'
  AND tes.PackageID = '$TaskSequenceID'
ORDER BY stat.Time DESC
"@

        $messages = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" `
            -ComputerName $script:sccmSiteServer `
            -Query $query `
            -ErrorAction Stop

        if ($messages -and $messages.Count -gt 0) {
            # Convert to structured format
            $messageList = @()

            foreach ($msg in $messages) {
                $messageList += [PSCustomObject]@{
                    Time = if ($msg.Time) { [System.Management.ManagementDateTimeConverter]::ToDateTime($msg.Time) } else { $null }
                    MachineName = $msg.MachineName
                    MessageID = $msg.MessageID
                    MessageType = $msg.MessageType
                    Severity = switch ($msg.Severity) {
                        0 { "Info" }
                        1 { "Warning" }
                        2 { "Error" }
                        default { "Unknown" }
                    }
                    State = $msg.MessageState
                    Component = $msg.Component
                    Details = @($msg.InsString1, $msg.InsString2, $msg.InsString3, $msg.InsString4,
                                $msg.InsString5, $msg.InsString6, $msg.InsString7, $msg.InsString8,
                                $msg.InsString9, $msg.InsString10) | Where-Object { -not [string]::IsNullOrEmpty($_) }
                }
            }

            # Save to JSON file
            $destFile = Join-Path $DestinationDirectory "$ComputerName`_TSStatusMessages.json"
            $messageList | ConvertTo-Json -Depth 5 | Out-File -FilePath $destFile -Encoding utf8 -Force

            Write-Host "  Saved $($messages.Count) TS status messages to file" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "  No TS status messages found in SCCM" -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Host "  Failed to get TS status messages: $_" -ForegroundColor Red
        return $false
    }
}

function Get-DevicePrimaryUser {
    param([string]$ComputerName)

    try {
        $currentLocation = Get-Location
        Set-Location "$($script:sccmSiteCode):" -ErrorAction Stop

        $device = Get-CMDevice -Name $ComputerName -ErrorAction SilentlyContinue

        Set-Location $currentLocation

        if ($device -and $device.PrimaryUser) {
            if ($device.PrimaryUser -match '\\(.+)$') {
                return $matches[1]
            }
            return $device.PrimaryUser
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
    param([string]$ComputerName)

    try {
        $currentLocation = Get-Location
        Set-Location "$($script:sccmSiteCode):" -ErrorAction Stop

        $device = Get-CMDevice -Name $ComputerName -ErrorAction SilentlyContinue

        Set-Location $currentLocation

        if ($device -and $device.OSVersion) {
            # Windows 11 builds: 22000, 22621, 26100, etc. (all 10.0.2xxxx and above)
            if ($device.OSVersion -match '10\.0\.(2[2-9]\d{3}|[3-9]\d{4})') {
                return "Windows 11"
            }
            return $device.OSVersion
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

#region Jumpbox Functions

$script:jumpboxSession = $null

function Connect-ToJumpbox {
    param([string]$JumpboxServer)

    try {
        Write-Host "Connecting to jumpbox: $JumpboxServer" -ForegroundColor Cyan

        # Close existing session if any
        if ($script:jumpboxSession) {
            Remove-PSSession $script:jumpboxSession -ErrorAction SilentlyContinue
        }

        $script:jumpboxSession = New-PSSession -ComputerName $JumpboxServer -ErrorAction Stop

        Write-Host "Connected to jumpbox successfully" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect to jumpbox: $_" -ForegroundColor Red
        return $false
    }
}

function Disconnect-FromJumpbox {
    if ($script:jumpboxSession) {
        Remove-PSSession $script:jumpboxSession -ErrorAction SilentlyContinue
        $script:jumpboxSession = $null
        Write-Host "Disconnected from jumpbox" -ForegroundColor Gray
    }
}

function Test-ComputerOnlineViaJumpbox {
    param(
        [string]$ComputerName,
        [int]$TimeoutMs = 1000
    )

    if (-not $script:jumpboxSession) {
        return $false
    }

    try {
        $result = Invoke-Command -Session $script:jumpboxSession -ScriptBlock {
            param($computer, $timeout)
            try {
                $ping = New-Object System.Net.NetworkInformation.Ping
                $pingResult = $ping.Send($computer, $timeout)
                return ($pingResult.Status -eq 'Success')
            }
            catch {
                return $false
            }
        } -ArgumentList $ComputerName, $TimeoutMs

        return $result
    }
    catch {
        return $false
    }
}

function Confirm-ComputerIdentityViaJumpbox {
    param(
        [string]$ExpectedName,
        [int]$TimeoutSeconds = 3
    )

    if (-not $script:jumpboxSession) {
        return @{
            IsValid = $false
            ActualName = "Unknown"
            ErrorMessage = "No jumpbox session"
        }
    }

    try {
        # Use simple Get-WmiObject from jumpbox
        $result = Invoke-Command -Session $script:jumpboxSession -ScriptBlock {
            param($computerName, $timeout)

            $verifyResult = @{
                IsValid = $false
                ActualName = "Unknown"
                ErrorMessage = ""
            }

            try {
                # Simple WMI query with timeout
                $wmiTimeout = $timeout * 1000  # Convert to milliseconds
                $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $computerName -ErrorAction Stop -AsJob | Wait-Job -Timeout $timeout | Receive-Job

                if ($computerSystem) {
                    $verifyResult.ActualName = $computerSystem.Name

                    if ($verifyResult.ActualName -eq $computerName) {
                        $verifyResult.IsValid = $true
                    }
                    else {
                        $verifyResult.ErrorMessage = "Hostname mismatch: Expected '$computerName', got '$($verifyResult.ActualName)'"
                    }
                }
                else {
                    $verifyResult.ErrorMessage = "WMI query returned null"
                }
            }
            catch {
                $verifyResult.ErrorMessage = "WMI error: $($_.Exception.Message)"
            }

            return $verifyResult
        } -ArgumentList $ExpectedName, $TimeoutSeconds

        return $result
    }
    catch {
        return @{
            IsValid = $false
            ActualName = "Unknown"
            ErrorMessage = "Jumpbox command failed: $($_.Exception.Message)"
        }
    }
}

function Get-DeferralLogDataViaJumpbox {
    param(
        [string]$LogPath,
        [string]$ComputerName
    )

    if (-not $script:jumpboxSession) {
        return @{
            DeferralCount = 0
            TSTriggerAttempted = $false
            TSTriggerSuccess = $false
            LastDeferralDate = "N/A"
            LastTriggerDate = "N/A"
            LogAvailable = $false
            ErrorMessage = "No jumpbox session"
        }
    }

    try {
        $result = Invoke-Command -Session $script:jumpboxSession -ScriptBlock {
            param($computer, $logPathLocal)

            $logResult = @{
                DeferralCount = 0
                TSTriggerAttempted = $false
                TSTriggerSuccess = $false
                LastDeferralDate = "N/A"
                LastTriggerDate = "N/A"
                LogAvailable = $false
                ErrorMessage = ""
            }

            try {
                $uncPath = "\\$computer\$($logPathLocal.Replace(':', '$'))"

                if (-not (Test-Path $uncPath)) {
                    $logResult.ErrorMessage = "Log file not found"
                    return $logResult
                }

                $logContent = Get-Content -Path $uncPath -ErrorAction Stop

                if ($logContent.Count -eq 0) {
                    $logResult.ErrorMessage = "Log file is empty"
                    return $logResult
                }

                $logResult.LogAvailable = $true

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

                $logResult.DeferralCount = $deferrals
                $logResult.TSTriggerAttempted = $triggerAttempted
                $logResult.TSTriggerSuccess = $triggerSuccess

                if ($lastDeferralDate) {
                    $logResult.LastDeferralDate = $lastDeferralDate.ToString("yyyy-MM-dd HH:mm:ss")
                }

                if ($lastTriggerDate) {
                    $logResult.LastTriggerDate = $lastTriggerDate.ToString("yyyy-MM-dd HH:mm:ss")
                }
            }
            catch {
                $logResult.ErrorMessage = $_.Exception.Message
            }

            return $logResult
        } -ArgumentList $ComputerName, $LogPath

        return $result
    }
    catch {
        return @{
            DeferralCount = 0
            TSTriggerAttempted = $false
            TSTriggerSuccess = $false
            LastDeferralDate = "N/A"
            LastTriggerDate = "N/A"
            LogAvailable = $false
            ErrorMessage = "Jumpbox command failed: $($_.Exception.Message)"
        }
    }
}

function Copy-DeferralLogViaJumpbox {
    param(
        [string]$ComputerName,
        [string]$SourcePath,
        [string]$DestinationDirectory
    )

    if (-not $script:jumpboxSession) {
        return $false
    }

    try {
        # Copy from client to jumpbox temp location using C$ share
        $tempPath = Invoke-Command -Session $script:jumpboxSession -ScriptBlock {
            param($computer, $sourcePath)

            # Build UNC path using C$ admin share
            # Example: \\COMPUTER\C$\Windows\ccm\logs\TaskSequenceDeferral.log
            $uncPath = "\\$computer\" + $sourcePath.Replace(':', '$')

            Write-Verbose "Attempting to access: $uncPath"

            if (-not (Test-Path $uncPath)) {
                Write-Warning "Log file not found: $uncPath"
                return $null
            }

            # Check if we need to update (compare with existing temp file)
            $tempFile = "$env:TEMP\$computer`_TaskSequenceDeferral.log"

            # Always copy to get latest version
            Copy-Item $uncPath $tempFile -Force -ErrorAction Stop
            Write-Verbose "Copied to temp: $tempFile"

            return $tempFile
        } -ArgumentList $ComputerName, $SourcePath -Verbose

        if ($tempPath) {
            # Copy from jumpbox temp to local destination
            $destFile = Join-Path $DestinationDirectory "$ComputerName`_TaskSequenceDeferral.log"

            # Only copy if destination doesn't exist or source is newer
            $shouldCopy = $false
            if (-not (Test-Path $destFile)) {
                $shouldCopy = $true
            }
            else {
                # Get file times from jumpbox
                $sourceTime = Invoke-Command -Session $script:jumpboxSession -ScriptBlock {
                    param($path)
                    (Get-Item $path).LastWriteTime
                } -ArgumentList $tempPath

                $destTime = (Get-Item $destFile).LastWriteTime

                if ($sourceTime -gt $destTime) {
                    $shouldCopy = $true
                }
            }

            if ($shouldCopy) {
                Copy-Item -FromSession $script:jumpboxSession -Path $tempPath -Destination $destFile -Force
                Write-Host "  Copied deferral log from $ComputerName" -ForegroundColor Green
            }
            else {
                Write-Host "  Deferral log already up to date for $ComputerName" -ForegroundColor Gray
            }

            # Clean up temp file on jumpbox
            Invoke-Command -Session $script:jumpboxSession -ScriptBlock {
                param($temp)
                Remove-Item $temp -Force -ErrorAction SilentlyContinue
            } -ArgumentList $tempPath

            return $true
        }

        Write-Host "  No deferral log found for $ComputerName" -ForegroundColor Yellow
        return $false
    }
    catch {
        Write-Host "  Failed to copy deferral log from ${ComputerName}: $_" -ForegroundColor Red
        return $false
    }
}

#endregion

#region Direct Connection Functions (fallback)

function Test-ComputerOnlineDirect {
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

        # Clean up TS status message JSON files
        $tsStatusFiles = Get-ChildItem -Path $LogDirectory -Filter "*_TSStatusMessages.json" -ErrorAction SilentlyContinue
        foreach ($file in $tsStatusFiles) {
            $computerName = $file.Name -replace '_TSStatusMessages\.json$', ''
            if ($computerName -notin $memberNames) {
                Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue
                Write-Host "Removed orphaned TS status messages: $($file.Name)" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "Error during orphaned log cleanup: $_" -ForegroundColor Red
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
    Write-Host "SCCM Deferral Monitor - Data Collection V2" -ForegroundColor Cyan
    Write-Host "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Load configuration
    $config = Load-Configuration -Path $ConfigFile

    $siteCode = $config.Configuration.Settings.SCCM.SiteCode
    $siteServer = $config.Configuration.Settings.SCCM.SiteServer
    $collectionID = $config.Configuration.Settings.SCCM.CollectionID
    $taskSequenceID = $config.Configuration.Settings.SCCM.TaskSequenceID
    $deferralLogPath = $config.Configuration.Settings.Logs.DeferralLogPath
    $webRootPath = $config.Configuration.Settings.WebServer.WebRootPath
    $dataDirectoryConfig = $config.Configuration.Settings.WebServer.DataDirectory
    $logsDirectoryConfig = $config.Configuration.Settings.WebServer.LogsDirectory
    $win11Override = [System.Convert]::ToBoolean($config.Configuration.Settings.Monitoring.Windows11OverrideSuccess)
    $hostnameVerificationTimeout = [int]$config.Configuration.Settings.Monitoring.HostnameVerificationTimeoutSeconds
    $pingTimeout = [int]$config.Configuration.Settings.Monitoring.PingTimeoutMs
    $useJumpbox = [System.Convert]::ToBoolean($config.Configuration.Settings.Monitoring.UseJumpbox)
    $jumpboxServer = $config.Configuration.Settings.Monitoring.JumpboxServer

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

    # Copy HTML file to web root
    $sourceHtmlPath = Join-Path $scriptDirectory "SCCMDeferralMonitor.html"
    $destHtmlPath = Join-Path $webRootPath "index.html"

    if (Test-Path $sourceHtmlPath) {
        if (-not (Test-Path $destHtmlPath) -or ((Get-Item $sourceHtmlPath).LastWriteTime -gt (Get-Item $destHtmlPath).LastWriteTime)) {
            Copy-Item $sourceHtmlPath $destHtmlPath -Force
            Write-Host "HTML file copied to: $destHtmlPath" -ForegroundColor Green
        }
    }

    # Connect to SCCM
    if (-not (Connect-ToSCCM -SiteCode $siteCode -SiteServer $siteServer)) {
        throw "Failed to connect to SCCM"
    }

    # Get collection and TS information
    Write-Host "`nGetting collection and task sequence information..." -ForegroundColor Cyan
    $collectionInfo = Get-CollectionInfo -CollectionID $collectionID
    $tsInfo = Get-TaskSequenceInfo -TaskSequenceID $taskSequenceID

    if (-not $collectionInfo) {
        throw "Failed to get collection information for $collectionID"
    }

    if (-not $tsInfo) {
        throw "Failed to get task sequence information for $taskSequenceID"
    }

    Write-Host "Collection: $($collectionInfo.Name) ($($collectionInfo.ID))" -ForegroundColor Green
    Write-Host "Task Sequence: $($tsInfo.Name) ($($tsInfo.PackageID))" -ForegroundColor Green

    # Connect to jumpbox if needed
    if ($useJumpbox -and -not [string]::IsNullOrEmpty($jumpboxServer)) {
        if (-not (Connect-ToJumpbox -JumpboxServer $jumpboxServer)) {
            throw "Failed to connect to jumpbox: $jumpboxServer"
        }
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
        $resourceID = $member.ResourceID

        Write-Host "[$count/$($members.Count)] Processing: $computerName" -ForegroundColor Yellow

        # Check if online
        if ($useJumpbox -and $script:jumpboxSession) {
            $isOnline = Test-ComputerOnlineViaJumpbox -ComputerName $computerName -TimeoutMs $pingTimeout
        }
        else {
            $isOnline = Test-ComputerOnlineDirect -ComputerName $computerName -TimeoutMs $pingTimeout
        }

        Write-Host "  Online: $isOnline" -ForegroundColor $(if ($isOnline) { "Green" } else { "Red" })

        # Hostname verification
        $hostnameVerified = $false
        $actualHostname = $computerName
        $verificationError = ""

        if ($isOnline) {
            Write-Host "  Verifying hostname..." -ForegroundColor Cyan

            if ($useJumpbox -and $script:jumpboxSession) {
                $verification = Confirm-ComputerIdentityViaJumpbox -ExpectedName $computerName -TimeoutSeconds $hostnameVerificationTimeout
            }
            else {
                # Direct verification - implement if needed
                $verification = @{ IsValid = $true; ActualName = $computerName; ErrorMessage = "" }
            }

            $hostnameVerified = $verification.IsValid
            $actualHostname = $verification.ActualName
            $verificationError = $verification.ErrorMessage

            if ($hostnameVerified) {
                Write-Host "  Hostname verified: $actualHostname" -ForegroundColor Green
            }
            else {
                Write-Host "  Hostname verification failed: $verificationError" -ForegroundColor Red
            }
        }

        # Get info from SCCM
        $primaryUser = Get-DevicePrimaryUser -ComputerName $computerName
        Write-Host "  Primary User: $primaryUser" -ForegroundColor Gray

        $osVersion = Get-DeviceOSVersion -ComputerName $computerName
        $isWindows11 = ($osVersion -eq "Windows 11")
        Write-Host "  OS: $osVersion" -ForegroundColor Gray

        # Get TS status from SCCM status messages
        $tsStatus = Get-TaskSequenceDeploymentStatusFromSCCM -ResourceID $resourceID -TaskSequenceID $taskSequenceID

        # Apply Windows 11 override
        if ($win11Override -and $isWindows11 -and $tsStatus -ne "Success") {
            Write-Host "  TS Status: $tsStatus -> Success (Win11 Override)" -ForegroundColor Green
            $tsStatus = "Success"
        }
        else {
            Write-Host "  TS Status: $tsStatus" -ForegroundColor Gray
        }

        # Download and save TS status messages from SCCM
        $tsMessagesDownloaded = Get-AndSaveTaskSequenceStatusMessages -ComputerName $computerName -ResourceID $resourceID -TaskSequenceID $taskSequenceID -DestinationDirectory $logsDirectory

        # Get deferral log data
        $deferralData = @{
            DeferralCount = "N/A"
            TSTriggerAttempted = "N/A"
            TSTriggerSuccess = "N/A"
            LogAvailable = $false
        }

        if ($isOnline -and $hostnameVerified) {
            if ($useJumpbox -and $script:jumpboxSession) {
                $deferralData = Get-DeferralLogDataViaJumpbox -LogPath $deferralLogPath -ComputerName $computerName

                if ($deferralData.LogAvailable) {
                    Write-Host "  Deferral Count: $($deferralData.DeferralCount)" -ForegroundColor Gray
                    Write-Host "  TS Trigger Attempted: $($deferralData.TSTriggerAttempted)" -ForegroundColor Gray
                    Write-Host "  TS Trigger Success: $($deferralData.TSTriggerSuccess)" -ForegroundColor Gray

                    # Copy deferral log
                    Copy-DeferralLogViaJumpbox -ComputerName $computerName -SourcePath $deferralLogPath -DestinationDirectory $logsDirectory
                }
            }
        }

        # Build device data object
        $deviceData = [PSCustomObject]@{
            DeviceName = $computerName
            ResourceID = $resourceID
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
            TSStatusMessagesPath = if ($tsMessagesDownloaded) { "$computerName`_TSStatusMessages.json" } else { "" }
            TSStatusMessagesAvailable = $tsMessagesDownloaded
            LastUpdated = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
        }

        $deviceDataList += $deviceData
        Write-Host ""
    }

    # Disconnect from jumpbox
    Disconnect-FromJumpbox

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
        CollectionName = $collectionInfo.Name
        TaskSequenceID = $taskSequenceID
        TaskSequenceName = $tsInfo.Name
        Windows11Override = $win11Override
        UseJumpbox = $useJumpbox
        JumpboxServer = if ($useJumpbox) { $jumpboxServer } else { "Direct Connection" }
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
    # Ensure jumpbox is disconnected on error
    Disconnect-FromJumpbox

    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "ERROR: $_" -ForegroundColor Red
    Write-Host "========================================`n" -ForegroundColor Red
    exit 1
}

#endregion
