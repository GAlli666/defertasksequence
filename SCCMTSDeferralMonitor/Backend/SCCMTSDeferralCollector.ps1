<#
.SYNOPSIS
    SCCM Task Sequence Deferral Monitor - Backend Data Collector

.DESCRIPTION
    Connects to SCCM, collects collection member data, retrieves TS execution status,
    downloads TaskSequenceDeferral.log files, and generates JSON metadata files.

    This is the backend data collection component. Use the WPF viewer to display the data.

.NOTES
    Date: 2025-11-28
    Requires: PowerShell 5.1, ConfigurationManager PowerShell Module, SCCM Admin Rights

    Features:
    - SCCM data collection (OS version, primary user, TS status)
    - Direct ping/WMI hostname verification
    - C$ share access for log collection
    - TS execution status from vSMS_TaskSequenceExecutionStatus
    - Windows 11 override support
    - Orphaned log cleanup
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
        [string]$TaskSequenceID,
        [string]$ComputerName
    )

    try {
        Write-Host "  [DEBUG] Getting TS deployment status for $ComputerName" -ForegroundColor Cyan
        Write-Host "  [DEBUG] ResourceID: $ResourceID, TaskSequenceID: $TaskSequenceID" -ForegroundColor Gray

        $currentLocation = Get-Location
        Set-Location "$($script:sccmSiteCode):" -ErrorAction Stop
        Write-Host "  [DEBUG] Changed location to site drive: $($script:sccmSiteCode):" -ForegroundColor Gray

        # Get task sequence name
        $ts = Get-CMTaskSequence -TaskSequencePackageId $TaskSequenceID -ErrorAction SilentlyContinue
        if (-not $ts) {
            Write-Host "  [DEBUG] Task sequence not found for ID: $TaskSequenceID" -ForegroundColor Yellow
            Set-Location $currentLocation
            return "Unknown"
        }

        Write-Host "  [DEBUG] Task Sequence Name: $($ts.Name)" -ForegroundColor Gray

        # Use cmdlet to get deployment status for this specific TS
        $deployments = Get-CMDeployment -SoftwareName $ts.Name -ErrorAction SilentlyContinue

        if (-not $deployments) {
            Write-Host "  [DEBUG] No deployments found for TS: $($ts.Name)" -ForegroundColor Yellow
            Set-Location $currentLocation
            return "Not Started"
        }

        Write-Host "  [DEBUG] Found $(@($deployments).Count) deployment(s)" -ForegroundColor Gray

        foreach ($deployment in $deployments) {
            Write-Host "  [DEBUG] Checking deployment ID: $($deployment.DeploymentID)" -ForegroundColor Gray

            # Get deployment status for this specific computer
            $statuses = Get-CMDeploymentStatus -DeploymentId $deployment.DeploymentID -ErrorAction SilentlyContinue

            if ($statuses) {
                Write-Host "  [DEBUG] Found $(@($statuses).Count) status record(s) for deployment" -ForegroundColor Gray

                # Try to find status for this computer
                $status = $statuses | Where-Object { $_.DeviceName -eq $ComputerName }

                if ($status) {
                    Write-Host "  [DEBUG] Found status for $ComputerName - StatusType: $($status.StatusType)" -ForegroundColor Green
                    Set-Location $currentLocation

                    # Map status to our values
                    switch ($status.StatusType) {
                        1 { return "Success" }
                        2 { return "In Progress" }
                        3 { return "Requirements Not Met" }
                        4 { return "Unknown" }
                        5 { return "Failed" }
                        default {
                            Write-Host "  [DEBUG] Unmapped StatusType: $($status.StatusType)" -ForegroundColor Yellow
                            return "Unknown"
                        }
                    }
                }
                else {
                    Write-Host "  [DEBUG] No status found for computer: $ComputerName in deployment $($deployment.DeploymentID)" -ForegroundColor Yellow
                }
            }
            else {
                Write-Host "  [DEBUG] No status records returned for deployment $($deployment.DeploymentID)" -ForegroundColor Yellow
            }
        }

        Write-Host "  [DEBUG] No matching status found across all deployments - returning 'Not Started'" -ForegroundColor Yellow
        Set-Location $currentLocation
        return "Not Started"
    }
    catch {
        Write-Host "  [ERROR] Exception in Get-TaskSequenceDeploymentStatusFromSCCM: $_" -ForegroundColor Red
        Write-Host "  [ERROR] Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        if ($currentLocation) {
            Set-Location $currentLocation -ErrorAction SilentlyContinue
        }
        return "Unknown"
    }
}

function Get-AndSaveTaskSequenceStatusMessages {
    <#
    .SYNOPSIS
        Downloads ALL TS execution status from SCCM for a computer and saves to JSON and HTML files
        Uses simple WQL queries (no JOINs) to get execution history with full ActionOutput
    #>
    param(
        [string]$ComputerName,
        [string]$ResourceID,
        [string]$TaskSequenceID,
        [string]$DestinationDirectory
    )

    try {
        Write-Host "  Downloading TS execution status from SCCM..." -ForegroundColor Cyan
        Write-Host "  [DEBUG] Parameters - Computer: $ComputerName, ResourceID: $ResourceID, TSID: $TaskSequenceID" -ForegroundColor Gray

        # Simple WQL query - no INNER JOIN (not supported in WQL)
        # Query the execution status view directly
        $query = "SELECT * FROM SMS_TaskExecutionStatus WHERE ResourceID = '$ResourceID' AND PackageID = '$TaskSequenceID'"
        Write-Host "  [DEBUG] WQL Query: $query" -ForegroundColor Gray
        Write-Host "  [DEBUG] Namespace: ROOT\SMS\site_$($script:sccmSiteCode), Server: $script:sccmSiteServer" -ForegroundColor Gray

        $messages = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" `
            -ComputerName $script:sccmSiteServer `
            -Query $query `
            -ErrorAction SilentlyContinue

        if ($messages -and $messages.Count -gt 0) {
            Write-Host "  [DEBUG] Found $($messages.Count) execution status message(s)" -ForegroundColor Green
            # Convert to structured format with full details
            $executionSteps = @()

            foreach ($msg in $messages) {
                # Get detailed status message if available
                $statusMsgQuery = "SELECT * FROM SMS_StatusMessage WHERE RecordID = '$($msg.StatusMessageID)'"
                $statusMsg = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" `
                    -ComputerName $script:sccmSiteServer `
                    -Query $statusMsgQuery `
                    -ErrorAction SilentlyContinue

                # Format execution time
                $executionTime = if ($msg.ExecutionTime) {
                    try {
                        [System.Management.ManagementDateTimeConverter]::ToDateTime($msg.ExecutionTime).ToString("yyyy-MM-dd HH:mm:ss")
                    }
                    catch { $msg.ExecutionTime }
                } else { "" }

                $executionSteps += [PSCustomObject]@{
                    ComputerName = $ComputerName
                    ExecutionTime = $executionTime
                    Step = if ($msg.Step) { $msg.Step } else { 0 }
                    ActionName = if ($msg.ActionName) { $msg.ActionName } else { "" }
                    GroupName = if ($msg.GroupName) { $msg.GroupName } else { "" }
                    LastStatusMsgName = if ($msg.LastStatusMessageName) { $msg.LastStatusMessageName } else { "" }
                    StatusMessageID = if ($statusMsg) { $statusMsg.MessageID } else { "" }
                    ExitCode = if ($null -ne $msg.ExitCode) { $msg.ExitCode } else { "" }
                    ActionOutput = if ($msg.ActionOutput) { $msg.ActionOutput } else { "" }  # Full output, no truncation
                }
            }

            # Sort by step number ascending (chronological order)
            $executionSteps = $executionSteps | Sort-Object Step

            # Save to JSON file
            $jsonFile = Join-Path $DestinationDirectory "$ComputerName`_TSStatusMessages.json"
            $executionSteps | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonFile -Encoding utf8 -Force

            # Create HTML log file
            $htmlFile = Join-Path $DestinationDirectory "$ComputerName`_TSExecutionLog.html"

            $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>TS Execution Log - $ComputerName</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            color: #2c3e50;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            background-color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin-bottom: 30px;
        }
        th {
            background-color: #3498db;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
            position: sticky;
            top: 0;
        }
        td {
            padding: 10px;
            border-bottom: 1px solid #e0e0e0;
        }
        tr:hover {
            background-color: #f8f9fa;
        }
        .action-output {
            background-color: #f8f9fa;
            padding: 15px;
            margin: 10px 0;
            border-left: 4px solid #3498db;
            font-family: 'Consolas', 'Courier New', monospace;
            font-size: 12px;
            white-space: pre-wrap;
            word-wrap: break-word;
            max-width: 100%;
            overflow-x: auto;
        }
        .step-row {
            font-weight: 600;
        }
        .success {
            color: #27ae60;
        }
        .failed {
            color: #e74c3c;
        }
        .info {
            color: #95a5a6;
            font-size: 11px;
        }
    </style>
</head>
<body>
    <h1>Task Sequence Execution Log</h1>
    <div class="info">
        <p><strong>Computer:</strong> $ComputerName</p>
        <p><strong>Task Sequence ID:</strong> $TaskSequenceID</p>
        <p><strong>Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
        <p><strong>Total Steps:</strong> $($executionSteps.Count)</p>
    </div>
    <table>
        <thead>
            <tr>
                <th>Step</th>
                <th>Execution Time</th>
                <th>Action Name</th>
                <th>Group Name</th>
                <th>Status Message</th>
                <th>Exit Code</th>
            </tr>
        </thead>
        <tbody>
"@

            foreach ($step in $executionSteps) {
                $exitCodeClass = if ($step.ExitCode -eq 0) { "success" } elseif ($step.ExitCode -ne "") { "failed" } else { "" }

                $htmlContent += @"
            <tr>
                <td class="step-row">$($step.Step)</td>
                <td>$($step.ExecutionTime)</td>
                <td>$($step.ActionName)</td>
                <td>$($step.GroupName)</td>
                <td>$($step.LastStatusMsgName)</td>
                <td class="$exitCodeClass">$($step.ExitCode)</td>
            </tr>
"@

                if (-not [string]::IsNullOrWhiteSpace($step.ActionOutput)) {
                    # Escape HTML characters in output
                    $escapedOutput = [System.Web.HttpUtility]::HtmlEncode($step.ActionOutput)
                    $htmlContent += @"
            <tr>
                <td colspan="6">
                    <strong>Action Output:</strong>
                    <div class="action-output">$escapedOutput</div>
                </td>
            </tr>
"@
                }
            }

            $htmlContent += @"
        </tbody>
    </table>
</body>
</html>
"@

            $htmlContent | Out-File -FilePath $htmlFile -Encoding utf8 -Force

            Write-Host "  Saved $($executionSteps.Count) TS execution steps (JSON + HTML)" -ForegroundColor Green
            return $true
        }
        else {
            Write-Host "  [DEBUG] No TS execution data found in SCCM for query" -ForegroundColor Yellow
            Write-Host "  [DEBUG] Messages variable state - IsNull: $($null -eq $messages), Count: $($messages.Count)" -ForegroundColor Yellow
            return $false
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to get TS execution status: $_" -ForegroundColor Red
        Write-Host "  [ERROR] Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
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

#region Direct Connection Functions

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
    param(
        [string]$ExpectedName,
        [int]$TimeoutSeconds = 3
    )

    $verifyResult = @{
        IsValid = $false
        ActualName = "Unknown"
        ErrorMessage = ""
    }

    try {
        # Simple WMI query with timeout
        $computerSystem = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ExpectedName -ErrorAction Stop -AsJob | Wait-Job -Timeout $TimeoutSeconds | Receive-Job

        if ($computerSystem) {
            $verifyResult.ActualName = $computerSystem.Name

            if ($verifyResult.ActualName -eq $ExpectedName) {
                $verifyResult.IsValid = $true
            }
            else {
                $verifyResult.ErrorMessage = "Hostname mismatch: Expected '$ExpectedName', got '$($verifyResult.ActualName)'"
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
}

function Get-DeferralLogData {
    param(
        [string]$LogPath,
        [string]$ComputerName
    )

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
        # Build UNC path using C$ admin share
        $uncPath = "\\$ComputerName\$($LogPath.Replace(':', '$'))"

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
}

function Copy-DeferralLog {
    param(
        [string]$ComputerName,
        [string]$SourcePath,
        [string]$DestinationDirectory
    )

    try {
        # Build UNC path using C$ admin share
        $uncPath = "\\$ComputerName\$($SourcePath.Replace(':', '$'))"

        if (-not (Test-Path $uncPath)) {
            Write-Host "  No deferral log found for $ComputerName" -ForegroundColor Yellow
            return $false
        }

        $destFile = Join-Path $DestinationDirectory "$ComputerName`_TaskSequenceDeferral.log"

        # Only copy if destination doesn't exist or source is newer
        $shouldCopy = $false
        if (-not (Test-Path $destFile)) {
            $shouldCopy = $true
        }
        else {
            $sourceTime = (Get-Item $uncPath).LastWriteTime
            $destTime = (Get-Item $destFile).LastWriteTime

            if ($sourceTime -gt $destTime) {
                $shouldCopy = $true
            }
        }

        if ($shouldCopy) {
            Copy-Item $uncPath $destFile -Force -ErrorAction Stop
            Write-Host "  Copied deferral log from $ComputerName" -ForegroundColor Green
        }
        else {
            Write-Host "  Deferral log already up to date for $ComputerName" -ForegroundColor Gray
        }

        return $true
    }
    catch {
        Write-Host "  Failed to copy deferral log from ${ComputerName}: $_" -ForegroundColor Red
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
    Write-Host "SCCM TS Deferral Monitor - Backend Collector" -ForegroundColor Cyan
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
        $isOnline = Test-ComputerOnline -ComputerName $computerName -TimeoutMs $pingTimeout

        Write-Host "  Online: $isOnline" -ForegroundColor $(if ($isOnline) { "Green" } else { "Red" })

        # Hostname verification
        $hostnameVerified = $false
        $actualHostname = $computerName
        $verificationError = ""

        if ($isOnline) {
            Write-Host "  Verifying hostname..." -ForegroundColor Cyan

            $verification = Confirm-ComputerIdentity -ExpectedName $computerName -TimeoutSeconds $hostnameVerificationTimeout

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

        # Get TS status from SCCM deployment status
        $tsStatus = Get-TaskSequenceDeploymentStatusFromSCCM -ResourceID $resourceID -TaskSequenceID $taskSequenceID -ComputerName $computerName

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
            $deferralData = Get-DeferralLogData -LogPath $deferralLogPath -ComputerName $computerName

            if ($deferralData.LogAvailable) {
                Write-Host "  Deferral Count: $($deferralData.DeferralCount)" -ForegroundColor Gray
                Write-Host "  TS Trigger Attempted: $($deferralData.TSTriggerAttempted)" -ForegroundColor Gray
                Write-Host "  TS Trigger Success: $($deferralData.TSTriggerSuccess)" -ForegroundColor Gray

                # Copy deferral log
                Copy-DeferralLog -ComputerName $computerName -SourcePath $deferralLogPath -DestinationDirectory $logsDirectory
            }
        }

        # Build device data object (explicitly cast booleans to ensure proper JSON serialization)
        $deviceData = [PSCustomObject]@{
            DeviceName = $computerName
            ResourceID = $resourceID
            PrimaryUser = $primaryUser
            IsOnline = [bool]$isOnline
            OnlineStatus = if ($isOnline) { "Online" } else { "Offline" }
            HostnameVerified = [bool]$hostnameVerified
            ActualHostname = $actualHostname
            VerificationError = $verificationError
            TSStatus = $tsStatus
            OSVersion = $osVersion
            IsWindows11 = [bool]$isWindows11
            DeferralCount = if ($deferralData.LogAvailable) { $deferralData.DeferralCount } else { "N/A" }
            TSTriggerAttempted = if ($deferralData.LogAvailable) { $deferralData.TSTriggerAttempted } else { "N/A" }
            TSTriggerSuccess = if ($deferralData.LogAvailable) { $deferralData.TSTriggerSuccess } else { "N/A" }
            LogAvailable = [bool]$deferralData.LogAvailable
            LastDeferralDate = if ($deferralData.LogAvailable) { $deferralData.LastDeferralDate } else { "N/A" }
            LastTriggerDate = if ($deferralData.LogAvailable) { $deferralData.LastTriggerDate } else { "N/A" }
            DeferralLogPath = "$computerName`_TaskSequenceDeferral.log"
            TSStatusMessagesPath = if ($tsMessagesDownloaded) { "$computerName`_TSStatusMessages.json" } else { "" }
            TSStatusMessagesAvailable = [bool]$tsMessagesDownloaded
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
        CollectionName = $collectionInfo.Name
        TaskSequenceID = $taskSequenceID
        TaskSequenceName = $tsInfo.Name
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
