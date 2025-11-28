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
        # Query WMI directly for per-device status using SMS_ClientAdvertisementStatus
        # This has actual device-level status, unlike Get-CMDeploymentStatus which returns summary data
        $query = @"
SELECT sys.Name0, adv.AdvertisementID, adv.LastStateName, adv.LastStatusMessageID
FROM SMS_ClientAdvertisementStatus adv
JOIN SMS_R_System sys ON adv.ResourceID = sys.ResourceId
WHERE adv.ResourceID = '$ResourceID' AND adv.AdvertisementID LIKE '%$TaskSequenceID%'
"@

        Write-Host "  [DEBUG] Querying SMS_ClientAdvertisementStatus for device-level TS status" -ForegroundColor Gray

        $status = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" `
            -ComputerName $script:sccmSiteServer `
            -Query $query `
            -ErrorAction SilentlyContinue

        if ($status) {
            Write-Host "  [DEBUG] Found advertisement status: $($status.LastStateName)" -ForegroundColor Green

            # Map common state names to our status values
            switch -Wildcard ($status.LastStateName) {
                "*Success*" { return "Success" }
                "*Running*" { return "In Progress" }
                "*Progress*" { return "In Progress" }
                "*Failed*" { return "Failed" }
                "*Error*" { return "Failed" }
                "*Waiting*" { return "Not Started" }
                default { return $status.LastStateName }
            }
        }
        else {
            Write-Host "  [DEBUG] No advertisement status found - checking if any execution history exists" -ForegroundColor Gray

            # Check if there's any execution history at all
            $execQuery = "SELECT TOP 1 * FROM SMS_TaskExecutionStatus WHERE ResourceID = '$ResourceID' AND PackageID = '$TaskSequenceID'"
            $execStatus = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" `
                -ComputerName $script:sccmSiteServer `
                -Query $execQuery `
                -ErrorAction SilentlyContinue

            if ($execStatus) {
                return "In Progress"
            }

            return "Not Started"
        }
    }
    catch {
        Write-Host "  [ERROR] Exception in Get-TaskSequenceDeploymentStatusFromSCCM: $_" -ForegroundColor Red
        Write-Host "  [ERROR] Stack Trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        return "Unknown"
    }
}

function Get-AllTaskSequenceDataFromSCCM {
    <#
    .SYNOPSIS
        Downloads ALL TS data (execution messages AND status) for entire collection at once using ConfigMgr cmdlets
        Uses Get-CMDeploymentStatusDetails to retrieve everything in one go
        Returns hashtable with both Messages and Status for each device
    #>
    param(
        [array]$CollectionMembers,
        [string]$TaskSequenceID,
        [string]$CollectionID,
        [string]$DestinationDirectory,
        [int]$DaysBack = 0
    )

    try {
        Write-Host "`nCollecting ALL TS data using ConfigMgr cmdlets..." -ForegroundColor Cyan
        Write-Host "[INFO] Collection $CollectionID is ONLY for getting list of machines to monitor" -ForegroundColor Cyan
        Write-Host "[INFO] Getting ALL deployments for TS $TaskSequenceID (from ANY collection)" -ForegroundColor Cyan

        # Get ALL deployments for this task sequence (NOT filtered by collection!)
        # The CollectionID parameter is ONLY used to get the list of machines to monitor
        # The TS could be deployed to ANY collection(s), not necessarily the monitoring collection
        $currentLocation = Get-Location
        Set-Location "$($script:sccmSiteCode):" -ErrorAction Stop

        $deployments = Get-CMTaskSequenceDeployment -TaskSequenceId $TaskSequenceID -ErrorAction SilentlyContinue

        Set-Location $currentLocation

        if (-not $deployments) {
            Write-Host "[WARNING] No deployments found for TaskSequence $TaskSequenceID" -ForegroundColor Yellow
            return @{}
        }

        Write-Host "[SUCCESS] Found $(@($deployments).Count) deployment(s) for this TS across ALL collections" -ForegroundColor Green

        # Collect all status details from all deployments using ConfigMgr cmdlets
        $allStatusDetails = @()

        foreach ($deployment in $deployments) {
            Write-Host "[DEBUG] Processing deployment: $($deployment.AdvertisementID)" -ForegroundColor Gray

            try {
                Set-Location "$($script:sccmSiteCode):" -ErrorAction Stop

                # Use ConfigMgr cmdlet to get deployment status
                $deploymentStatus = Get-CMDeploymentStatus -DeploymentId $deployment.AdvertisementID -ErrorAction SilentlyContinue

                if ($deploymentStatus) {
                    # Use ConfigMgr cmdlet to get detailed status (this has EVERYTHING)
                    $statusDetails = $deploymentStatus | Get-CMDeploymentStatusDetails -ErrorAction SilentlyContinue

                    if ($statusDetails) {
                        # Filter by date if needed - use SummarizationTime property
                        if ($DaysBack -gt 0) {
                            $startDate = (Get-Date).AddDays(-$DaysBack)
                            $statusDetails = $statusDetails | Where-Object {
                                $_.SummarizationTime -and [datetime]$_.SummarizationTime -ge $startDate
                            }
                            Write-Host "[DEBUG] Date filter: Last $DaysBack days (since $($startDate.ToString('yyyy-MM-dd')))" -ForegroundColor Gray
                        }

                        $allStatusDetails += $statusDetails
                        Write-Host "[DEBUG] Retrieved $(@($statusDetails).Count) status detail records" -ForegroundColor Gray
                    }
                }

                Set-Location $currentLocation
            }
            catch {
                Write-Host "[ERROR] Failed to get status details for deployment $($deployment.AdvertisementID): $_" -ForegroundColor Red
                Set-Location $currentLocation
            }
        }

        if (-not $allStatusDetails -or $allStatusDetails.Count -eq 0) {
            Write-Host "[WARNING] No TS data found for any deployments" -ForegroundColor Yellow
            return @{}
        }

        Write-Host "[SUCCESS] Found $($allStatusDetails.Count) total status detail records across ALL deployments" -ForegroundColor Green

        # Create a set of machine names from our monitoring collection for filtering
        $monitoringMachines = @{}
        foreach ($member in $CollectionMembers) {
            $monitoringMachines[$member.Name] = $true
        }
        Write-Host "[DEBUG] Monitoring collection has $($monitoringMachines.Count) machines" -ForegroundColor Gray

        # Filter to only devices in our monitoring collection, then group by DeviceName
        $filteredDetails = $allStatusDetails | Where-Object { $monitoringMachines.ContainsKey($_.DeviceName) }
        Write-Host "[DEBUG] Filtered to $($filteredDetails.Count) records for machines in monitoring collection" -ForegroundColor Green

        $messagesByDevice = $filteredDetails | Group-Object -Property DeviceName
        Write-Host "[DEBUG] Data grouped into $($messagesByDevice.Count) device(s)" -ForegroundColor Gray

        # Create hashtable: DeviceName -> { Messages: [...], Status: "..." }
        $resultHash = @{}
        foreach ($group in $messagesByDevice) {
            if ($group.Name) {
                # Get the most recent status for this device - use SummarizationTime property
                $latestRecord = $group.Group | Sort-Object SummarizationTime -Descending | Select-Object -First 1

                # Determine status from StatusDescription
                $tsStatus = "Unknown"
                if ($latestRecord.StatusDescription) {
                    # Map status description to our status values
                    switch -Wildcard ($latestRecord.StatusDescription) {
                        "*Success*" { $tsStatus = "Success" }
                        "*Complete*" { $tsStatus = "Success" }
                        "*finished successfully*" { $tsStatus = "Success" }
                        "*Running*" { $tsStatus = "In Progress" }
                        "*Progress*" { $tsStatus = "In Progress" }
                        "*reboot*" { $tsStatus = "In Progress" }
                        "*Failed*" { $tsStatus = "Failed" }
                        "*Error*" { $tsStatus = "Failed" }
                        default { $tsStatus = "In Progress" }
                    }
                }

                $resultHash[$group.Name] = @{
                    Messages = $group.Group
                    Status = $tsStatus
                }
            }
        }

        Write-Host "[SUCCESS] Processed $($resultHash.Count) devices with TS data from cmdlets" -ForegroundColor Green

        # Now get detailed execution messages via WMI (SMS_TaskSequenceExecutionStatus) - collection-wide
        Write-Host "`n[INFO] Querying SMS_TaskSequenceExecutionStatus WMI class for detailed step execution data..." -ForegroundColor Cyan
        Write-Host "[INFO] This class has: ExecutionTime, Step, ActionName, GroupName, ExitCode, ActionOutput" -ForegroundColor Cyan

        # Build ResourceID list for WMI query
        $resourceIDs = ($CollectionMembers | ForEach-Object { $_.ResourceID }) -join "','"

        # Query SMS_TaskSequenceExecutionStatus for TS execution messages
        # This is the correct class that has ActionOutput!
        $wmiQuery = "SELECT * FROM SMS_TaskSequenceExecutionStatus WHERE ResourceID IN ('$resourceIDs') AND PackageID = '$TaskSequenceID'"

        if ($DaysBack -gt 0) {
            $startDate = (Get-Date).AddDays(-$DaysBack)
            $wmiDate = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($startDate)
            $wmiQuery += " AND ExecutionTime >= '$wmiDate'"
        }

        Write-Host "[DEBUG] WMI Query: $wmiQuery" -ForegroundColor Gray

        try {
            $executionMessages = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" `
                -ComputerName $script:sccmSiteServer `
                -Query $wmiQuery `
                -ErrorAction SilentlyContinue

            if ($executionMessages) {
                Write-Host "[SUCCESS] Retrieved $(@($executionMessages).Count) detailed execution messages from WMI" -ForegroundColor Green

                # Create ResourceID to ComputerName mapping
                $resourceIDtoName = @{}
                foreach ($member in $CollectionMembers) {
                    $resourceIDtoName[$member.ResourceID] = $member.Name
                }

                # Group by ResourceID and add to existing results
                $messagesByResource = $executionMessages | Group-Object -Property ResourceID

                foreach ($group in $messagesByResource) {
                    $resourceID = $group.Name
                    $computerName = $resourceIDtoName[$resourceID]

                    if ($computerName) {
                        if ($resultHash.ContainsKey($computerName)) {
                            # Add detailed messages to existing entry
                            $resultHash[$computerName].DetailedMessages = $group.Group
                        }
                        else {
                            # Create new entry
                            $resultHash[$computerName] = @{
                                Messages = @()
                                Status = "Unknown"
                                DetailedMessages = $group.Group
                            }
                        }
                    }
                }

                Write-Host "[SUCCESS] Added detailed execution messages to $($messagesByResource.Count) device(s)" -ForegroundColor Green
            }
            else {
                Write-Host "[WARNING] No detailed execution messages found in SMS_TaskSequenceExecutionStatus" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host "[ERROR] Failed to query SMS_TaskSequenceExecutionStatus: $_" -ForegroundColor Red
            Write-Host "[ERROR] Query was: $wmiQuery" -ForegroundColor Red
        }

        Write-Host "[SUCCESS] Processed $($resultHash.Count) total devices with TS data" -ForegroundColor Green
        return $resultHash
    }
    catch {
        Write-Host "[ERROR] Failed to collect TS data: $_" -ForegroundColor Red
        Write-Host "[ERROR] Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        Set-Location $currentLocation -ErrorAction SilentlyContinue
        return @{}
    }
}

function Save-TaskSequenceExecutionLog {
    <#
    .SYNOPSIS
        Saves TS execution messages for a single computer to JSON and HTML files
    #>
    param(
        [string]$ComputerName,
        [string]$ResourceID,
        [string]$TaskSequenceID,
        [array]$Messages,
        [string]$DestinationDirectory
    )

    try {
        if (-not $Messages -or $Messages.Count -eq 0) {
            Write-Host "  No TS execution messages for $ComputerName" -ForegroundColor Gray
            return $false
        }

        Write-Host "  Processing $($Messages.Count) TS execution message(s) for $ComputerName" -ForegroundColor Cyan

        # Convert to structured format with full details
        $executionSteps = @()

        foreach ($msg in $Messages) {
                # SMS_TaskSequenceExecutionStatus WMI class has properties WITHOUT spaces
                # ExecutionTime is WMI datetime format that needs conversion

                # Format execution time from WMI datetime
                $executionTime = ""
                if ($msg.ExecutionTime) {
                    try {
                        $executionTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($msg.ExecutionTime).ToString("yyyy-MM-dd HH:mm:ss")
                    }
                    catch {
                        $executionTime = $msg.ExecutionTime
                    }
                }

                $executionSteps += [PSCustomObject]@{
                    ComputerName = $ComputerName
                    ExecutionTime = $executionTime
                    Step = if ($null -ne $msg.Step) { $msg.Step } else { 0 }
                    ActionName = if ($msg.ActionName) { $msg.ActionName } else { "" }
                    GroupName = if ($msg.GroupName) { $msg.GroupName } else { "" }
                    LastStatusMsgName = if ($msg.LastStatusMsgName) { $msg.LastStatusMsgName } else { "" }
                    LastMessageID = if ($msg.LastStatusMsgID) { $msg.LastStatusMsgID } else { "" }
                    ExitCode = if ($null -ne $msg.ExitCode) { $msg.ExitCode } else { "" }
                    ActionOutput = if ($msg.ActionOutput) { $msg.ActionOutput } else { "" }  # Full output, NO truncation!
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

function Get-DeviceOSInfo {
    param(
        [string]$ComputerName,
        [string]$ResourceID
    )

    try {
        # Query SMS_R_System for detailed OS information including build number
        $query = "SELECT Build01, Caption0, Version0 FROM SMS_R_System WHERE ResourceID = '$ResourceID'"

        $system = Get-WmiObject -Namespace "ROOT\SMS\site_$($script:sccmSiteCode)" `
            -ComputerName $script:sccmSiteServer `
            -Query $query `
            -ErrorAction SilentlyContinue

        if ($system) {
            $build = $system.Build01
            $version = $system.Version0

            # Determine OS name ONLY from build number (SCCM Caption0 is unreliable)
            $osName = "Unknown"
            $isWindows11 = $false
            $buildNumber = ""

            if ($build -and $build -match '10\.0\.(\d+)') {
                $buildNumber = $matches[1]
                $buildNum = [int]$buildNumber

                # Determine OS from build number only
                if ($buildNum -ge 22000) {
                    # Windows 11 version mapping
                    if ($buildNum -ge 26100) {
                        $osName = "Windows 11 24H2"
                    }
                    elseif ($buildNum -ge 22631) {
                        $osName = "Windows 11 23H2"
                    }
                    elseif ($buildNum -ge 22621) {
                        $osName = "Windows 11 22H2"
                    }
                    elseif ($buildNum -ge 22000) {
                        $osName = "Windows 11 21H2"
                    }
                    $isWindows11 = $true
                }
                elseif ($buildNum -ge 10240) {
                    # Windows 10 version mapping
                    if ($buildNum -ge 19045) {
                        $osName = "Windows 10 22H2"
                    }
                    elseif ($buildNum -ge 19044) {
                        $osName = "Windows 10 21H2"
                    }
                    elseif ($buildNum -ge 19043) {
                        $osName = "Windows 10 21H1"
                    }
                    elseif ($buildNum -ge 19042) {
                        $osName = "Windows 10 20H2"
                    }
                    elseif ($buildNum -ge 19041) {
                        $osName = "Windows 10 2004"
                    }
                    elseif ($buildNum -ge 18363) {
                        $osName = "Windows 10 1909"
                    }
                    elseif ($buildNum -ge 18362) {
                        $osName = "Windows 10 1903"
                    }
                    elseif ($buildNum -ge 17763) {
                        $osName = "Windows 10 1809"
                    }
                    elseif ($buildNum -ge 17134) {
                        $osName = "Windows 10 1803"
                    }
                    elseif ($buildNum -ge 16299) {
                        $osName = "Windows 10 1709"
                    }
                    elseif ($buildNum -ge 15063) {
                        $osName = "Windows 10 1703"
                    }
                    elseif ($buildNum -ge 14393) {
                        $osName = "Windows 10 1607"
                    }
                    elseif ($buildNum -ge 10586) {
                        $osName = "Windows 10 1511"
                    }
                    else {
                        $osName = "Windows 10 1507"
                    }
                }
                else {
                    # Older Windows versions
                    $osName = "Windows (Build $buildNumber)"
                }
            }

            return @{
                OSName = $osName
                Build = $build
                BuildNumber = $buildNumber
                Version = $version
                IsWindows11 = $isWindows11
            }
        }

        return @{
            OSName = "Unknown"
            Build = ""
            BuildNumber = ""
            Version = ""
            IsWindows11 = $false
        }
    }
    catch {
        Write-Host "  [ERROR] Failed to get OS info: $_" -ForegroundColor Red
        return @{
            OSName = "Unknown"
            Build = ""
            BuildNumber = ""
            Version = ""
            IsWindows11 = $false
        }
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
    $tsStatusMessagesDaysBack = [int]$config.Configuration.Settings.Monitoring.TSStatusMessagesDaysBack

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

    # Collect ALL TS data (messages AND status) for the collection at once - NO PER-DEVICE QUERIES!
    # Uses ONLY ConfigMgr cmdlets: Get-CMTaskSequenceDeployment, Get-CMDeploymentStatus, Get-CMDeploymentStatusDetails
    Write-Host "`n============================================" -ForegroundColor Cyan
    Write-Host "COLLECTION-WIDE DATA PULL (ConfigMgr Cmdlets Only)" -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    $allTSData = Get-AllTaskSequenceDataFromSCCM -CollectionMembers $members -TaskSequenceID $taskSequenceID -CollectionID $collectionID -DestinationDirectory $logsDirectory -DaysBack $tsStatusMessagesDaysBack

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

        $osInfo = Get-DeviceOSInfo -ComputerName $computerName -ResourceID $resourceID
        $isWindows11 = $osInfo.IsWindows11
        Write-Host "  OS: $($osInfo.OSName) (Build: $($osInfo.BuildNumber))" -ForegroundColor Gray

        # Get TS status from the collection-wide data we already pulled - NO MORE WMI QUERIES!
        $tsStatus = "Unknown"
        $tsMessagesDownloaded = $false

        if ($allTSData.ContainsKey($computerName)) {
            # Extract status from the collection-wide pull
            $tsStatus = $allTSData[$computerName].Status
            Write-Host "  TS Status: $tsStatus (from collection-wide data)" -ForegroundColor Gray

            # Apply Windows 11 override
            if ($win11Override -and $isWindows11 -and $tsStatus -ne "Success") {
                Write-Host "  TS Status: $tsStatus -> Success (Win11 Override)" -ForegroundColor Green
                $tsStatus = "Success"
            }

            # Save TS execution messages from the collection-wide data
            # Prefer DetailedMessages (from WMI) if available, otherwise use Messages (from cmdlet)
            $messagesToSave = $null
            if ($allTSData[$computerName].DetailedMessages) {
                $messagesToSave = $allTSData[$computerName].DetailedMessages
                Write-Host "  Using detailed WMI messages (SMS_StatusMessage)" -ForegroundColor Cyan
            }
            elseif ($allTSData[$computerName].Messages) {
                $messagesToSave = $allTSData[$computerName].Messages
                Write-Host "  Using cmdlet messages (SMS_ClassicDeploymentAssetDetails)" -ForegroundColor Cyan
            }

            if ($messagesToSave) {
                $tsMessagesDownloaded = Save-TaskSequenceExecutionLog -ComputerName $computerName -ResourceID $resourceID -TaskSequenceID $taskSequenceID -Messages $messagesToSave -DestinationDirectory $logsDirectory
            }
        }
        else {
            Write-Host "  No TS data found in collection-wide pull" -ForegroundColor Gray

            # Apply Windows 11 override even if no data
            if ($win11Override -and $isWindows11) {
                Write-Host "  TS Status: Unknown -> Success (Win11 Override)" -ForegroundColor Green
                $tsStatus = "Success"
            }
        }

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
            OSName = $osInfo.OSName
            OSVersion = $osInfo.Build
            OSBuildNumber = $osInfo.BuildNumber
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
