<#
.SYNOPSIS
    Setup script for SCCM Deferral Monitor Scheduled Task

.DESCRIPTION
    Creates a Windows Scheduled Task to run the SCCMDeferralMonitor.ps1 script
    every hour. Requires administrative privileges.

.NOTES
    Date: 2025-11-28
    Requires: Administrator rights
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$ScriptPath = "",

    [Parameter(Mandatory=$false)]
    [string]$ConfigPath = "",

    [Parameter(Mandatory=$false)]
    [int]$IntervalMinutes = 60,

    [Parameter(Mandatory=$false)]
    [string]$TaskName = "SCCM Deferral Monitor"
)

# Check if running as administrator
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "ERROR: This script must be run as Administrator" -ForegroundColor Red
    Write-Host "Please run PowerShell as Administrator and try again" -ForegroundColor Yellow
    exit 1
}

try {
    # Resolve script directory
    $scriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Path

    # Resolve paths
    if ([string]::IsNullOrEmpty($ScriptPath)) {
        $ScriptPath = Join-Path $scriptDirectory "SCCMDeferralMonitor.ps1"
    }

    if ([string]::IsNullOrEmpty($ConfigPath)) {
        $ConfigPath = Join-Path $scriptDirectory "SCCMDeferralMonitorConfig.xml"
    }

    # Validate files exist
    if (-not (Test-Path $ScriptPath)) {
        throw "Script not found: $ScriptPath"
    }

    if (-not (Test-Path $ConfigPath)) {
        throw "Config file not found: $ConfigPath"
    }

    Write-Host "`nSCCM Deferral Monitor - Scheduled Task Setup" -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    Write-Host "Script Path: $ScriptPath" -ForegroundColor Gray
    Write-Host "Config Path: $ConfigPath" -ForegroundColor Gray
    Write-Host "Interval: Every $IntervalMinutes minutes" -ForegroundColor Gray
    Write-Host "Task Name: $TaskName`n" -ForegroundColor Gray

    # Check if task already exists
    $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

    if ($existingTask) {
        Write-Host "Task already exists. Removing old task..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    }

    # Create task action
    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`" -ConfigFile `"$ConfigPath`""

    # Create task trigger (repeating every X minutes)
    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Minutes $IntervalMinutes) -RepetitionDuration ([TimeSpan]::MaxValue)

    # Create task settings
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -RunOnlyIfNetworkAvailable

    # Create task principal (run as SYSTEM)
    $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

    # Register the task
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Collects SCCM deferral data and generates web reports every $IntervalMinutes minutes"

    Write-Host "`nScheduled task created successfully!" -ForegroundColor Green
    Write-Host "Task will run every $IntervalMinutes minutes" -ForegroundColor Green
    Write-Host "`nTo verify the task, run: Get-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Cyan
    Write-Host "To run the task manually now, run: Start-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Cyan
    Write-Host "To remove the task, run: Unregister-ScheduledTask -TaskName '$TaskName'" -ForegroundColor Cyan

    # Ask if user wants to run it now
    $runNow = Read-Host "`nWould you like to run the task now? (Y/N)"
    if ($runNow -eq 'Y' -or $runNow -eq 'y') {
        Write-Host "Starting task..." -ForegroundColor Cyan
        Start-ScheduledTask -TaskName $TaskName
        Write-Host "Task started!" -ForegroundColor Green
    }

    exit 0
}
catch {
    Write-Host "`nERROR: $_" -ForegroundColor Red
    exit 1
}
