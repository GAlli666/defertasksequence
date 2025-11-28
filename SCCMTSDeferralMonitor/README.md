# SCCM Task Sequence Deferral Monitor

A two-part solution for monitoring SCCM Task Sequence deferrals with backend data collection and WPF frontend viewer.

## Architecture

### Backend (Data Collector)
- **Purpose**: Collects data from SCCM, downloads logs from client machines
- **Location**: `Backend/SCCMTSDeferralCollector.ps1`
- **Runs On**: SCCM server or jumpbox with SCCM access
- **Schedule**: Via Windows Task Scheduler (hourly recommended)
- **Output**: JSON files (metadata + device data) and log files

### Frontend (WPF Viewer)
- **Purpose**: Displays collected data in a modern WPF interface
- **Location**: `Frontend/SCCMTSDeferralViewer.ps1`
- **Runs On**: Any Windows machine with network access to data folder
- **Input**: Reads JSON files from configured data path

## Setup Instructions

### Part 1: Backend Setup

1. **Configure Settings**
   - Edit `Backend/SCCMDeferralMonitorConfig.xml`
   - Set SCCM site code, server, collection ID, and task sequence ID
   - Configure output paths for data and logs

2. **Deploy to SCCM Server/Jumpbox**
   ```powershell
   # Copy files to SCCM server
   Copy-Item Backend\* \\sccmserver\c$\Scripts\SCCMDeferral\
   ```

3. **Setup Scheduled Task**
   ```powershell
   # Run on SCCM server
   cd C:\Scripts\SCCMDeferral
   .\Setup-DeferralMonitorScheduledTask.ps1
   ```
   - Choose schedule: Hourly recommended
   - Task will run as SYSTEM with SCCM permissions

4. **Test Initial Run**
   ```powershell
   .\SCCMTSDeferralCollector.ps1
   ```
   - Verify data files created in configured paths
   - Check logs directory for downloaded files

### Part 2: Frontend Setup

1. **Configure Data Path**
   - Option 1: Edit script variable at top of `SCCMTSDeferralViewer.ps1`
     ```powershell
     $script:DataPath = "\\sccmserver\share\Data"
     $script:LogsPath = "\\sccmserver\share\Logs"
     ```
   - Option 2: Use Settings button in UI after launch

2. **Launch Viewer**
   ```powershell
   .\Frontend\SCCMTSDeferralViewer.ps1
   ```

3. **Usage**
   - **Refresh**: Click refresh button to reload data
   - **Search**: Filter by device name or user
   - **Filters**: Filter by online/offline or TS status
   - **Sort**: Click column headers to sort
   - **View Logs**: Click buttons to open log files

## Configuration Reference

### Backend Config (SCCMDeferralMonitorConfig.xml)

```xml
<Configuration>
  <Settings>
    <SCCM>
      <SiteCode>PR1</SiteCode>
      <SiteServer>sccmserver.domain.com</SiteServer>
      <CollectionID>PR10077A</CollectionID>
      <TaskSequenceID>PR100867</TaskSequenceID>
    </SCCM>
    <WebServer>
      <WebRootPath>C:\SCCMDeferralMonitor</WebRootPath>
      <DataDirectory>Data</DataDirectory>
      <LogsDirectory>Logs</LogsDirectory>
    </WebServer>
    <Monitoring>
      <Windows11OverrideSuccess>true</Windows11OverrideSuccess>
      <HostnameVerificationTimeoutSeconds>3</HostnameVerificationTimeoutSeconds>
      <PingTimeoutMs>1000</PingTimeoutMs>
    </Monitoring>
    <Logs>
      <DeferralLogPath>C:\Windows\ccm\logs\TaskSequenceDeferral.log</DeferralLogPath>
    </Logs>
  </Settings>
</Configuration>
```

### Frontend Config

Edit top of `SCCMTSDeferralViewer.ps1`:
```powershell
$script:DataPath = "C:\SCCMDeferralMonitor\Data"  # Path to JSON files
$script:LogsPath = "C:\SCCMDeferralMonitor\Logs"  # Path to log files
```

## Features

### Backend Features
- ✅ SCCM data collection (OS version, primary user, collection membership)
- ✅ TS execution status from `vSMS_TaskSequenceExecutionStatus`
- ✅ Direct ping and WMI hostname verification
- ✅ C$ share access for log downloads
- ✅ TaskSequenceDeferral.log collection
- ✅ Windows 11 success override
- ✅ Orphaned log cleanup
- ✅ JSON export for viewer consumption

### Frontend Features
- ✅ Modern WPF interface with dark theme
- ✅ Real-time data refresh
- ✅ Sortable columns
- ✅ Multi-filter support (search, status, TS status)
- ✅ One-click log file viewing
- ✅ Configurable data path
- ✅ Responsive grid layout

## File Outputs

### Data Directory
- `metadata.json` - Collection info, counts, last update
- `devicedata.json` - Array of device objects with all properties

### Logs Directory
- `{ComputerName}_TaskSequenceDeferral.log` - Deferral logs from clients
- `{ComputerName}_TSStatusMessages.json` - TS execution history from SCCM

## Requirements

### Backend
- Windows Server with SCCM Console installed
- PowerShell 5.1+
- ConfigurationManager PowerShell module
- SCCM admin rights
- Network access to client machines (C$ shares)

### Frontend
- Windows client or server
- PowerShell 5.1+
- .NET Framework 4.5+ (for WPF)
- Network access to data/logs directories

## Troubleshooting

### Backend Issues

**No data collected:**
- Check SCCM connection: `Get-CMSite`
- Verify collection ID exists
- Check permissions for C$ share access

**WMI errors:**
- Ensure Windows Firewall allows WMI
- Check admin rights on client machines
- Verify network connectivity

**Empty logs directory:**
- Check TaskSequenceDeferral.log exists on clients
- Verify C$ share access: `dir \\computername\c$`
- Check configured DeferralLogPath in config

### Frontend Issues

**No data shown:**
- Click Settings and verify Data Path
- Check network access to data directory
- Verify JSON files exist and are readable

**Can't open log files:**
- Check Logs Path in Settings
- Verify network access to logs directory
- Ensure log files were downloaded by backend

## Network Share Setup (Recommended)

For multi-user access, share the output directory:

1. **On SCCM Server/Jumpbox:**
   ```powershell
   New-SmbShare -Name "SCCMDeferral" -Path "C:\SCCMDeferralMonitor" -ReadAccess "Domain Users"
   ```

2. **On Client Machines:**
   ```powershell
   # Configure frontend to use UNC path
   $script:DataPath = "\\sccmserver\SCCMDeferral\Data"
   $script:LogsPath = "\\sccmserver\SCCMDeferral\Logs"
   ```

## License

Internal use only. SCCM admin rights required.
