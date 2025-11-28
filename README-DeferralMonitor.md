# SCCM Task Sequence Deferral Monitor

A web-based monitoring tool for tracking SCCM Task Sequence deferrals across a collection of devices. This tool provides real-time visibility into deferral status, Task Sequence progress, and device health.

## Features

- **Real-time Monitoring**: Track Task Sequence deferral status across your device collection
- **Device Status**: Monitor online/offline status, primary users, and OS versions
- **Windows 11 Override**: Optionally mark Windows 11 devices as "Success" automatically
- **Log Collection**: Automatic collection and storage of deferral logs and TS logs from client machines
- **Web Interface**: Clean, responsive HTML interface with sorting and filtering capabilities
- **Automated Updates**: Runs on a schedule (default: hourly) via Windows Scheduled Task
- **Log Management**: Automatic cleanup of logs from removed devices

## Components

### 1. Configuration File
**SCCMDeferralMonitorConfig.xml** - XML configuration file containing:
- SCCM connection details (Site Code, Site Server, Collection ID, Task Sequence ID)
- Log paths and retention settings
- Windows 11 override setting
- Web server configuration

### 2. Backend Script
**SCCMDeferralMonitor.ps1** - PowerShell script that:
- Connects to SCCM
- Retrieves collection member information
- Collects deferral logs from client machines via UNC paths
- Collects Task Sequence logs from clients
- Generates JSON data files for the web frontend
- Cleans up orphaned log files

### 3. Scheduled Task Setup
**Setup-DeferralMonitorScheduledTask.ps1** - Helper script to:
- Create a Windows Scheduled Task to run the backend script
- Configure hourly execution (configurable interval)
- Run as SYSTEM account with appropriate permissions

### 4. Web Frontend
**SCCMDeferralMonitor.html** - Static HTML/CSS/JavaScript page that:
- Loads device data from JSON files
- Displays devices in a sortable, filterable table
- Provides download links for logs
- Requires no server-side processing (reads from static files only)

## Installation

### Prerequisites
- Windows Server with SCCM Console installed
- ConfigurationManager PowerShell module
- SCCM Admin rights
- Web server (IIS or simple HTTP server) to host the HTML page
- Network access to client machines via UNC paths

### Setup Steps

1. **Extract Files**
   - Copy all files to a directory on your SCCM server (e.g., `C:\SCCMDeferralMonitor\`)

2. **Configure Settings**
   - Edit `SCCMDeferralMonitorConfig.xml`
   - Set your SCCM connection details:
     - `SiteCode`: Your SCCM site code (e.g., "ABC")
     - `SiteServer`: Your SCCM server FQDN (e.g., "sccm01.contoso.com")
     - `CollectionID`: Target collection ID (e.g., "ABC00123")
     - `TaskSequenceID`: Task Sequence package ID to monitor (e.g., "ABC00456")
   - Configure log paths:
     - `DeferralLogPath`: Path to deferral log on clients (default: `C:\Windows\ccm\logs\TaskSequenceDeferral.log`)
     - `LocalLogStoragePath`: Where to store copied logs (e.g., `C:\SCCMDeferralMonitor\Logs`)
     - `TSLogRetentionDays`: How many days of TS logs to collect (default: 7)
   - Set monitoring options:
     - `RunIntervalMinutes`: How often to run (default: 60)
     - `Windows11OverrideSuccess`: Set to `true` or `false`
   - Configure web server:
     - `DataDirectory`: Where to store JSON files for web frontend (e.g., `C:\SCCMDeferralMonitor\WebData`)

3. **Create Scheduled Task**
   - Run PowerShell as Administrator
   - Execute: `.\Setup-DeferralMonitorScheduledTask.ps1`
   - This will create a scheduled task named "SCCM Deferral Monitor" that runs every hour
   - Optionally run the task immediately when prompted

4. **Configure Web Server**
   - Option A: IIS
     - Create a new website or virtual directory in IIS
     - Point the physical path to your installation directory
     - Ensure the website can serve static HTML files
     - Configure MIME types for .json files if needed
   - Option B: Simple HTTP Server (for testing)
     - Run: `python -m http.server 8080` in the installation directory
     - Access via `http://localhost:8080/SCCMDeferralMonitor.html`

5. **Update HTML Configuration**
   - Edit `SCCMDeferralMonitor.html`
   - Update the `DATA_URL`, `METADATA_URL`, and `LOG_BASE_PATH` constants to match your web server setup
   - Ensure the paths correctly reference the JSON files and log directory

6. **Test**
   - Run the scheduled task manually: `Start-ScheduledTask -TaskName "SCCM Deferral Monitor"`
   - Verify JSON files are created in the `DataDirectory`
   - Verify logs are copied to `LocalLogStoragePath`
   - Open the HTML page in a browser and verify data is displayed

## Configuration Details

### XML Configuration Options

```xml
<SCCM>
    <SiteCode>ABC</SiteCode>                      <!-- Your SCCM site code -->
    <SiteServer>sccm01.contoso.com</SiteServer>   <!-- SCCM server FQDN -->
    <CollectionID>ABC00123</CollectionID>         <!-- Target collection -->
    <TaskSequenceID>ABC00456</TaskSequenceID>     <!-- TS package ID -->
</SCCM>

<Logs>
    <DeferralLogPath>C:\Windows\ccm\logs\TaskSequenceDeferral.log</DeferralLogPath>
    <LocalLogStoragePath>C:\SCCMDeferralMonitor\Logs</LocalLogStoragePath>
    <TSLogRetentionDays>7</TSLogRetentionDays>
</Logs>

<Monitoring>
    <RunIntervalMinutes>60</RunIntervalMinutes>
    <Windows11OverrideSuccess>true</Windows11OverrideSuccess>  <!-- Override TS status for Win11 -->
</Monitoring>

<WebServer>
    <Port>8080</Port>
    <DataDirectory>C:\SCCMDeferralMonitor\WebData</DataDirectory>
</WebServer>
```

### Windows 11 Override

When `Windows11OverrideSuccess` is set to `true`:
- Devices detected as running Windows 11 will have their TS Status automatically set to "Success"
- This is useful if Windows 11 devices should be considered compliant regardless of actual TS status
- The override is applied during data collection and reflected in the web UI

### Log Collection

The backend script collects two types of logs:

1. **TaskSequenceDeferral.log**
   - Copied from each online device at every run
   - Stored as `<ComputerName>_TaskSequenceDeferral.log`
   - Only updated if a newer version is available on the client

2. **Task Sequence Logs**
   - Collected from `C:\Windows\CCM\Logs\SMSTSLog` on clients
   - Only logs from the last N days (configurable) are collected
   - Stored in `<ComputerName>_TSLogs\` directory

### Orphaned Log Cleanup

When a device is removed from the collection:
- Its deferral log file is automatically deleted
- Its TS logs directory is automatically deleted
- This prevents accumulation of stale log data

## Web Interface Features

### Display Columns
- **Device Name**: Computer name
- **Primary User**: Primary user from SCCM
- **Status**: Online/Offline indicator with traffic light
- **TS Status**: Task Sequence status (Success, Failed, In Progress, Not Started)
- **Deferral Count**: Number of deferrals from log (or "N/A" if log not available)
- **TS Trigger Attempted**: Whether TS trigger was attempted (Yes/No/N/A)
- **TS Trigger Success**: Whether TS trigger succeeded (Yes/No/N/A)
- **Actions**: Buttons to view TS logs and deferral log

### Filtering Options
- **Search**: Filter by device name or primary user
- **Status Filter**: Show all, online only, or offline only
- **TS Status Filter**: Show all, or filter by TS status

### Sorting
- Click any column header to sort by that column
- Click again to reverse sort order
- Visual indicator shows current sort column and direction

### Log Access
- **TS Logs button**: Opens the directory containing TS logs for that device
- **Deferral Log button**: Opens the deferral log file in a new window
- Buttons are disabled if the device is offline or logs are not available

## Data Flow

1. **Scheduled Task** runs `SCCMDeferralMonitor.ps1` every hour
2. **Backend Script**:
   - Connects to SCCM
   - Retrieves collection members
   - Pings each device to check online status
   - Collects device information (primary user, OS version, TS status)
   - Reads deferral log via UNC path (if online)
   - Copies deferral log to local storage
   - Copies recent TS logs to local storage
   - Generates `devicedata.json` and `metadata.json`
3. **Web Frontend**:
   - Loads JSON files via HTTP
   - Displays data in table
   - Provides filtering and sorting
   - Links to log files

## Troubleshooting

### Script Errors

**"Failed to connect to SCCM"**
- Verify ConfigurationManager module is installed
- Check site code and site server in config
- Ensure you have SCCM admin rights

**"No members found in collection"**
- Verify collection ID is correct
- Check that collection has members
- Ensure SCCM permissions allow viewing collection

**"Log file not found"**
- Verify deferral log path in config matches client configuration
- Check network connectivity to client machines
- Verify admin shares (C$) are accessible

### Web Interface Issues

**"Error loading data"**
- Verify JSON files exist in DataDirectory
- Check web server configuration
- Verify file paths in HTML are correct
- Check browser console for errors

**"No devices match the current filters"**
- Clear all filters and search
- Verify data was collected successfully
- Check scheduled task execution history

**Log files won't open**
- Verify web server is configured to serve the Logs directory
- Check LOG_BASE_PATH in HTML matches your web server setup
- Ensure file permissions allow web server to access logs

### Scheduled Task Issues

**Task not running**
- Check Task Scheduler for errors
- Verify SYSTEM account has necessary permissions
- Check script path in task action
- Review task history for failure codes

**Task runs but no data appears**
- Check script output in a manual run
- Verify all paths in config file exist
- Check SCCM connection details
- Review event logs for errors

## Security Considerations

- The scheduled task runs as SYSTEM for SCCM access
- Ensure web server has appropriate authentication configured
- Consider restricting web interface to internal network only
- Log files may contain sensitive device information
- Use HTTPS for production deployments

## Maintenance

### Regular Tasks
- Monitor disk space in log storage directory
- Review scheduled task execution history
- Verify data is being updated regularly
- Check for orphaned log files

### Adjusting Collection Interval
To change from hourly to a different interval:
1. Edit `SCCMDeferralMonitorConfig.xml` and update `RunIntervalMinutes`
2. Re-run `Setup-DeferralMonitorScheduledTask.ps1` with `-IntervalMinutes` parameter
   ```powershell
   .\Setup-DeferralMonitorScheduledTask.ps1 -IntervalMinutes 30
   ```

### Changing Collections or Task Sequences
1. Update `SCCMDeferralMonitorConfig.xml` with new IDs
2. Run the backend script manually to test
3. Scheduled task will use new config on next run

## File Structure

```
C:\SCCMDeferralMonitor\
├── SCCMDeferralMonitor.ps1                    # Backend data collection script
├── SCCMDeferralMonitorConfig.xml              # Configuration file
├── Setup-DeferralMonitorScheduledTask.ps1     # Scheduled task setup script
├── SCCMDeferralMonitor.html                   # Web frontend
├── README-DeferralMonitor.md                  # This file
├── Logs\                                      # Log storage (created automatically)
│   ├── COMPUTER1_TaskSequenceDeferral.log
│   ├── COMPUTER1_TSLogs\
│   │   ├── smsts.log
│   │   └── ...
│   └── ...
└── WebData\                                   # JSON data files (created automatically)
    ├── devicedata.json
    └── metadata.json
```

## Version History

- **v1.0** (2025-11-28): Initial release

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Review script output and logs
3. Verify configuration settings
4. Contact your IT department or SCCM administrator

## License

Internal use only. Modify as needed for your environment.
