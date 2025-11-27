# SCCM Task Sequence Deferral Log Reader Tool

## Overview

A diagnostic tool that connects to SCCM, retrieves collection members, and reads `TaskSequenceDeferral.log` files from each member via UNC path. The tool displays deferral status, trigger attempts, and success/failure information in a modern WPF UI with traffic light indicators.

## Features

- **SCCM Integration**: Connects to SCCM using the ConfigurationManager PowerShell module
- **Collection Scanning**: Retrieves all members from a specified Collection ID
- **Remote Log Reading**: Reads log files from remote computers via UNC paths
- **Online/Offline Detection**: Checks if computers are reachable via ping
- **Traffic Light Indicators**: Visual status indicators for quick assessment
  - 🟢 Green: Good/Success (0 deferrals)
  - 🟠 Orange: Warning (1-2 deferrals)
  - 🔴 Red: Critical (3+ deferrals or failed)
  - ⚪ Gray: Not attempted
- **Offline Highlighting**: Offline computers shown with red border and background
- **Detailed Status Display**: Shows:
  - Computer name
  - Online/offline status
  - Overall status (Ready, Deferred, Success, Failed, Offline, No Log File)
  - Deferral count with traffic light
  - Deferral status description
  - Task Sequence trigger attempted (Yes/No)
  - Task Sequence trigger success (Yes/No)
  - Task Sequence Package ID
  - Last deferral date/time
  - Last TS trigger date/time
  - Error messages (if any)
- **CSV Export**: Export scan results to CSV for reporting
- **Modern UI**: Clean, bright interface with light blue accent (#5DADE2), 1500x900 resolution

## Prerequisites

1. **PowerShell 5.1** or higher
2. **SCCM Admin Rights** to query collections
3. **ConfigurationManager PowerShell Module** (installed with SCCM Console)
4. **Network Access** to target computers (for UNC path access)
5. **Administrative Credentials** (may be required for remote file access)

## Installation

1. Download `SCCMLogReaderTool.ps1` to your administrative workstation
2. Ensure you have the SCCM Console installed (provides ConfigurationManager module)
3. Run PowerShell as Administrator

## Usage

### Step 1: Launch the Tool

```powershell
.\SCCMLogReaderTool.ps1
```

### Step 2: Connect to SCCM

1. Enter your **Site Code** (e.g., "ABC")
2. Enter your **Site Server** FQDN (e.g., "sccm01.contoso.com")
3. Click **Connect**
4. Wait for connection confirmation

### Step 3: Configure Scan Settings

1. Enter the **Collection ID** you want to scan (e.g., "ABC00123")
2. Review/modify the **Log File Path** (default: `C:\Windows\ccm\logs\TaskSequenceDeferral.log`)
   - This is the local path on each client
   - The tool will automatically convert to UNC path (e.g., `\\computer\C$\Windows\ccm\logs\TaskSequenceDeferral.log`)

### Step 4: Scan Collection

1. Click **Scan Collection**
2. The tool will:
   - Retrieve all collection members from SCCM
   - Check each computer's online status (ping)
   - Read log files from online computers
   - Parse log data and extract deferral/trigger information
   - Display results in real-time in the data grid
3. Progress bar shows scan progress

### Step 5: Review Results

The data grid shows all computers with their status:

- **Computer Name**: Name of the computer
- **Online**: 🟢 Green = Online, 🔴 Red = Offline
- **Status**: Overall status (Ready, Deferred, Success, Failed, Offline, No Log File, Error)
- **Deferral Count**: Number with traffic light indicator
  - 🟢 0 deferrals
  - 🟠 1-2 deferrals
  - 🔴 3+ deferrals
- **Deferral Status**: Text description (e.g., "Deferred (2)", "TS Started", "TS Failed")
- **TS Trigger Attempted**: 🟢 Yes / ⚪ No
- **TS Trigger Success**: 🟢 Success / 🔴 Failed
- **TS Package ID**: Task Sequence Package ID (extracted from logs)
- **Last Deferral**: Date/time of last deferral
- **Last TS Trigger**: Date/time of last TS trigger attempt
- **Error**: Any error messages

**Offline computers** are highlighted with:
- Red border (2px)
- Light red background (#FADBD8)

### Step 6: Export Results (Optional)

1. Click **Export to CSV**
2. Choose save location
3. CSV file contains all scan data for reporting/analysis

## Log File Format

The tool parses logs created by `deferTS.ps1` with the format:

```
[2025-11-27 10:13:45] [Info] Message here
[2025-11-27 10:15:22] [Warning] Warning message
[2025-11-27 10:16:33] [Error] Error message
```

### Key Log Patterns Detected

- **Deferral Increment**: `Deferral count incremented immediately: X / Y`
- **Deferral Reset**: `Deferral count reset to 0`
- **User Deferred**: `User chose to defer`
- **TS Trigger Attempt**: `Attempting to start Task Sequence: ABC00123`
- **TS Trigger Success**: `Task Sequence triggered successfully!`
- **TS Started**: `Task Sequence started successfully`
- **TS Trigger Failed**: `Failed to start Task Sequence` or `Failed to trigger schedule`

### Scenario Handling

The tool correctly handles complex scenarios:

1. **Multiple Deferrals**: Tracks cumulative deferral count
2. **Deferral + Trigger Success**: Shows deferral count was reset, TS started successfully
3. **Deferral + Trigger Failed**: Shows trigger was attempted but failed, deferral count remains
4. **Trigger Retry After Failure**: Detects reset after failed trigger, then subsequent successful trigger
5. **Multiple TS Packages**: Extracts and displays the Package ID being triggered

## Troubleshooting

### Connection Issues

**Error**: "Failed to connect to SCCM"

- Verify Site Code and Site Server are correct
- Ensure you have SCCM Console installed on this machine
- Check you have permissions to query SCCM
- Verify ConfigurationManager module is available:
  ```powershell
  Get-Module -ListAvailable -Name ConfigurationManager
  ```

### Collection Issues

**Error**: "Failed to retrieve collection members"

- Verify Collection ID is correct
- Check you have permissions to view the collection
- Ensure collection exists and is not empty

### Computer Offline

**Status**: "Offline" with red border

- Computer may be powered off
- Computer may be disconnected from network
- Firewall may be blocking ICMP (ping)
- Check computer status in SCCM console

### Log File Not Found

**Status**: "No Log File"

- The deferral tool may not have run yet on this computer
- Log file path may be incorrect
- Check the log path configuration in `DeferTSConfig.xml` on the client
- Verify the application has been deployed and executed

### Access Denied

**Status**: "Error" with "Access Denied" message

- Run the tool as an account with administrative access to target computers
- Verify administrative shares (C$) are enabled on target computers
- Check Windows Firewall settings on target computers

### Slow Scanning

- Scanning is performed sequentially (one computer at a time)
- Offline computers timeout after 1 second (ping timeout)
- Log file reading depends on network speed and file size
- For large collections (100+ computers), expect several minutes

## Technical Details

### Architecture

- **Language**: PowerShell 5.1
- **UI Framework**: WPF (Windows Presentation Foundation)
- **SCCM Interface**: ConfigurationManager PowerShell module
- **Network Protocol**: SMB/CIFS (for UNC path access)
- **Connectivity Test**: ICMP (ping)

### Classes

- `ComputerStatus`: Stores status information for each computer
- `LogEntry`: Represents a single log line (timestamp, level, message)

### Functions

- `Test-ComputerOnline`: Pings computer to check if reachable
- `Parse-LogFile`: Reads and parses log file, extracts deferral/trigger information
- `Connect-ToSCCM`: Establishes connection to SCCM site
- `Get-CollectionMembers`: Retrieves all members of a collection
- `Show-MainWindow`: Creates and displays the WPF UI

## Security Considerations

1. **Credentials**: The tool runs under your current user context. Ensure you have:
   - SCCM admin rights
   - Network access to target computers
   - Permission to read files on target computers

2. **Data Exposure**: Log files may contain sensitive information:
   - Computer names
   - User interaction timestamps
   - Task Sequence details
   - Export CSV files should be stored securely

3. **Network Traffic**: The tool generates network traffic:
   - SCCM queries (minimal)
   - ICMP ping to each computer
   - SMB connections to read log files

## Best Practices

1. **Test First**: Test on a small collection before scanning large collections
2. **Network Timing**: Run during low-usage hours for large collections
3. **Permissions**: Use a dedicated service account with appropriate permissions
4. **Regular Scans**: Schedule regular scans to monitor deferral trends
5. **Export Data**: Export and archive results for historical tracking
6. **Filter Collections**: Create specific collections for monitoring rather than scanning "All Systems"

## Limitations

1. **Sequential Processing**: Computers are scanned one at a time (not parallel)
2. **No Real-time Updates**: Scan is a point-in-time snapshot
3. **Network Dependent**: Requires network connectivity to all target computers
4. **Windows Only**: Works only with Windows SCCM clients
5. **Log Format Specific**: Designed for logs created by `deferTS.ps1` script

## Future Enhancements (Potential)

- Parallel scanning (multi-threading) for faster processing
- Real-time monitoring mode with auto-refresh
- Filtering and sorting options in UI
- Save/load connection settings
- Historical trend analysis
- Email alerting for critical statuses
- Integration with other SCCM reports

## Support

For issues or questions:
1. Review the Troubleshooting section above
2. Check SCCM logs on the site server
3. Verify network connectivity to target computers
4. Review Windows Event Logs on target computers

## Version History

- **1.0.0** (2025-11-27): Initial release
  - SCCM connection and collection scanning
  - Log file parsing with deferral/trigger detection
  - WPF UI with traffic lights
  - Online/offline detection
  - CSV export functionality

## Related Files

- `deferTS.ps1`: The client-side deferral tool that creates the logs
- `DeferTSConfig.xml`: Configuration file for the deferral tool
- `SCCM_LOG_READER_README.md`: This file

## License

This tool is provided as-is for diagnostic and monitoring purposes within SCCM environments.
