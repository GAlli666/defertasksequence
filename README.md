# Task Sequence Deferral Tool

A modern, user-friendly PowerShell tool for providing deferral options when deploying Task Sequences via SCCM/ConfigMgr Application Model.

## Features

- **Modern WPF UI** (900x700 pixels) with customizable branding
- **Deferral tracking** via registry with configurable limits
- **Configurable timeout** on main window to prevent indefinite delays
- **Two-step confirmation** for installations
- **Automatic countdown** when deferral limit reached
- **Fully configurable** via XML file (colors, text, settings)
- **Comprehensive logging** for troubleshooting
- **Exit code 1** on deferral to trigger SCCM retry
- **Hidden console** support for SCCM deployments via VBScript wrapper
- **Task Sequence detection** - automatically exits if Task Sequence already running

## Files Included

| File | Description |
|------|-------------|
| `deferTS.ps1` | Main deferral tool script with WPF UI |
| `LaunchDeferTS.vbs` | VBScript wrapper for hidden console execution from SCCM |
| `DeferTSConfig.xml` | Configuration file for all settings |
| `Detect-Windows11.ps1` | Detection method script for Windows 11 |
| `README.md` | This file - deployment instructions |
| `banner.png` | Your company banner image (900x125 pixels) - **YOU NEED TO ADD THIS** |

## Prerequisites

- SCCM/ConfigMgr environment
- PowerShell 5.1 or higher on target machines
- Task Sequence already deployed as "Available"
- Windows 7+ target machines (WPF support)

## Setup Instructions

### 1. Prepare Your Files

1. **Add your banner image:**
   - Create or use an existing banner image (900x125 pixels recommended)
   - Name it `banner.png` (or update `DeferTSConfig.xml` to match your filename)
   - Place it in the same folder as the scripts

2. **Edit `DeferTSConfig.xml`:**
   ```xml
   <!-- Update with your Task Sequence Package ID -->
   <PackageID>ABC00123</PackageID>

   <!-- Customize deferral limit (default: 3) -->
   <!-- This is the number of times users can defer - if set to 3, users see "defer 3 more times" on first run -->
   <MaxDeferrals>3</MaxDeferrals>

   <!-- Customize main window timeout in minutes (default: 60) -->
   <MainWindowTimeoutMinutes>60</MainWindowTimeoutMinutes>

   <!-- Customize UI messages -->
   <MainMessage>Your custom message here...</MainMessage>

   <!-- Customize colors (RGB: 0-255) -->
   <AccentColor>20, 40, 70</AccentColor>
   <FontColor>45, 45, 45</FontColor>
   <BackgroundColor>245, 247, 250</BackgroundColor>
   ```

### 2. Create Application in SCCM

1. **Open SCCM Console**
   - Navigate to: `Software Library > Application Management > Applications`

2. **Create New Application**
   - Right-click `Applications` > `Create Application`
   - Type: `Manually specify the application information`
   - Name: `Task Sequence Deferral - Windows 11 Upgrade`
   - Publisher: `Your Company`

3. **Add Deployment Type**
   - Type: `Script Installer`
   - Name: `PowerShell Deferral Script`

   - **Content location:**
     - Create a network share (e.g., `\\server\share\DeferralTool`)
     - Copy all files (deferTS.ps1, LaunchDeferTS.vbs, DeferTSConfig.xml, banner.png) to this location
     - Set content location to this path

   - **Installation program (RECOMMENDED - Hidden Console):**
     ```cmd
     wscript.exe ".\LaunchDeferTS.vbs"
     ```
     **Note:** This is the recommended method as it ensures no console window is visible to users.

   - **Alternative Installation program (PowerShell direct):**
     ```powershell
     powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ".\deferTS.ps1"
     ```
     **Note:** This method may briefly show a console window on some systems.

   - **Uninstall program:**
     ```powershell
     cmd.exe /c exit 0
     ```

   - **Detection Method:**
     - Type: `Use a custom script to detect this deployment type`
     - Script Type: `PowerShell`
     - Script: Browse and select `Detect-Windows11.ps1`
     - Run script as 32-bit process: `No`

   - **User Experience:**
     - Installation behavior: `Install for system`
     - Logon requirement: `Whether or not a user is logged on`
     - Installation program visibility: `Normal`
     - Maximum allowed run time: `120 minutes`
     - Estimated installation time: `60 minutes`

   - **Return Codes:**
     - Add custom return code:
       - Code: `1`
       - Type: `Soft Reboot`
       - Description: `User deferred installation`

4. **Distribute Content**
   - Right-click the application
   - Select `Distribute Content`
   - Select your distribution points

### 3. Deploy the Application

1. **Create Deployment**
   - Right-click your application
   - Select `Deploy`

   - **Collection:**
     - Choose your target collection (Windows 10 devices)

   - **Deployment Settings:**
     - Purpose: `Required` (Mandatory)
     - Make available to: `Only Configuration Manager Clients`

   - **Scheduling:**
     - Available time: `As soon as possible`
     - Installation deadline: Set to your desired deadline
     - Schedule re-evaluation: `Yes`
     - Rerun behavior: `Rerun if failed previous attempt`

   - **User Experience:**
     - User notifications: `Display in Software Center and show all notifications`
     - Software Installation: `Allow`
     - System restart: `If required by Software Center, allow user to restart`

   - **Deployment Schedule (Recurring):**
     - To create recurring schedule for deferrals:
       - Set installation deadline with recurring schedule
       - Example: Every 4 hours during business hours
       - This ensures the deferral prompt reappears after user defers

### 4. Configure Recurring Schedule

For the deferral system to work properly:

1. **Deployment Schedule Options:**
   - Set a recurring installation deadline schedule
   - Example schedule:
     - Recur every: `4 hours`
     - Between: `8:00 AM and 5:00 PM`
     - On days: `Monday through Friday`

2. **Rerun Behavior:**
   - Ensure `Rerun if failed previous attempt` is enabled
   - This allows the script to exit with code 1 (defer) and retry later

## How It Works

### User Flow

1. **First Launch (Deferrals Available):**
   - **Deferral count incremented IMMEDIATELY** (before UI shows)
   - Modern UI appears with company branding
   - Shows main message with text: "You can defer this installation X more times"
   - Two buttons: "Defer" and "Install Now"
   - **Window controls disabled:** No X, minimize, or maximize buttons
   - **Alt+F4 blocked:** Cannot force-close the window
   - **Timeout warning displayed:** Red text showing deadline and live countdown
   - **Auto-install on timeout:** If no button clicked within configured time, installation begins automatically
   - Must click a button to proceed (Defer or Install)

2. **User Clicks "Defer":**
   - Script exits with code 1 (triggers SCCM retry)
   - UI closes
   - Deferral count remains incremented (no additional increment)

3. **User Clicks "Install Now":**
   - Secondary confirmation appears
   - "Are you sure?" message with two buttons:
     - "Yes, Install Now" - proceeds with installation, resets deferral count to 0
     - "No, Go Back" - returns to main screen

4. **User Attempts Force-Close:**
   - Alt+F4, Esc, and system close methods are blocked
   - Window cannot be closed except by clicking a button
   - If user kills process (Task Manager), deferral already counted

5. **Deferral Limit Reached:**
   - **Main dialog skipped entirely** - goes straight to countdown
   - Only the countdown screen shows (no buttons, no options)
   - Final message with countdown timer
   - Auto-starts Task Sequence when countdown reaches 0
   - Cannot be cancelled or deferred

### Technical Flow

```
Script Start
    ↓
Load Configuration (XML)
    ↓
Check if TSManager.exe Running
    ├─ YES → Exit 1618 (TS already running - prevent interference)
    └─ NO → Continue
    ↓
Read Registry (Current Deferral Count)
    ↓
INCREMENT DEFERRAL COUNT IMMEDIATELY (if not at limit)
    ↓
Show WPF UI (Alt+F4 blocked, no window controls)
    ↓
Deferrals Remaining?
    ├─ YES → Show Main Dialog
    │         ├─ Defer → Exit 1 (count already incremented)
    │         └─ Install → Confirmation → Start TS → Reset Count to 0 → Exit 0
    │
    └─ NO → Skip Main Dialog → Show Countdown Only
             └─ Auto Start TS → Reset Count to 0 → Exit 0

Note: Force-close blocked. Task Manager kill = deferral already counted.
```

## Registry Details

**Registry Structure:**

The script uses a Package ID-specific subfolder to allow reuse for multiple Task Sequences:

```
HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\<PackageID>\
```

**Example with Package ID "ABC00123":**
```
HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123\
    DeferralCount (DWORD)       - Current deferral count
    Vendor (String)             - Company/vendor name (from config)
    Product (String)            - Product name (from config)
    Version (String)            - Version number (from config)
    FirstRunDate (String)       - Date/time of first script execution
```

**Registry Values:**

| Value | Type | Description |
|-------|------|-------------|
| `DeferralCount` | DWORD | Number of times installation has been deferred |
| `Vendor` | String | Vendor/company name from config file |
| `Product` | String | Product name from config file |
| `Version` | String | Version number from config file |
| `FirstRunDate` | String | Timestamp when script first ran (auto-set) |

**Behavior:**
- **DeferralCount incremented IMMEDIATELY when script starts** (before UI shows)
- This ensures force-closing the window counts as a deferral
- Clicking "Defer" button does NOT increment again (already incremented)
- Only clicking "Install" and successfully starting TS resets count to 0
- **Metadata set on first run:** Vendor, Product, Version, FirstRunDate
- **Version updated on subsequent runs** if changed in config
- FirstRunDate never changes after initial set
- Persists across reboots
- Can be manually modified if needed for testing

**Reusability:**
- Each Task Sequence gets its own subfolder (by Package ID)
- Can deploy multiple TS deferral tools without conflicts
- Example:
  - Windows 11 Upgrade (ABC00123) → `...\ABC00123\`
  - Office 365 Install (XYZ00456) → `...\XYZ00456\`
  - Security Patches (DEF00789) → `...\DEF00789\`

**Example Scenario (with MaxDeferrals=3 in config):**
1. Script starts with count = 0
2. Count immediately increments to 1 (before UI shows)
3. Metadata set: Vendor, Product, Version, FirstRunDate
4. User sees "You can defer this installation 3 more times" (config value: 3)
5. User clicks "Defer" - count stays at 1, exits with code 1
6. Next run: count = 1, increments to 2, shows "2 more times"
7. Next run: count = 2, increments to 3, shows "1 more time"
8. Next run: count = 3, increments to 4, limit reached - no defer option

## Customization Guide

### Colors

Edit `DeferTSConfig.xml` and adjust RGB values:

```xml
<!-- Darker Navy Blue (Banner) -->
<AccentColor>20, 40, 70</AccentColor>

<!-- Dark Gray (Text) -->
<FontColor>45, 45, 45</FontColor>

<!-- Light Blue-Gray (Background) -->
<BackgroundColor>245, 247, 250</BackgroundColor>
```

**Color Examples:**
- Navy Blue: `0, 51, 102`
- Dark Blue: `20, 40, 70`
- Royal Blue: `65, 105, 225`
- White: `255, 255, 255`
- Black: `0, 0, 0`
- Gray: `128, 128, 128`

### Messages

Customize all text in `DeferTSConfig.xml`:

- `WindowTitle` - Window title bar text
- `BannerText` - Banner text (if no image)
- `MainMessage` - Initial screen message
- `SecondaryMessage` - Confirmation screen message
- `FinalMessage` - Countdown screen message

### Deferral Settings

```xml
<!-- Number of times user can defer -->
<!-- This is what users will see - if set to 3, users see "defer 3 more times" on first run -->
<MaxDeferrals>3</MaxDeferrals>

<!-- Main window timeout in minutes (default: 60) -->
<!-- If user doesn't make a choice, installation begins automatically -->
<MainWindowTimeoutMinutes>60</MainWindowTimeoutMinutes>

<!-- Countdown duration (seconds) when limit reached -->
<FinalMessageDuration>30</FinalMessageDuration>

<!-- Timeout warning text (shown in red on main window) -->
<TimeoutWarningText>Installation will begin</TimeoutWarningText>
<TimeoutNoInputText>if no user input is given</TimeoutNoInputText>
```

### Banner Image

**Requirements:**
- Recommended size: 900x125 pixels
- Supported formats: PNG, JPG, BMP
- Place in same folder as script
- Update filename in config if different:

```xml
<BannerImagePath>your-custom-banner.png</BannerImagePath>
```

**Design Tips:**
- Use company logo and branding
- Keep text minimal (will be overlaid)
- Use high-contrast colors
- Test on different screen resolutions

## Troubleshooting

### Script Doesn't Run

1. Check execution policy:
   ```powershell
   Get-ExecutionPolicy
   ```

2. Verify content location accessible:
   ```powershell
   Test-Path "\\server\share\DeferralTool\deferTS.ps1"
   ```

3. Check SCCM client logs:
   - AppEnforce.log
   - AppDiscovery.log

### UI Doesn't Appear

1. Check WPF prerequisites installed
2. Verify config file exists and is valid XML
3. Check log file: `C:\Windows\Temp\TaskSequenceDeferral.log`

### Task Sequence Doesn't Start

1. Verify Package ID is correct in config
2. Check Task Sequence is deployed as "Available"
3. Test manually via Software Center
4. Review SCCM client logs (SCClient.log)

### Deferrals Not Tracking

1. Check registry path permissions
2. Verify registry path in config is correct
3. Manually check registry (replace ABC00123 with your Package ID):
   ```powershell
   Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123"
   ```
4. Check all registry values:
   ```powershell
   Get-Item -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123" | Select-Object -ExpandProperty Property
   ```

### Detection Method Issues

1. Test detection script manually:
   ```powershell
   powershell.exe -ExecutionPolicy Bypass -File ".\Detect-Windows11.ps1"
   ```

2. Should output "Detected" on Windows 11
3. Should exit with code 0 on Windows 11, code 1 on Windows 10

## Testing

### Test Scenario 1: Fresh Install with Defer

1. Reset registry (replace ABC00123 with your Package ID):
   ```powershell
   Remove-Item -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123" -Recurse -Force -ErrorAction SilentlyContinue
   ```
2. Run script manually
3. **Check registry BEFORE clicking anything** - should show deferral count AND metadata:
   ```powershell
   Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123"
   ```
   - DeferralCount should be 1
   - Vendor, Product, Version, FirstRunDate should all be set
4. Verify UI shows "You can defer this installation 3 more times" (if MaxDeferrals = 3 in config)
5. Verify window has NO X, minimize, or maximize buttons
6. Try Alt+F4 - should NOT close the window
7. Click "Defer" - script exits with code 1
8. **Check registry again** - DeferralCount should still be 1 (not incremented again)
9. Run script again - DeferralCount should increment to 2
10. Verify UI shows "You can defer this installation 2 more times"

### Test Scenario 2: Window Controls Blocked

1. Run script manually
2. Try to close via Alt+F4 - should be blocked
3. Try Esc key - should be blocked
4. Verify NO X button in title bar (WindowStyle=None)
5. Only way to close is clicking "Defer" or "Install Now"
6. If process is killed via Task Manager, deferral already counted (registry already incremented)

### Test Scenario 2b: Main Window Timeout

1. Edit DeferTSConfig.xml and set MainWindowTimeoutMinutes to 1 (for testing)
2. Run script manually
3. Verify timeout warning displays in red:
   - "Installation will begin"
   - Deadline date/time (format: yyyy-MM-dd HH:mm, no seconds)
   - "if no user input is given"
   - Live countdown: "Time remaining: X minutes, Y seconds"
4. Wait for countdown to reach 0
5. Verify installation begins automatically
6. Reset MainWindowTimeoutMinutes to normal value (e.g., 60) after testing

### Test Scenario 3: Install Accepted

1. Run script
2. Click "Install Now"
3. Verify confirmation appears
4. Click "Yes, Install Now"
5. Verify Task Sequence starts
6. **Check registry** - should be reset to 0

### Test Scenario 4: Deferral Limit Reached (Important!)

1. Manually set registry value to max deferrals (replace ABC00123 with your Package ID):
   ```powershell
   # If MaxDeferrals=3 in config, set to 4 (script adds 1 internally)
   Set-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123" -Name "DeferralCount" -Value 4
   ```
2. Run script
3. Registry should NOT increment further (already at limit)
4. **Main dialog should be SKIPPED entirely** - no buttons, no defer option
5. Verify countdown screen appears immediately
6. Verify NO buttons present (cannot cancel or defer)
7. Verify Task Sequence auto-starts after countdown completes

### Test Scenario 5: Metadata Verification

1. Reset registry and run script
2. Check all registry values:
   ```powershell
   Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123"
   ```
3. Verify all values present:
   - DeferralCount (DWORD)
   - Vendor (String)
   - Product (String)
   - Version (String)
   - FirstRunDate (String with timestamp)
4. Update Version in DeferTSConfig.xml (e.g., 1.0.0 → 1.0.1)
5. Run script again
6. Verify Version updated in registry, but FirstRunDate unchanged

### Reset for Testing

```powershell
# Complete reset - removes entire Package ID subfolder (replace ABC00123 with your Package ID)
Remove-Item -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123" -Recurse -Force -ErrorAction SilentlyContinue

# Or just reset deferral count
Set-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123" -Name "DeferralCount" -Value 0

# View all registry values
Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123"
```

## Best Practices

1. **Test thoroughly** in pilot group before production
2. **Set reasonable deferral limits** (3-5 deferrals recommended)
3. **Schedule during business hours** for better user experience
4. **Monitor deployment compliance** regularly
5. **Provide help desk documentation** for user questions
6. **Keep banner image professional** and on-brand
7. **Use clear, concise messaging** in UI
8. **Test on various screen resolutions**
9. **Have rollback plan** if issues arise
10. **Monitor logs** for early issue detection

## Advanced Configuration

### Custom Registry Path

To use a different registry location:

```xml
<RegistryPath>HKLM:\SOFTWARE\MyCompany\MyCustomPath</RegistryPath>
<RegistryValueName>MyDeferralCount</RegistryValueName>
```

### Custom Logging

Enable or change log location:

```xml
<!-- Enable logging -->
<LogPath>C:\Logs\TSDefer.log</LogPath>

<!-- Disable logging -->
<LogPath></LogPath>
```

### Multiple Task Sequences

Create separate config files for different Task Sequences:

```powershell
# Launch with custom config
powershell.exe -ExecutionPolicy Bypass -File ".\deferTS.ps1" -ConfigFile ".\CustomConfig.xml"
```

## Support

For issues or questions:

1. Check log file: `C:\Windows\Temp\TaskSequenceDeferral.log`
2. Review SCCM client logs (AppEnforce.log, AppDiscovery.log)
3. Verify configuration XML is valid
4. Test scripts manually on problem devices
5. Check registry for deferral count issues

## License

Created by Claude for internal use. Modify as needed for your environment.

## Version History

- **v1.1.0** (2025-11-21)
  - Added VBScript wrapper (LaunchDeferTS.vbs) for completely hidden console execution from SCCM
  - Added TSManager.exe detection to prevent running when Task Sequence already in progress
  - Changed exit code to 0 when Task Sequence trigger is attempted (success or failure)
  - Updated documentation with recommended SCCM installation command

- **v1.0.1** (2025-11-19)
  - Added configurable main window timeout feature
  - Auto-installs if user doesn't make a choice within configured time
  - Red warning text with deadline and live countdown
  - All timeout text configurable via XML

- **v1.0** (2025-11-19)
  - Initial release
  - WPF UI with modern design
  - Deferral tracking via registry with Package ID subfolders
  - Metadata tracking (Vendor, Product, Version, FirstRunDate)
  - Immediate deferral count increment (prevents gaming)
  - Window close prevention (Alt+F4, X button blocked)
  - Skip main dialog when deferral limit reached
  - Configurable via XML
  - Windows 11 detection method
