# Task Sequence Deferral Tool

A modern, user-friendly PowerShell tool for providing deferral options when deploying Task Sequences via SCCM/ConfigMgr Application Model.

## Features

- **Modern WPF UI** (900x700 pixels) with customizable branding
- **Deferral tracking** via registry with configurable limits
- **Two-step confirmation** for installations
- **Automatic countdown** when deferral limit reached
- **Fully configurable** via XML file (colors, text, settings)
- **Comprehensive logging** for troubleshooting
- **Exit code 1** on deferral to trigger SCCM retry

## Files Included

| File | Description |
|------|-------------|
| `deferTS.ps1` | Main deferral tool script with WPF UI |
| `DeferTSConfig.xml` | Configuration file for all settings |
| `Detect-Windows11.ps1` | Detection method script for Windows 11 |
| `README.md` | This file - deployment instructions |
| `banner.png` | Your company banner image (900x125 pixels) - **YOU NEED TO ADD THIS** |

## Prerequisites

- SCCM/ConfigMgr environment
- PowerShell 3.0 or higher on target machines
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
   <MaxDeferrals>3</MaxDeferrals>

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
     - Copy all files (deferTS.ps1, DeferTSConfig.xml, banner.png) to this location
     - Set content location to this path

   - **Installation program:**
     ```powershell
     powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ".\deferTS.ps1"
     ```

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
   - Shows main message and updated deferral count
   - Two buttons: "Defer" and "Install Now"
   - **Important:** Force-closing the window counts as a deferral (count already incremented)

2. **User Clicks "Defer":**
   - Script exits with code 1 (triggers SCCM retry)
   - UI closes
   - Deferral count remains incremented (no additional increment)

3. **User Clicks "Install Now":**
   - Secondary confirmation appears
   - "Are you sure?" message with two buttons:
     - "Yes, Install Now" - proceeds with installation, resets deferral count to 0
     - "No, Go Back" - returns to main screen

4. **User Force-Closes Window:**
   - Deferral count already incremented at script start
   - Script exits with code 1 (triggers SCCM retry)
   - This prevents users from gaming the system by force-closing

5. **Deferral Limit Reached:**
   - Final message appears automatically
   - Countdown timer (configurable seconds)
   - No buttons - auto-starts Task Sequence
   - Installation begins when countdown reaches 0

### Technical Flow

```
Script Start
    ↓
Load Configuration (XML)
    ↓
Read Registry (Current Deferral Count)
    ↓
INCREMENT DEFERRAL COUNT IMMEDIATELY (if not at limit)
    ↓
Show WPF UI (displays already-incremented count)
    ↓
User Choice?
    ├─ Defer → Exit 1 (count already incremented)
    ├─ Install → Confirmation → Start Task Sequence → Reset Count to 0 → Exit 0
    ├─ Force Close → Exit 1 (count already incremented)
    └─ Limit Reached → Countdown → Start Task Sequence → Reset Count to 0 → Exit 0
```

## Registry Details

**Default Path:**
```
HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral
Value: DeferralCount (DWORD)
```

**Behavior:**
- **Incremented IMMEDIATELY when script starts** (before UI shows)
- This ensures force-closing the window counts as a deferral
- Clicking "Defer" button does NOT increment again (already incremented)
- Only clicking "Install" and successfully starting TS resets count to 0
- Persists across reboots
- Can be manually modified if needed for testing

**Example Scenario:**
1. Script starts with count = 0
2. Count immediately increments to 1 (before UI shows)
3. User sees "You have 2 deferral(s) remaining" (3 max - 1 used = 2 remaining)
4. User clicks "Defer" - count stays at 1, exits with code 1
5. Next time script runs: count = 1, increments to 2, shows "1 remaining"

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
<MaxDeferrals>3</MaxDeferrals>

<!-- Countdown duration (seconds) when limit reached -->
<FinalMessageDuration>30</FinalMessageDuration>
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
3. Manually check registry:
   ```powershell
   Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral"
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

1. Reset registry: `Remove-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral" -Name "DeferralCount" -ErrorAction SilentlyContinue`
2. Run script manually
3. **Check registry BEFORE clicking anything** - should already be incremented to 1
4. Verify UI shows "You have 2 deferral(s) remaining" (assuming max = 3)
5. Click "Defer" - script exits with code 1
6. **Check registry again** - should still be 1 (not incremented again)
7. Run script again - registry should increment to 2
8. Verify UI shows "You have 1 deferral(s) remaining"

### Test Scenario 2: Force Close (Important!)

1. Reset registry to 0
2. Run script manually
3. **Check registry immediately** - should already be 1
4. Force close the window (X button or Alt+F4)
5. **Check registry** - should still be 1 (deferral counted)
6. This proves force-close = deferral used

### Test Scenario 3: Install Accepted

1. Run script
2. Click "Install Now"
3. Verify confirmation appears
4. Click "Yes, Install Now"
5. Verify Task Sequence starts
6. **Check registry** - should be reset to 0

### Test Scenario 4: Deferral Limit Reached

1. Manually set registry value to max deferrals:
   ```powershell
   Set-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral" -Name "DeferralCount" -Value 3
   ```
2. Run script
3. Registry should NOT increment further (already at limit)
4. Verify countdown appears immediately
5. Verify Task Sequence auto-starts after countdown

### Reset for Testing

```powershell
# Reset deferral count
Remove-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral" -Name "DeferralCount" -ErrorAction SilentlyContinue

# Or set to specific value
Set-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral" -Name "DeferralCount" -Value 0
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

- **v1.0** (2025-11-19)
  - Initial release
  - WPF UI with modern design
  - Deferral tracking via registry
  - Configurable via XML
  - Windows 11 detection method
