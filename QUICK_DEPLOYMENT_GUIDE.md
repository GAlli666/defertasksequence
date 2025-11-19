# Quick Deployment Guide

## Pre-Deployment Checklist

- [ ] Task Sequence created and deployed as "Available"
- [ ] Task Sequence Package ID noted (e.g., ABC00123)
- [ ] Banner image created (900x125 pixels)
- [ ] Network share created for content
- [ ] DeferTSConfig.xml edited with correct Package ID
- [ ] Test collection created

## 5-Minute Setup

### 1. Edit Configuration (2 minutes)

Open `DeferTSConfig.xml` and update:

```xml
<PackageID>YOUR_TS_PACKAGE_ID</PackageID>
```

Optionally customize:
- MaxDeferrals (default: 3)
- Messages
- Colors
- Registry path

### 2. Create Network Share (1 minute)

```powershell
# Create share
New-Item -Path "C:\SCCMSources\DeferralTool" -ItemType Directory
New-SmbShare -Name "DeferralTool" -Path "C:\SCCMSources\DeferralTool" -ReadAccess "Everyone"

# Copy files
Copy-Item .\* -Destination "C:\SCCMSources\DeferralTool\" -Force
```

### 3. Create SCCM Application (2 minutes)

**Application:**
- Name: `Task Sequence Deferral - Windows 11 Upgrade`
- Type: Script Installer

**Deployment Type:**

| Setting | Value |
|---------|-------|
| Content Location | `\\server\share\DeferralTool` |
| Install Command | `powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File ".\deferTS.ps1"` |
| Uninstall Command | `cmd.exe /c exit 0` |
| Detection Method | Script: `Detect-Windows11.ps1` |
| Install for | System |
| Installation program visibility | Normal |
| Maximum run time | 120 minutes |

**Return Code:**
- Code `1` = Soft Reboot (User deferred)

### 4. Deploy

**Deployment Settings:**
- Purpose: **Required**
- Available: As soon as possible
- Deadline: Set recurring (e.g., every 4 hours)
- Rerun: **Rerun if failed previous attempt** ✓

## Important: Deferral Logic & UI Behavior

**Key Behavior:**
- Deferral count increments **IMMEDIATELY when script starts** (before UI shows)
- This ensures even Task Manager kills count as deferrals
- Clicking "Defer" button does NOT increment again (already incremented)
- Only successful "Install" resets count to 0

**UI Features:**
- **No window controls:** No X, minimize, or maximize buttons
- **Alt+F4 blocked:** Cannot force-close the window
- **Must click a button:** Only way to proceed is Defer or Install
- **Limit reached:** Skips main dialog, shows countdown only (no buttons)

**Example:**
1. Start with count = 0
2. Script runs → count becomes 1 (before UI)
3. UI shows "You can defer this installation 2 more times" (3 max - 1 used)
4. User clicks "Defer" → count stays 1, exits with code 1
5. Next run → count becomes 2, shows "You can defer this installation 1 more time"
6. Next run → count becomes 3 (at limit)
7. **Main dialog skipped** → countdown shows immediately → auto-installs

## Test Commands

### Test Main Script
```powershell
# Run interactively
powershell.exe -ExecutionPolicy Bypass -File ".\deferTS.ps1"

# Check registry (replace ABC00123 with your Package ID)
# Should show DeferralCount AND metadata (Vendor, Product, Version, FirstRunDate)
Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123"
```

### Test Window Controls Blocked (Important!)
```powershell
# Reset count (replace ABC00123 with your Package ID)
Remove-Item -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123" -Recurse -Force -ErrorAction SilentlyContinue

# Run script
powershell.exe -ExecutionPolicy Bypass -File ".\deferTS.ps1"

# Try Alt+F4 - should NOT close
# Try Esc - should NOT close
# Verify NO X button in window
# Must click "Defer" or "Install Now" to proceed

# In another PowerShell window, check registry BEFORE clicking anything
Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123"
# Should show DeferralCount=1 AND all metadata values
```

### Test Metadata
```powershell
# Check all registry values (replace ABC00123 with your Package ID)
Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123"

# Expected output:
# DeferralCount : 1
# Vendor        : Your Company Name
# Product       : Windows 11 Upgrade
# Version       : 1.0.0
# FirstRunDate  : 2025-11-19 14:30:25
```

### Test Deferral Limit Reached
```powershell
# Set count to max (replace ABC00123 with your Package ID)
Set-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123" -Name "DeferralCount" -Value 3

# Run script
powershell.exe -ExecutionPolicy Bypass -File ".\deferTS.ps1"

# Main dialog should be SKIPPED
# Only countdown screen shows (no buttons)
# Auto-starts Task Sequence after countdown
```

### Test Detection
```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Detect-Windows11.ps1"
# Should output "Detected" on Windows 11
```

### Reset Deferrals
```powershell
# Complete reset (replace ABC00123 with your Package ID)
Remove-Item -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\ABC00123" -Recurse -Force -ErrorAction SilentlyContinue
```

## Common Issues

| Issue | Solution |
|-------|----------|
| UI doesn't appear | Check log: `C:\Windows\Temp\TaskSequenceDeferral.log` |
| TS doesn't start | Verify Package ID in config matches TS |
| Deferrals don't track | Check registry path permissions |
| Detection fails | Test detection script manually |

## Deployment Timeline

| Time | Action |
|------|--------|
| Day 1 | Deploy to pilot group (10-20 devices) |
| Day 2-3 | Monitor compliance, logs, user feedback |
| Day 4 | Adjust settings if needed |
| Day 5 | Deploy to broader test group (100+ devices) |
| Week 2 | Deploy to production collections |

## Monitoring

**SCCM Console:**
- Deployment Status
- Compliance %
- Error Codes

**Client Logs:**
- `AppEnforce.log` - Installation attempts
- `AppDiscovery.log` - Detection method
- `C:\Windows\Temp\TaskSequenceDeferral.log` - Script log

**Registry Check:**
```powershell
# Check deferral counts across devices (replace ABC00123 with your Package ID)
Invoke-Command -ComputerName (Get-Content .\computers.txt) -ScriptBlock {
    $packageID = "ABC00123"
    $regPath = "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral\$packageID"

    if (Test-Path $regPath) {
        Get-ItemProperty -Path $regPath | Select-Object PSComputerName, DeferralCount, Vendor, Product, Version, FirstRunDate
    }
    else {
        [PSCustomObject]@{
            PSComputerName = $env:COMPUTERNAME
            Status = "Not Started"
        }
    }
}

# View all Package IDs being tracked
Get-ChildItem -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral" -ErrorAction SilentlyContinue | Select-Object PSChildName
```

## Rollback Plan

If issues occur:

1. **Suspend Deployment** in SCCM
2. **Delete Deployment** (doesn't remove application)
3. **Fix issues** in scripts/config
4. **Update content** on distribution points
5. **Re-deploy** to pilot group
6. **Monitor** before broader rollout

## Support Resources

- **Full Documentation:** See README.md
- **Logs:** `C:\Windows\Temp\TaskSequenceDeferral.log`
- **Registry:** `HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral`
- **SCCM Logs:** `C:\Windows\CCM\Logs\AppEnforce.log`
