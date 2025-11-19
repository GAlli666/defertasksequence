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

## Test Commands

### Test Main Script
```powershell
# Run interactively
powershell.exe -ExecutionPolicy Bypass -File ".\deferTS.ps1"

# Check registry
Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral" -Name DeferralCount
```

### Test Detection
```powershell
powershell.exe -ExecutionPolicy Bypass -File ".\Detect-Windows11.ps1"
# Should output "Detected" on Windows 11
```

### Reset Deferrals
```powershell
Remove-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral" -Name DeferralCount -ErrorAction SilentlyContinue
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
# Check deferral counts across devices
Invoke-Command -ComputerName (Get-Content .\computers.txt) -ScriptBlock {
    Get-ItemProperty -Path "HKLM:\SOFTWARE\YourCompany\TaskSequenceDeferral" -Name DeferralCount -ErrorAction SilentlyContinue
}
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
