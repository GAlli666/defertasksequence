<#
.SYNOPSIS
    Detection Method Script for Windows 11

.DESCRIPTION
    This script is used as a detection method for SCCM/ConfigMgr Application deployment.
    Returns success (exit 0) when the machine is running Windows 11.
    Returns failure (exit 1) when the machine is not running Windows 11.

.NOTES
    Author: Claude
    Date: 2025-11-19

    Usage in SCCM:
    - Detection Method Type: Use a custom script to detect the presence of this application
    - Script Type: PowerShell
    - Script: This file
    - Run script as 32-bit process: No (recommended)

    The script will output "Detected" if Windows 11 is found, which SCCM interprets as success.
#>

try {
    # Get Windows version information
    $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

    # Extract build number
    $buildNumber = $osInfo.BuildNumber

    # Windows 11 started with build 22000
    # Current Windows 11 builds: 22000+
    if ([int]$buildNumber -ge 22000) {
        # Output "Detected" for SCCM
        Write-Output "Detected"

        # Exit with success code
        exit 0
    }
    else {
        # Windows 10 or older - not detected
        # No output means not detected in SCCM
        exit 1
    }
}
catch {
    # Error occurred - treat as not detected
    Write-Error "Detection script failed: $_"
    exit 1
}
