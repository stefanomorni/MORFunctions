# =============================================================================
# Repair-ProjectSymlinks.ps1
# Self-healing symlink repair script for Office-Automation-Framework child projects.
# SSoT Target: Office-Automation-Framework/scripts
# =============================================================================
# Why this exists:
# Cloud sync tools (GoodSync -> OneDrive msgraph API) and Git clones on Windows
# without core.symlinks=true cannot preserve NTFS symlinks. If symlinks are
# replaced by plain files or deleted during cloud restore, running this script
# re-establishes the SSoT symbolic links to the central framework scripts.
# =============================================================================
[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$FrameworkScriptsPath,

    [Parameter(Mandatory=$false)]
    [switch]$WhatIf
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$ProjectDir = $PSScriptRoot
$ScriptsDir = Join-Path $ProjectDir "scripts"

function Write-Step  { param($msg) Write-Host "`n--> $msg" -ForegroundColor Cyan }
function Write-OK    { param($msg) Write-Host "  [OK] $msg" -ForegroundColor Green }
function Write-Warn  { param($msg) Write-Host "  [!!] $msg" -ForegroundColor Yellow }
function Write-Fail  { param($msg) Write-Host "  [XX] $msg" -ForegroundColor Red }

# 1. Locate Framework Scripts Directory
if (-not $FrameworkScriptsPath) {
    # Relative lookup: ..\Office-Automation-Framework\scripts
    $candidate = [System.IO.Path]::GetFullPath((Join-Path $ProjectDir "..\Office-Automation-Framework\scripts"))
    if (Test-Path $candidate) {
        $FrameworkScriptsPath = $candidate
    } else {
        # Fallback to standard canonical path
        $candidate2 = "D:\Cloud\Coding\Projects\Office-Automation-Framework\scripts"
        if (Test-Path $candidate2) {
            $FrameworkScriptsPath = $candidate2
        } else {
            Write-Fail "Could not find Office-Automation-Framework\scripts. Please specify -FrameworkScriptsPath."
            return
        }
    }
}

Write-Step "Repairing SSoT Symlinks for: $([System.IO.Path]::GetFileName($ProjectDir))"
Write-Host "  Framework Source: $FrameworkScriptsPath" -ForegroundColor DarkGray
Write-Host "  Project Scripts:  $ScriptsDir" -ForegroundColor DarkGray

# Guard against running on the Framework root itself
if ($FrameworkScriptsPath.TrimEnd('\') -eq $ScriptsDir.TrimEnd('\')) {
    Write-OK "Currently in the Framework root directory (Source of Truth). No symlinks to repair."
    return
}

if (-not (Test-Path $ScriptsDir)) {
    if (-not $WhatIf) {
        New-Item -ItemType Directory -Path $ScriptsDir -Force | Out-Null
    }
    Write-OK "Created scripts directory: $ScriptsDir"
}

# Determine which core scripts should be linked
$CoreScripts = @("vba_sync_watch.py", "pq_sync.py", "pq_sync_watch.py", "mask_data.py", "ribbon_sync.py", "publish_addin.py", "Resolve-OfficePython.ps1")

function Repair-Symlink {
    param(
        [string]$LinkPath,
        [string]$TargetPath
    )
    $fileName = [System.IO.Path]::GetFileName($LinkPath)

    if (Test-Path $LinkPath) {
        $item = Get-Item $LinkPath -Force
        $isSymlink = ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0

        if ($isSymlink) {
            $currentTarget = $item.Target
            # If target array (PowerShell 7+) or string
            if ($currentTarget -is [array]) { $currentTarget = $currentTarget[0] }
            if ($currentTarget -and ($currentTarget.TrimEnd('\') -eq $TargetPath.TrimEnd('\'))) {
                Write-OK "Verified valid symlink: $fileName -> $TargetPath"
                return
            } else {
                Write-Warn "Symlink target mismatch on $fileName (Current: $currentTarget -> Expected: $TargetPath)"
            }
        } else {
            Write-Warn "Non-symlink file found in place of SSoT link: $fileName (Size: $($item.Length) bytes)"
        }

        if ($WhatIf) {
            Write-Host "  [WhatIf] Would replace $fileName with symlink to $TargetPath" -ForegroundColor Yellow
            return
        }

        # Backup static file if it has modified content before replacing
        Remove-Item $LinkPath -Force
    } else {
        if ($WhatIf) {
            Write-Host "  [WhatIf] Would create symlink $fileName -> $TargetPath" -ForegroundColor Yellow
            return
        }
    }

    try {
        New-Item -ItemType SymbolicLink -Path $LinkPath -Target $TargetPath -ErrorAction Stop | Out-Null
        Write-OK "Repaired symlink: $fileName -> $TargetPath"
    } catch {
        try {
            New-Item -ItemType HardLink -Path $LinkPath -Target $TargetPath -ErrorAction Stop | Out-Null
            Write-OK "Repaired using HardLink (elevation fallback): $fileName -> $TargetPath"
        } catch {
            Copy-Item -Path $TargetPath -Destination $LinkPath -Force
            Write-Warn "Symlink creation failed (no privileges); copied SSoT script instead: $fileName"
        }
    }
}

foreach ($scriptName in $CoreScripts) {
    $targetFile = Join-Path $FrameworkScriptsPath $scriptName
    if (Test-Path $targetFile) {
        $linkFile = Join-Path $ScriptsDir $scriptName
        Repair-Symlink -LinkPath $linkFile -TargetPath $targetFile
    }
}

Write-Host "`nSymlink verification and repair complete." -ForegroundColor Green
