#Requires -RunAsAdministrator

param(
    [string]$IsoPath = ""
)

$ErrorActionPreference = "Continue"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$bundledFontsPath = Join-Path $scriptDir "Fonts"
$systemFontsPath = "$env:WINDIR\Fonts"

Write-Host "`n=======================================`n  Windows Font Restoration Tool v1.0`n=======================================" -ForegroundColor Cyan
Write-Host "Fixes Adobe Acrobat font errors & corrupted system fonts`n" -ForegroundColor Gray

Write-Host "This script will:" -ForegroundColor Yellow
Write-Host "  - Stop font cache services" -ForegroundColor Yellow
Write-Host "  - Clear the system font cache" -ForegroundColor Yellow
Write-Host "  - Take ownership of files in $env:WINDIR\Fonts" -ForegroundColor Yellow
Write-Host "  - Overwrite system font files" -ForegroundColor Yellow
Write-Host "  - Modify font registry entries" -ForegroundColor Yellow
Write-Host ""
$confirm = Read-Host "Type YES to continue or press Enter to abort"
if ($confirm -ne 'YES') {
    Write-Host "Aborted. No changes were made." -ForegroundColor Yellow
    exit 0
}
Write-Host ""

# Ensure Fonts directory exists
if (-not (Test-Path $bundledFontsPath)) {
    New-Item -ItemType Directory -Path $bundledFontsPath | Out-Null
}

# === FONT ACQUISITION ===
Write-Host "[0/6] Checking font bundle..." -ForegroundColor Yellow

$criticalFonts = @(
    "arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf",
    "times.ttf", "timesbd.ttf", "timesi.ttf", "timesbi.ttf",
    "cour.ttf", "courbd.ttf", "couri.ttf", "courbi.ttf"
)
$warningFonts = @(
    "verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf",
    "tahoma.ttf", "tahomabd.ttf",
    "georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf",
    "calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf",
    "segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "seguisym.ttf"
)

$missingCritical = $criticalFonts | Where-Object { -not (Test-Path (Join-Path $bundledFontsPath $_)) }
$missingWarning = $warningFonts  | Where-Object { -not (Test-Path (Join-Path $bundledFontsPath $_)) }

if ($missingCritical.Count -eq 0) {
    $fontCount = @(Get-ChildItem "$bundledFontsPath\*.ttf" -ErrorAction SilentlyContinue).Count
    Write-Host "  ✓ Critical fonts present ($fontCount total) — skipping acquisition" -ForegroundColor Green
    if ($missingWarning.Count -gt 0) {
        Write-Warning "  $($missingWarning.Count) non-critical fonts absent: $($missingWarning -join ', ')"
        Write-Warning "  Re-run with installation media to acquire them, or proceed without them."
    }
}
else {
    Write-Host "  Missing $($missingCritical.Count) critical fonts — locating Windows installation media..." -ForegroundColor Yellow

    # === RESOLUTION CHAIN ===
    $isoFile = $null
    $wimPath = $null
    $isoMounted = $false

    # Tier 1: -IsoPath parameter
    if ($IsoPath -and (Test-Path $IsoPath -PathType Leaf)) {
        $isoFile = $IsoPath
        Write-Host "  Source: ISO specified via -IsoPath" -ForegroundColor Gray
    }

    # Tier 2: .iso file in script directory
    if (-not $isoFile) {
        $isoFile = Get-ChildItem -Path $scriptDir -Filter "*.iso" -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
        if ($isoFile) {
            Write-Host "  Source: ISO auto-detected in script folder: $isoFile" -ForegroundColor Gray
        }
    }

    # Tier 3: Windows installation USB or optical media (removable=2, CD-ROM=5)
    if (-not $isoFile) {
        Write-Host "  No ISO found — scanning for Windows installation media (USB/DVD)..." -ForegroundColor Yellow
        $removableDrives = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -in @(2, 5) }
        $usbMatches = @()
        foreach ($drive in $removableDrives) {
            foreach ($candidate in @("install.wim", "install.esd")) {
                $candidatePath = Join-Path "$($drive.DeviceID)\sources" $candidate
                if (Test-Path $candidatePath) {
                    $usbMatches += [PSCustomObject]@{ Drive = $drive.DeviceID; Path = $candidatePath }
                    break
                }
            }
        }
        if ($usbMatches.Count -gt 0) {
            $wimPath = $usbMatches[0].Path
            Write-Host "  ✓ Windows media found at $($usbMatches[0].Drive) — using directly (no ISO mount needed)" -ForegroundColor Green
            if ($usbMatches.Count -gt 1) {
                Write-Warning "  Multiple media drives detected. Using $($usbMatches[0].Drive). Others: $(($usbMatches[1..($usbMatches.Count-1)] | ForEach-Object { $_.Drive }) -join ', ')"
            }
        }
    }

    # No source found
    if (-not $isoFile -and -not $wimPath) {
        Write-Error @"
No Windows installation media found. Critical fonts missing: $($missingCritical -join ', ')

To obtain a Windows 11 ISO (official source only):
  1. Install a User-Agent switcher browser extension (Chrome or Firefox)
  2. Set your UA to a non-Windows string (e.g., iOS Safari or Android Chrome)
  3. Visit: https://www.microsoft.com/en-us/software-download/windows11
     The page will show a direct ISO download instead of the Media Creation Tool.
  4. Select your language -> 64-bit Download
  5. Place the downloaded .iso in: $scriptDir
     Or run: .\Restore-WindowsFonts.ps1 -IsoPath "C:\path\to\windows11.iso"
     Or plug in a Windows installation USB drive.
"@
        exit 1
    }

    # === ISO MOUNT (only if no USB/DVD WIM was found directly) ===
    if (-not $wimPath) {
        Write-Host "  ISO: $isoFile" -ForegroundColor Gray
        Write-Host "  Mounting ISO..." -ForegroundColor Yellow
        Mount-DiskImage -ImagePath $isoFile | Out-Null
        $driveLetter = $null
        for ($attempt = 0; $attempt -lt 6 -and -not $driveLetter; $attempt++) {
            Start-Sleep -Milliseconds 500
            $driveLetter = (Get-DiskImage -ImagePath $isoFile | Get-Volume).DriveLetter
        }
        if (-not $driveLetter) {
            Dismount-DiskImage -ImagePath $isoFile | Out-Null
            Write-Error "ISO mounted but no drive letter was assigned. Try again."
            exit 1
        }
        $driveLetter = $driveLetter + ":"
        $isoMounted = $true
        Write-Host "  ✓ ISO mounted at $driveLetter" -ForegroundColor Green

        foreach ($candidate in @("install.wim", "install.esd")) {
            $full = Join-Path "$driveLetter\sources" $candidate
            if (Test-Path $full) { $wimPath = $full; break }
        }

        if (-not $wimPath) {
            Dismount-DiskImage -ImagePath $isoFile | Out-Null
            Write-Error "Could not find install.wim or install.esd under ${driveLetter}\sources\"
            exit 1
        }
    }

    # === WIM MOUNT ===
    $wimMountPath = Join-Path $scriptDir "WimMount"
    if (-not (Test-Path $wimMountPath)) { New-Item -ItemType Directory -Path $wimMountPath | Out-Null }

    $wimIndex = (Get-WindowsImage -ImagePath $wimPath | Select-Object -First 1).ImageIndex

    $wimJob = Start-Job -ArgumentList $wimPath, $wimIndex, $wimMountPath -ScriptBlock {
        param($wimPath, $wimIndex, $wimMountPath)
        Import-Module Dism
        Mount-WindowsImage -ImagePath $wimPath -Index $wimIndex -Path $wimMountPath -ReadOnly | Out-Null
    }
    $elapsed = 0
    while ($wimJob.State -eq 'Running') {
        Write-Progress -Activity "Mounting WIM image" -Status "Elapsed: ${elapsed}s — please wait..." -PercentComplete -1
        Start-Sleep -Seconds 1
        $elapsed++
    }
    Write-Progress -Activity "Mounting WIM image" -Completed

    if ($wimJob.State -eq 'Failed') {
        $jobError = Receive-Job -Job $wimJob 2>&1
        Remove-Job -Job $wimJob
        if ($isoMounted) { Dismount-DiskImage -ImagePath $isoFile | Out-Null }
        Write-Error "Failed to mount WIM image. Ensure DISM is available and the WIM is not corrupted.`n$jobError"
        exit 1
    }
    Remove-Job -Job $wimJob
    Write-Host "  ✓ WIM mounted" -ForegroundColor Green

    Write-Host "  Copying fonts from WIM..." -ForegroundColor Yellow
    $wimFontsPath = Join-Path $wimMountPath "Windows\Fonts"
    $copiedCount = 0
    foreach ($font in ($criticalFonts + $warningFonts)) {
        $src = Join-Path $wimFontsPath $font
        if (Test-Path $src) {
            Copy-Item $src -Destination $bundledFontsPath -Force
            $copiedCount++
        }
    }
    Write-Host "  ✓ Copied $copiedCount fonts" -ForegroundColor Green

    Dismount-WindowsImage -Path $wimMountPath -Discard | Out-Null
    Remove-Item $wimMountPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "  ✓ WIM dismounted" -ForegroundColor Green

    if ($isoMounted) {
        Dismount-DiskImage -ImagePath $isoFile | Out-Null
        Write-Host "  ✓ ISO dismounted" -ForegroundColor Green
    }

    $missingCritical = $criticalFonts | Where-Object { -not (Test-Path (Join-Path $bundledFontsPath $_)) }
    if ($missingCritical.Count -gt 0) {
        Write-Error "Critical fonts still missing after acquisition: $($missingCritical -join ', ')"
        exit 1
    }

    $missingWarning = $warningFonts | Where-Object { -not (Test-Path (Join-Path $bundledFontsPath $_)) }
    if ($missingWarning.Count -gt 0) {
        Write-Warning "  Non-critical fonts not found in media: $($missingWarning -join ', ')"
    }

    $fontCount = @(Get-ChildItem "$bundledFontsPath\*.ttf" -ErrorAction SilentlyContinue).Count
    Write-Host "  ✓ Acquisition complete — $fontCount fonts ready`n" -ForegroundColor Green
}

# === STEP 1: STOP SERVICES ===
Write-Host "[1/6] Stopping font services..." -ForegroundColor Yellow
$services = @("FontCache", "FontCache3.0.0.0")
$stoppedCount = 0

foreach ($svc in $services) {
    $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq "Running") {
        Stop-Service -Name $svc -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
        $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
        if ($service.Status -eq "Stopped") {
            Write-Host "  ✓ Stopped $svc" -ForegroundColor Green
            $stoppedCount++
        }
        else {
            Write-Warning "  Could not stop $svc (may be locked)"
        }
    }
}

if ($stoppedCount -eq 0) {
    Write-Host "  All services already stopped" -ForegroundColor Gray
}

Start-Sleep -Seconds 2

# === STEP 2: CLEAR FONT CACHE ===
Write-Host "`n[2/6] Clearing font cache..." -ForegroundColor Yellow
$cachePaths = @(
    "$env:WINDIR\ServiceProfiles\LocalService\AppData\Local\FontCache",
    "$env:WINDIR\ServiceProfiles\LocalService\AppData\Local\FontCache3.0.0.0",
    "$env:WINDIR\System32\FNTCACHE.DAT",
    "$env:LOCALAPPDATA\Microsoft\Windows\Caches"
)

$clearedCount = 0
foreach ($cachePath in $cachePaths) {
    if (Test-Path $cachePath) {
        $isContainer = (Get-Item $cachePath -ErrorAction SilentlyContinue).PSIsContainer
        if ($isContainer) {
            $itemCount = @(Get-ChildItem $cachePath -Recurse -ErrorAction SilentlyContinue).Count
            Remove-Item "$cachePath\*" -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  ✓ Cleared $cachePath ($itemCount items)" -ForegroundColor Green
        }
        else {
            Remove-Item $cachePath -Force -ErrorAction SilentlyContinue
            Write-Host "  ✓ Deleted $cachePath" -ForegroundColor Green
        }
        $clearedCount++
    }
}

Write-Host "  Summary: $clearedCount cache locations cleared" -ForegroundColor Cyan

# === STEP 3: TAKE OWNERSHIP OF FONTS FOLDER ===
Write-Host "`n[3/6] Taking ownership of Fonts folder..." -ForegroundColor Yellow
takeown /f "$systemFontsPath\*.ttf" /a 2>&1 | Out-Null
icacls "$systemFontsPath\*.ttf" /grant "Administrators:F" /t /c /q 2>&1 | Out-Null
Write-Host "  ✓ Ownership acquired" -ForegroundColor Green

# === STEP 4: INSTALL BUNDLED FONTS ===
Write-Host "`n[4/6] Installing bundled fonts..." -ForegroundColor Yellow

$requiredFonts = @(
    "arial.ttf", "arialbd.ttf", "arialbi.ttf", "ariali.ttf",
    "times.ttf", "timesbd.ttf", "timesbi.ttf", "timesi.ttf",
    "cour.ttf", "courbd.ttf", "courbi.ttf", "couri.ttf",
    "verdana.ttf", "verdanab.ttf", "verdanai.ttf", "verdanaz.ttf",
    "tahoma.ttf", "tahomabd.ttf",
    "georgia.ttf", "georgiab.ttf", "georgiai.ttf", "georgiaz.ttf",
    "calibri.ttf", "calibrib.ttf", "calibrii.ttf", "calibriz.ttf",
    "segoeui.ttf", "segoeuib.ttf", "segoeuii.ttf", "seguisym.ttf"
)

$installedCount = 0
$failedFonts = @()

foreach ($fontFile in $requiredFonts) {
    $sourcePath = Join-Path $bundledFontsPath $fontFile
    $destPath = Join-Path $systemFontsPath $fontFile
    
    if (-not (Test-Path $sourcePath)) {
        Write-Warning "  ⚠ Missing from bundle: $fontFile"
        $failedFonts += $fontFile
        continue
    }
    
    # Force overwrite with robocopy for locked files
    $sourceDir = Split-Path $sourcePath
    $destDir = Split-Path $destPath
    robocopy "$sourceDir" "$destDir" "$fontFile" /IS /IT /COPY:DAT /R:2 /W:1 /NP /NDL /NJH /NJS | Out-Null
    
    if (Test-Path $destPath) {
        $sourceHash = (Get-FileHash $sourcePath -Algorithm MD5).Hash
        $destHash = (Get-FileHash $destPath -Algorithm MD5).Hash
        if ($sourceHash -eq $destHash) {
            Write-Host "  ✓ $fontFile" -ForegroundColor Green
            $installedCount++
        }
        else {
            Write-Warning "  ✗ $fontFile (copy failed — hash mismatch after install)"
            $failedFonts += $fontFile
        }
    }
    else {
        Write-Warning "  ✗ Failed to install: $fontFile"
        $failedFonts += $fontFile
    }
}

Write-Host "`n  Summary: $installedCount installed, $($failedFonts.Count) failed" -ForegroundColor Cyan
if ($failedFonts.Count -gt 0) {
    Write-Warning "  Failed fonts: $($failedFonts -join ', ')"
}

# === STEP 5: RESTORE FONT REGISTRY ===
Write-Host "`n[5/6] Restoring font registry..." -ForegroundColor Yellow

$regContent = @"
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts]
"Arial (TrueType)"="arial.ttf"
"Arial Bold (TrueType)"="arialbd.ttf"
"Arial Bold Italic (TrueType)"="arialbi.ttf"
"Arial Italic (TrueType)"="ariali.ttf"
"Times New Roman (TrueType)"="times.ttf"
"Times New Roman Bold (TrueType)"="timesbd.ttf"
"Times New Roman Bold Italic (TrueType)"="timesbi.ttf"
"Times New Roman Italic (TrueType)"="timesi.ttf"
"Courier New (TrueType)"="cour.ttf"
"Courier New Bold (TrueType)"="courbd.ttf"
"Courier New Bold Italic (TrueType)"="courbi.ttf"
"Courier New Italic (TrueType)"="couri.ttf"
"Verdana (TrueType)"="verdana.ttf"
"Verdana Bold (TrueType)"="verdanab.ttf"
"Verdana Italic (TrueType)"="verdanai.ttf"
"Verdana Bold Italic (TrueType)"="verdanaz.ttf"
"Tahoma (TrueType)"="tahoma.ttf"
"Tahoma Bold (TrueType)"="tahomabd.ttf"
"Georgia (TrueType)"="georgia.ttf"
"Georgia Bold (TrueType)"="georgiab.ttf"
"Georgia Italic (TrueType)"="georgiai.ttf"
"Georgia Bold Italic (TrueType)"="georgiaz.ttf"
"Calibri (TrueType)"="calibri.ttf"
"Calibri Bold (TrueType)"="calibrib.ttf"
"Calibri Italic (TrueType)"="calibrii.ttf"
"Calibri Bold Italic (TrueType)"="calibriz.ttf"
"Segoe UI (TrueType)"="segoeui.ttf"
"Segoe UI Bold (TrueType)"="segoeuib.ttf"
"Segoe UI Italic (TrueType)"="segoeuii.ttf"
"Segoe UI Symbol (TrueType)"="seguisym.ttf"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\FontSubstitutes]
"Arial"=-
"Arial Bold"=-
"Times New Roman"=-
"Times New Roman Bold"=-
"Courier New"=-
"Verdana"=-
"Tahoma"=-
"Georgia"=-
"Calibri"=-
"Segoe UI"=-
"@

$tempReg = "$env:TEMP\RestoreFonts_$(Get-Date -Format 'yyyyMMddHHmmss').reg"
$regContent | Out-File -FilePath $tempReg -Encoding ASCII -Force
$process = Start-Process -FilePath "reg.exe" -ArgumentList "import `"$tempReg`"" -Wait -PassThru -NoNewWindow

if ($process.ExitCode -eq 0) {
    Write-Host "  ✓ Registry restored successfully" -ForegroundColor Green
}
else {
    Write-Warning "  Registry import returned exit code $($process.ExitCode) (may be non-critical)"
}

Remove-Item $tempReg -Force -ErrorAction SilentlyContinue

# === STEP 6: RESTART SERVICES ===
Write-Host "`n[6/6] Restarting font services..." -ForegroundColor Yellow
foreach ($svc in $services) {
    Start-Service -Name $svc -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 500
    $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq "Running") {
        Write-Host "  ✓ Started $svc" -ForegroundColor Green
    }
    else {
        Write-Warning "  Could not start $svc (will start automatically on reboot)"
    }
}

# === COMPLETION SUMMARY ===
Write-Host "`n=======================================`n       Restoration Complete ✓`n=======================================" -ForegroundColor Cyan

Write-Host "`n✓ Installed/verified $installedCount fonts" -ForegroundColor White
Write-Host "✓ Cleared $clearedCount cache locations" -ForegroundColor White
Write-Host "✓ Restored font registry" -ForegroundColor White
Write-Host "✓ Services restarted" -ForegroundColor White

Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Yellow
Write-Host "1. RESTART your computer NOW (required for changes to take effect)" -ForegroundColor White
Write-Host "2. After reboot, test Adobe Acrobat with problematic PDFs" -ForegroundColor White
Write-Host "3. The 'Arial-BoldMT' and 'Times New Roman' errors should be gone" -ForegroundColor White

Write-Host "`nVerify fonts after reboot:" -ForegroundColor Gray
Write-Host "  Get-ChildItem $systemFontsPath\arial*.ttf | Select Name`n" -ForegroundColor Gray

$reboot = Read-Host "Restart now? (Y/N)"
if ($reboot -eq 'Y' -or $reboot -eq 'y') {
    Write-Host "`nRestarting in 10 seconds... (Press Ctrl+C to cancel)" -ForegroundColor Yellow
    for ($i = 10; $i -gt 0; $i--) {
        Write-Host "  $i..." -NoNewline
        Start-Sleep -Seconds 1
    }
    Write-Host "`n`nRestarting now..." -ForegroundColor Red
    Restart-Computer -Force
}
else {
    Write-Warning "`nRemember to restart manually. Changes will not take full effect until you reboot."
}
