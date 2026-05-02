@echo off
IF "%~1"=="" (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "Start-Process powershell.exe -ArgumentList '-ExecutionPolicy Bypass -NoProfile -File \"%~dp0Restore-WindowsFonts.ps1\"' -Verb RunAs"
) ELSE (
    powershell.exe -ExecutionPolicy Bypass -NoProfile -Command "Start-Process powershell.exe -ArgumentList '-ExecutionPolicy Bypass -NoProfile -File \"%~dp0Restore-WindowsFonts.ps1\" -IsoPath \"%~1\"' -Verb RunAs"
)
