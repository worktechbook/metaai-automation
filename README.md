# MailCx — Android ADB UI Automation Utility

Windows Batch + PowerShell automation utility that drives Android UI flows over USB ADB (`uiautomator`), fetches disposable Mail.cx inboxes and 6-digit OTP verification codes via the Mail.cx REST API, and records session outputs into a dependency-free `.xlsx` workbook built with .NET `System.IO.Compression.ZipFile`.

## Features
- Auto-discovers `adb.exe` across PATH, WinGet, and Android SDK locations
- Multi-user Android profile detection (`pm list users` / `am start-user`)
- XML UI hierarchy parsing via `uiautomator dump /dev/tty`
- Parallel Mail.cx inbox polling with automatic OTP extraction and inbox cleanup
- Native `.xlsx` generation without requiring Microsoft Excel

## Prerequisites
- Windows 10/11 with PowerShell 5.1+
- Android Platform-Tools (`adb.exe`) and a USB-debugging-enabled Android device

## Configuration
Set environment variables (see `.env.example`) or pass parameters directly:
- `ADB_SERIAL`: Optional target device serial (auto-detects if omitted)
- `METAAI_DEFAULT_PASSWORD`: Default password used if left blank at the prompt

## Usage
Run `MetaAI Phone Fast.cmd` or invoke `metaai_factory_fast.ps1` directly from PowerShell.
