# Dell Battery Manager

Windows app to view and change the battery settings stored in a Dell laptop's BIOS, in the spirit of the former Dell Power Manager.

A single PowerShell + WPF script. Nothing to install.

## Features

- **Battery**: charge level, status, power, voltage, current and design capacity, health, cycle count.
- **Charge mode**: Standard, ExpressCharge, Primarily AC use, Adaptive, Custom (start and stop thresholds).
- **Advanced Charge**: on/off and per-day schedule.
- **Peak Shift**: on/off, battery threshold and per-day schedule.
- **English** and **French** interface.

## Requirements

- A Dell laptop that exposes the Dell BIOS WMI provider (`root\dcim\sysman`).
- Windows 10 or 11 (Windows PowerShell 5.1 is built in).
- Administrator rights.
- [Dell Command | Configure](https://www.dell.com/support/kbdoc/000178000/dell-command-configure), only for the Advanced Charge and Peak Shift schedules. Without it, everything else works and a banner offers to install it.

Tested on a Dell XPS 17 9730.

## Usage

1. Download `DellBatteryManager.ps1` and `DellBatteryManager.cmd` into the same folder.
2. Double-click `DellBatteryManager.cmd` and accept the administrator prompt.

If Windows blocks the downloaded script:

```powershell
Unblock-File .\DellBatteryManager.ps1
```

### Parameters

```powershell
powershell -STA -File DellBatteryManager.ps1 [-Demo] [-Language en|fr] [-ShowConsole]
```

| Parameter | Effect |
|---|---|
| `-Demo` | Simulated settings, no administrator rights, nothing written to the BIOS. |
| `-Language en\|fr` | Forces the language for this run. |
| `-ShowConsole` | Keeps the console visible (troubleshooting). |

The language can also be changed from the **Language** menu (English by default). The choice is remembered.

## How it works

- Charge mode, thresholds, on/off switches and Peak Shift threshold: Dell BIOS WMI, instant.
- Schedules: `cctk.exe` from Dell Command | Configure, about 20 s per call, read only when their tab is opened.
- Before each write: validation, a confirmation listing old → new values, then the BIOS is read back.
- BIOS administrator password: asked at the bottom of the window if one is set, never saved.

Local files, in `%LOCALAPPDATA%\DellBatteryManager\`:

| File | Content |
|---|---|
| `settings.json` | Chosen language. |
| `plannings.json` | Last known schedules (the BIOS does not return the schedule of a disabled mode). |
| `erreurs.log` | Error log. |

## Limitations

- Settings are written to the BIOS. Use at your own risk.
- Available modes and ranges depend on the BIOS; other models may behave differently.
- Dell Optimizer can change some settings: **Refresh** (F5) reads the BIOS again.
- If a BIOS password is set, it is passed to `cctk` on the command line while a schedule is written, so other administrator processes can see it.
