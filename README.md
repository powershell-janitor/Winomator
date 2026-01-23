# Winomator

**Winomator** is a Windows third-party app update runner built around **winget**, designed to run safely as **SYSTEM** (RMM / scheduled task) while still providing **user-visible notifications** and **interactive “app is open” prompts**.

It’s built for real-world enterprise pain points:
- winget output parsing that won’t get confused by progress spinners / ANSI sequences
- avoids Microsoft Store source agreement prompts by forcing `--source winget`
- can prompt the logged-on user when an app is open (Proceed vs Snooze)
- “Snooze” writes per-app timers under `%ProgramData%\Winomator\Snooze`
- toast notifications via **BurntToast** executed in the interactive user context (scheduled task w/ `InteractiveToken`)

---

## Features

- **Weekly bulk update** using `winget update` + `winget upgrade`
- **Per-app exclusions**
  - by winget Id (`$ExcludedWingetIds`)
  - by app name regex patterns (`$ExcludedNamePatterns`)
- **Open app detection**
  - maps winget Ids to process names (`$AppProcessMap`)
  - prompts user to proceed (auto-close app) or snooze
- **Toast notifications**
  - runs as SYSTEM but displays toasts as the logged-on user
  - uses BurntToast if available/installed
- **Robust logging**
  - logs to: `%ProgramData%\Winomator\Logs\Winomator-Weekly.log`
  - captures winget stdout/stderr per app in `%ProgramData%\Winomator\`

---

## Requirements

- Windows 10/11
- PowerShell 5.1+ (works well with Windows PowerShell)
- winget available (App Installer)
  - Script includes logic to download/extract winget runtime for SYSTEM if needed
- For toast notifications:
  - BurntToast module (script can install it to `AllUsers`)

> Note: winget behavior under SYSTEM can vary depending on device configuration, proxy/DNS, and whether apps are per-user vs machine-wide installs.

---

## How it works (high level)

1. Script runs as **SYSTEM** (RMM agent / scheduled task).
2. Discovers upgradeable packages via `winget update --source winget`.
3. For each target:
   - checks snooze timer
   - checks if mapped process names are currently running
   - if running: prompts the logged-on user (Proceed/Snooze)
   - runs `winget upgrade --id <id> --exact --silent` and falls back to a non-silent retry in specific cases
4. Summarizes results and sends a completion toast.

---

## Installation / Usage

### Option A: Run manually (admin PowerShell)
```powershell
Set-ExecutionPolicy Bypass -Scope Process -Force
.\src\Winomator.ps1
