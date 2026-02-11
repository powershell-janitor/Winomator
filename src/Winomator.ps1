<#
.SYNOPSIS
  Winomator - Winget Bulk Updater + User Toasts + "App is open" Snooze/Proceed prompts.

.DESCRIPTION
  PowerShell-based third-party patching runner built around winget, designed to run as SYSTEM
  with user-visible notifications and interactive prompts.

.PARAMETER EnableToasts
  Enables/disables toast notifications (default: enabled).

.PARAMETER EnablePopups
  Enables/disables interactive "app is open" popups (default: enabled).

.PARAMETER WhatIf
  Runs in report-only mode (no upgrades performed).

.PARAMETER ExcludedWingetIds
  Winget package IDs to exclude from upgrades.

.PARAMETER ExcludedNamePatterns
  Regex patterns to exclude packages by Name.

.PARAMETER OrganizationName
  Name shown in popup text (default: "Your IT department").

.PARAMETER MaxConcurrentUpgrades
  Max parallel upgrades for non-interactive packages (default: 3).

.PARAMETER MaxLogSizeMB
  Max log file size in MB before rotation (default: 5).

.NOTES
  Author: Daniel Davies
  Repo:   https://github.com/DanielDavies-reveleer/Winomator

<#
Copyright (c) 2026 Daniel Davies - MIT License
#>

#>

[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [bool]$EnableToasts = $true,
    [bool]$EnablePopups = $true,
    [string[]]$ExcludedWingetIds = @("VMware.HorizonClient","Microsoft.Office","Microsoft.Edge","Dell.PeripheralManager","Cisco.Webex","Perimeter81.HarmonySASE","Microsoft.VisualStudio.2022.Professional","Microsoft.VisualStudio.2025.Professional"),
    [string[]]$ExcludedNamePatterns = @(),
    [string]$OrganizationName = "Reveleer IT",
    [int]$MaxConcurrentUpgrades = 3,
    [int]$MaxLogSizeMB = 5
)
$ErrorActionPreference = 'Stop'

$LogRoot     = "$env:ProgramData\Winomator\Logs"
$LogFile     = Join-Path $LogRoot "Winomator-Weekly.log"
$SummaryFile = Join-Path $LogRoot "Winomator-LastRun.json"
$RootDir     = Join-Path $env:ProgramData "Winomator"
$ToastDir    = Join-Path $RootDir "Toast"
$SnoozeDir   = Join-Path $RootDir "Snooze"
$LockFile    = Join-Path $RootDir "winomator.lock"

$PopupResponseTimeoutSeconds = 120
$DefaultSnoozeMinutes        = 60
$PopupSnoozeOptionsMinutes   = @(15, 30, 60, 120)

$AppProcessMap = @{

    # ── Browsers ──────────────────────────────────────────────────────────
    "Google.Chrome"                = @("chrome")
    "Google.Chrome.Beta"           = @("chrome")
    "Google.Chrome.Dev"            = @("chrome")
    "Google.Chrome.Canary"         = @("chrome")
    "Mozilla.Firefox"              = @("firefox")
    "Mozilla.Firefox.ESR"          = @("firefox")
    "Mozilla.Firefox.DeveloperEdition" = @("firefox")
    "BraveSoftware.BraveBrowser"   = @("brave")
    "Opera.Opera"                  = @("opera")
    "Opera.OperaGX"                = @("opera")
    "Vivaldi.Vivaldi"              = @("vivaldi")
    "ArcBrowser.Arc"               = @("Arc")
    # Note: Microsoft Edge is updated via Windows Update, not typically winget

    # ── IDEs & Code Editors ───────────────────────────────────────────────
    "Microsoft.VisualStudioCode"           = @("Code")
    "Microsoft.VisualStudioCode.Insiders"  = @("Code - Insiders")
    "Microsoft.VisualStudio.2022.Community"    = @("devenv")
    "Microsoft.VisualStudio.2022.Professional" = @("devenv")
    "Microsoft.VisualStudio.2022.Enterprise"   = @("devenv")
    "Microsoft.VisualStudio.2025.Community"    = @("devenv")
    "Microsoft.VisualStudio.2025.Professional" = @("devenv")
    "Microsoft.VisualStudio.2025.Enterprise"   = @("devenv")
    "JetBrains.IntelliJIDEA.Community"     = @("idea64","idea")
    "JetBrains.IntelliJIDEA.Ultimate"      = @("idea64","idea")
    "JetBrains.PyCharm.Community"          = @("pycharm64","pycharm")
    "JetBrains.PyCharm.Professional"       = @("pycharm64","pycharm")
    "JetBrains.WebStorm"                   = @("webstorm64","webstorm")
    "JetBrains.Rider"                      = @("rider64","rider")
    "JetBrains.GoLand"                     = @("goland64","goland")
    "JetBrains.CLion"                      = @("clion64","clion")
    "JetBrains.DataGrip"                   = @("datagrip64","datagrip")
    "JetBrains.PHPStorm"                   = @("phpstorm64","phpstorm")
    "JetBrains.RubyMine"                   = @("rubymine64","rubymine")
    "JetBrains.RustRover"                  = @("rustrover64")
    "JetBrains.Fleet"                      = @("Fleet")
    "JetBrains.Toolbox"                    = @("jetbrains-toolbox")
    "Notepad++.Notepad++"                  = @("notepad++")
    "SublimeHQ.SublimeText.4"             = @("sublime_text")
    "SublimeHQ.SublimeMerge"              = @("sublime_merge")
    "Cursor.Cursor"                        = @("Cursor")
    "Zed.Zed"                              = @("zed")

    # ── Communication & Collaboration ─────────────────────────────────────
    "Microsoft.Teams"              = @("ms-teams","Teams")
    "SlackTechnologies.Slack"      = @("slack")
    "Zoom.Zoom"                    = @("Zoom","ZoomWebHost")
    "Zoom.ZoomWorkplace"           = @("Zoom","ZoomWebHost")
    "Discord.Discord"              = @("Discord")
    "OpenWhisperSystems.Signal"    = @("Signal")
    "Telegram.TelegramDesktop"     = @("Telegram")
    "RingCentral.RingCentral"      = @("RingCentral")
    "Cisco.Webex"                  = @("webex","CiscoCollabHost","ptoneclk","AtmgNode","CiscoWebExStart")
    "Cisco.CiscoSecureClient"      = @("vpnui","vpnagent")
    "BlueJeans.BlueJeans"          = @("BlueJeans")
    "GoTo.Connect"                 = @("GoTo","GoToMeeting")
    "GoTo.Meeting"                 = @("GoToMeeting","g2mcomm")

    # ── Microsoft Office / 365 ────────────────────────────────────────────
    # Note: Office is usually updated via Click-to-Run, but in case winget picks it up
    "Microsoft.Office"             = @("WINWORD","EXCEL","POWERPNT","OUTLOOK","ONENOTE","MSACCESS","MSPUB","lync")

    # ── Cloud Storage & Sync ──────────────────────────────────────────────
    "Microsoft.OneDrive"           = @("OneDrive","OneDriveStandaloneUpdater")
    "Dropbox.Dropbox"              = @("Dropbox")
    "Google.GoogleDrive"           = @("GoogleDriveFS")
    "BoxDrive.BoxDrive"            = @("Box")

    # ── VPN & Remote Access ───────────────────────────────────────────────
    "Citrix.Workspace"             = @("SelfService","Receiver","wfcrun32","wfica32","AuthManagerSvc","concentr")
    "VMware.HorizonClient"         = @("vmware-view","vmware-horizon-client")
    "PaloAltoNetworks.GlobalProtect" = @("PanGPA","PanGPS")
    "OpenVPNTechnologies.OpenVPN"  = @("openvpn","openvpn-gui")
    "WireGuard.WireGuard"          = @("wireguard")
    "tailscale.tailscale"          = @("tailscaled","tailscale-ipn")
    "Fortinet.FortiClient"         = @("FortiClient","FortiTray","FortiSSLVPNdaemon")

    # ── Database & DB Tools ───────────────────────────────────────────────
    "dbeaver.dbeaver"              = @("dbeaver")
    "Microsoft.AzureDataStudio"    = @("azuredatastudio")
    "DBngin.DBngin"                = @("DBngin")
    "TablePlus.TablePlus"          = @("TablePlus")
    "HeidiSQL.HeidiSQL"            = @("heidisql")
    "MongoDB.Compass.Full"         = @("MongoDBCompass")
    "pgAdmin.pgAdmin"              = @("pgAdmin4")
    "PostgreSQL.pgAdmin"           = @("pgAdmin4")
    "RedisInsight.RedisInsight"    = @("RedisInsight")

    # ── DevOps / Containers / Terminal ────────────────────────────────────
    "Docker.DockerDesktop"         = @("Docker Desktop","com.docker.backend","com.docker.proxy")
    "Microsoft.WindowsTerminal"    = @("WindowsTerminal","wt")
    "Git.Git"                      = @("git","git-bash")
    "GitHub.GitHubDesktop"         = @("GitHubDesktop")
    "Atlassian.Sourcetree"         = @("SourceTree")
    "Axosoft.GitKraken"            = @("gitkraken")
    "Postman.Postman"              = @("Postman")
    "Insomnia.Insomnia"            = @("Insomnia")

    # ── PDF & Document Viewers ────────────────────────────────────────────
    "Adobe.Acrobat.Reader.32-bit"  = @("AcroRd32")
    "Adobe.Acrobat.Reader.64-bit"  = @("Acrobat","AcroCEF","RdrCEF")
    "Adobe.Acrobat.Pro"            = @("Acrobat","AcroCEF")
    "Foxit.FoxitReader"            = @("FoxitPDFReader")
    "Foxit.PhantomPDF"             = @("FoxitPhantomPDF")
    "TrackerSoftware.PDF-XChangeEditor" = @("PDFXEdit")
    "SumatraPDF.SumatraPDF"        = @("SumatraPDF")
    "Nitro.NitroPDF"               = @("NitroPDF","NitroPro")

    # ── Media Players ─────────────────────────────────────────────────────
    "VideoLAN.VLC"                 = @("vlc")
    "Spotify.Spotify"              = @("Spotify")
    "AIMP.AIMP"                    = @("AIMP")
    "clsid2.mpc-hc"               = @("mpc-hc64","mpc-hc")
    "OBSProject.OBSStudio"        = @("obs64","obs32")

    # ── Graphics & Design ─────────────────────────────────────────────────
    "GIMP.GIMP"                    = @("gimp-2.10","gimp")
    "Figma.Figma"                  = @("Figma")
    "Inkscape.Inkscape"            = @("inkscape")
    "BlenderFoundation.Blender"    = @("blender")
    "paint.net.paintdotnet"        = @("paintdotnet","PaintDotNet")
    "IrfanSkiljan.IrfanView"       = @("i_view64","i_view32")
    "ShareX.ShareX"                = @("ShareX")
    "Greenshot.Greenshot"          = @("Greenshot")
    "ScreenToGif.ScreenToGif"      = @("ScreenToGif")
    "Snagit.Snagit"                = @("Snagit32","SnagitEditor")

    # ── File Management & Utilities ───────────────────────────────────────
    "7zip.7zip"                    = @("7zFM","7zG")
    "RARLab.WinRAR"                = @("WinRAR")
    "voidtools.Everything"         = @("Everything")
    "WinSCP.WinSCP"                = @("WinSCP")
    "PuTTY.PuTTY"                  = @("putty","pageant","plink")
    "FileZilla.FileZilla"          = @("filezilla")
    "SFTPFileZilla.FileZilla"      = @("filezilla")
    "Ghisler.TotalCommander"       = @("TOTALCMD64","TOTALCMD")
    "Microsoft.PowerToys"          = @("PowerToys","PowerToys.Settings")

    # ── Password Managers ─────────────────────────────────────────────────
    "AgileBits.1Password"          = @("1Password")
    "Bitwarden.Bitwarden"          = @("Bitwarden")
    "KeePassXCTeam.KeePassXC"     = @("KeePassXC")
    "LastPass.LastPass"            = @("LastPass","lpwinmetro")
    "Dashlane.Dashlane"            = @("Dashlane")
    "DominikReichl.KeePass"        = @("KeePass")
    "Oracle.MySQLWorkbench"        = @("MySQLWorkbench")
    "Mattermost.MattermostDesktop" = @("Mattermost")
    "Perimeter81.HarmonySASE"      = @("Perimeter81","Perimeter81.Service")

    # ── Note-Taking & Knowledge ───────────────────────────────────────────
    "Notion.Notion"                = @("Notion")
    "Obsidian.Obsidian"            = @("Obsidian")
    "Logseq.Logseq"               = @("Logseq")
    "evernote.evernote"            = @("Evernote")

    # ── Email Clients ─────────────────────────────────────────────────────
    "Mozilla.Thunderbird"          = @("thunderbird")
    "eM-Client.eMClient"           = @("MailClient")

    # ── Virtualization ────────────────────────────────────────────────────
    "Oracle.VirtualBox"            = @("VirtualBox","VirtualBoxVM","VBoxSVC")
    "VMware.WorkstationPlayer"     = @("vmplayer","vmware")
    "VMware.WorkstationPro"        = @("vmware","vmplayer")

    # ── Security & Antivirus ──────────────────────────────────────────────
    "CrowdStrike.FalconSensor"     = @("CSFalconService","CSFalconContainer")
    "SentinelOne.SentinelOneAgent" = @("SentinelAgent","SentinelServiceHost")
    "Malwarebytes.Malwarebytes"    = @("mbam","MBAMService")
    "CarbonBlack.CarbonBlackCloud" = @("RepMgr","cb")

    # ── Monitoring & IT Tools ─────────────────────────────────────────────
    "Wireshark.Wireshark"          = @("Wireshark")
    "Insecure.Nmap"                = @("nmap","zenmap")
    "WinDirStat.WinDirStat"       = @("windirstat")
    "REALiX.HWiNFO"               = @("HWiNFO64","HWiNFO32")
    "CPUID.CPU-Z"                  = @("cpuz")

    # ── Node.js / Runtime (usually safe but can lock files) ───────────────
    "OpenJS.NodeJS"                = @("node")
    "OpenJS.NodeJS.LTS"            = @("node")
    "Python.Python.3.12"           = @("python","pythonw")
    "Python.Python.3.13"           = @("python","pythonw")

    # ── Miscellaneous Enterprise ──────────────────────────────────────────
    "Microsoft.PowerBI"            = @("PBIDesktop")
    "Tableau.Desktop"              = @("tableau")
    "Tableau.Reader"               = @("tabreader")
    "Salesforce.DataLoader"        = @("dataloader")
    "Amazon.WorkSpaces"            = @("workspaces")
    "Grammarly.Grammarly"         = @("Grammarly")
    "Loom.Loom"                    = @("Loom")
    "Krisp.Krisp"                  = @("krisp")
    "1Password.1PasswordCLI"       = @("op")
}

$ToastImageMap = @{
    Info    = (Join-Path $ToastDir "info.png")
    Success = (Join-Path $ToastDir "success.png")
    Warning = (Join-Path $ToastDir "warning.png")
    Error   = (Join-Path $ToastDir "error.png")
}
$ToastBuiltinCandidates = @{
    Info    = @("$env:windir\System32\SecurityAndMaintenance.png")
    Success = @("$env:windir\System32\SecurityAndMaintenance.png")
    Warning = @("$env:windir\System32\SecurityAndMaintenance_Alert.png")
    Error   = @("$env:windir\System32\SecurityAndMaintenance_Error.png")
}

$script:RunResults = [System.Collections.ArrayList]::new()

#region Concurrency Guard
$script:WinomatorMutex = $null

function Enter-WinomatorLock {
    $created = $false
    try {
        $script:WinomatorMutex = [System.Threading.Mutex]::new($false, "Global\WinomatorExclusiveRun", [ref]$created)
        if (-not $script:WinomatorMutex.WaitOne(0)) { return $false }
        New-Item -ItemType Directory -Path $RootDir -Force | Out-Null
        "$PID $(Get-Date -Format o)" | Set-Content -LiteralPath $LockFile -Encoding UTF8 -Force
        return $true
    } catch {
        return $false
    }
}

function Exit-WinomatorLock {
    try {
        if ($script:WinomatorMutex) {
            $script:WinomatorMutex.ReleaseMutex()
            $script:WinomatorMutex.Dispose()
            $script:WinomatorMutex = $null
        }
    } catch { }
    Remove-Item -LiteralPath $LockFile -Force -ErrorAction SilentlyContinue
}
#endregion

#region Logging
function Invoke-LogRotation {
    if (-not (Test-Path -LiteralPath $LogFile)) { return }
    try {
        $fi = Get-Item -LiteralPath $LogFile -ErrorAction SilentlyContinue
        if ($fi -and $fi.Length -gt ($MaxLogSizeMB * 1MB)) {
            $arc = Join-Path $LogRoot "Winomator-Weekly-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
            Move-Item -LiteralPath $LogFile -Destination $arc -Force
            Get-ChildItem -Path $LogRoot -Filter "Winomator-Weekly-*.log" -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -Skip 5 |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }
    } catch { }
}

function Write-Log {
    param([string]$Message, [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO")
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
    New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    Write-Output $line
}
#endregion

#region Atomic File Write
function Set-ContentAtomic {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Content,
        [string]$Encoding = "UTF8"
    )
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $tmp = "$Path.tmp.$PID"
    try {
        Set-Content -LiteralPath $tmp -Value $Content -Encoding $Encoding -Force
        Move-Item -LiteralPath $tmp -Destination $Path -Force
    } catch {
        Set-Content -LiteralPath $Path -Value $Content -Encoding $Encoding -Force
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}
#endregion

#region Result Tracking
function Add-RunResult {
    param(
        [string]$Name, [string]$Id, [string]$VersionBefore, [string]$VersionAfter,
        [string]$VersionAvailable,
        [ValidateSet("Success","Failed","Skipped","Snoozed","WhatIf")][string]$Outcome,
        [string]$Detail = "", [double]$DurationSeconds = 0
    )
    [void]$script:RunResults.Add([pscustomobject]@{
        Name             = $Name
        Id               = $Id
        VersionBefore    = $VersionBefore
        VersionAfter     = $VersionAfter
        VersionAvailable = $VersionAvailable
        Outcome          = $Outcome
        Detail           = $Detail
        DurationSeconds  = [math]::Round($DurationSeconds, 1)
        Timestamp        = (Get-Date).ToString("o")
    })
}

function Export-RunSummary {
    param([int]$SuccessCount, [int]$FailedCount, [int]$TotalAvailable, [bool]$RebootPending)
    $summary = [pscustomobject]@{
        RunTime        = (Get-Date).ToString("o")
        TotalAvailable = $TotalAvailable
        SuccessCount   = $SuccessCount
        FailedCount    = $FailedCount
        RebootPending  = $RebootPending
        WhatIfMode     = [bool]$WhatIfPreference
        Results        = @($script:RunResults)
    }
    Set-ContentAtomic -Path $SummaryFile -Content ($summary | ConvertTo-Json -Depth 6)
    Write-Log "Summary report written to: $SummaryFile"
}
#endregion

#region Reboot Detection
function Test-RebootPending {
    $keys = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
    )
    foreach ($k in $keys) {
        if (Test-Path -LiteralPath $k) { return $true }
    }
    try {
        $v = Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" `
            -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
        if ($v -and $v.PendingFileRenameOperations) { return $true }
    } catch { }
    return $false
}
#endregion

#region Text Helpers
function Remove-Ansi {
    param([string]$Text)
    if ($null -eq $Text) { return $Text }
    return ($Text -replace "`e\[[0-?]*[ -/]*[@-~]", "")
}

function Remove-ProgressJunk {
    param([string]$Text)
    if ($null -eq $Text) { return $Text }
    $t = $Text -replace '[\u2580-\u259F]', '' -replace '(?:â–.)+', ''
    if ($t.Trim() -match '^[-\\|/\s]+$') { return '' }
    return $t
}

function Convert-ToSafeFileToken {
    param([Parameter(Mandatory)][string]$Text)
    return ($Text -replace '[^A-Za-z0-9\.\-_]', '_')
}

function Get-MergedLogPreview {
    param(
        [Parameter(Mandatory)][string]$StdOutFile,
        [Parameter(Mandatory)][string]$StdErrFile,
        [int]$MaxLines = 14
    )
    $lines = @()
    if (Test-Path $StdOutFile) {
        $lines += Get-Content -Path $StdOutFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    if (Test-Path $StdErrFile) {
        $lines += Get-Content -Path $StdErrFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    $lines = $lines | ForEach-Object { Remove-ProgressJunk (Remove-Ansi "$_") } |
        Where-Object { $_ -and $_.Trim() }
    return (($lines | Select-Object -First $MaxLines) -join " | ")
}

function Get-CleanedWingetOutput {
    param(
        [Parameter(Mandatory)][string]$StdOutFile,
        [Parameter(Mandatory)][string]$StdErrFile
    )
    $lines = @()
    if (Test-Path $StdOutFile) {
        $lines += Get-Content -Path $StdOutFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    if (Test-Path $StdErrFile) {
        $lines += Get-Content -Path $StdErrFile -Encoding UTF8 -ErrorAction SilentlyContinue
    }
    $cleaned = $lines | ForEach-Object {
        $t = Remove-Ansi "$_"
        $t = $t -replace '[\u2580-\u259F]', ''
        $t = $t -replace '(?:â–.)+', ''
        if ($t.Trim() -match '^[-\\|/\s]+$') { return }
        if ($t.Trim()) { $t.Trim() }
    } | Where-Object { $_ }
    return ($cleaned -join " | ")
}
#endregion


#region Truncation-Aware ID Matching

function Test-WingetIdMatch {
    param(
        [Parameter(Mandatory)][string]$ParsedId,
        [Parameter(Mandatory)][string]$KnownId
    )
    $ellipsis = [string][char]0x2026
    if ($ParsedId -eq $KnownId) { return $true }
    if ($ParsedId.EndsWith($ellipsis)) {
        $prefix = $ParsedId.TrimEnd([char]0x2026)
        if ($KnownId.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    if ($KnownId.EndsWith($ellipsis)) {
        $prefix = $KnownId.TrimEnd([char]0x2026)
        if ($ParsedId.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Resolve-WingetId {
    <#
    .SYNOPSIS
      Given a (possibly truncated) parsed ID and a list of known full IDs,
      returns the matching full ID or the original parsed ID if no match found.
    .DESCRIPTION
      Useful for resolving a truncated ID back to its full form before using
      it in file paths, cooldown records, or logging.
    #>
    param(
        [Parameter(Mandatory)][string]$ParsedId,
        [Parameter(Mandatory)][string[]]$KnownIds
    )
    foreach ($k in $KnownIds) {
        if (Test-WingetIdMatch -ParsedId $ParsedId -KnownId $k) { return $k }
    }
    return $ParsedId
}

#endregion


#region Winget Discovery
function Test-WingetOutputIndicatesSuccess {
    param([Parameter(Mandatory)][string]$T)
    if ($T -match '(?i)\bNo installed package found\b') { return $false }
    return (
        $T -match '(?i)\bSuccessfully installed\b' -or
        $T -match '(?i)\bNo available upgrade found\b' -or
        $T -match '(?i)\bNo upgrade available\b' -or
        $T -match '(?i)\bNo newer package versions\b' -or
        $T -match '(?i)\bNo applicable update found\b' -or
        $T -match '(?i)\bNo updates available\b'
    )
}

function Test-WingetOutputIndicatesHardFailure {
    param([Parameter(Mandatory)][string]$T)
    return (
        $T -match '(?i)\bInstaller failed with exit code:\s*\d+\b' -or
        $T -match '(?i)\bUninstall failed with exit code:\s*\d+\b' -or
        $T -match '(?i)\bInstall failed with exit code:\s*\d+\b' -or
        $T -match '(?i)\bError:\b'
    )
}

function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $m = Get-Item "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe" `
        -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($m) { return $m.FullName }
    return $null
}

function Get-WingetEmbeddedExitCode {
    param([string]$Text)
    if ($Text -match '(?i)\b(?:Installer|Uninstall|Install) failed with exit code:\s*(\d+)\b') {
        return [int]$matches[1]
    }
    return $null
}

function Get-WingetNoRetryReason {
    param([Parameter(Mandatory)][string]$Text)
    if ($Text -match '(?i)\bNo installed package found\b') {
        return "Not installed / not visible to SYSTEM"
    }
    if ($Text -match '(?i)\bNo package found matching input criteria\b') {
        return "Not found in configured sources"
    }
    if ($Text -match '(?i)\bAnother installation is already in progress\b' -or $Text -match '(?i)\b1618\b') {
        return "Windows Installer busy (1618)"
    }
    if ($Text -match '(?i)\bAccess is denied\b|\b0x80070005\b') {
        return "Access denied (0x80070005)"
    }
    if ($Text -match '(?i)\b0x80072EE7\b|\b0x80072EFD\b|\b0x80072F8F\b') {
        return "Network/DNS/TLS failure"
    }
    return $null
}

function Invoke-WingetToFiles {
    param(
        [Parameter(Mandatory)][string]$WingetPath,
        [Parameter(Mandatory)][string[]]$Arguments,
        [Parameter(Mandatory)][string]$StdOutFile,
        [Parameter(Mandatory)][string]$StdErrFile,
        [int]$TimeoutSeconds = 3600
    )
    New-Item -ItemType Directory -Path (Split-Path $StdOutFile) -Force | Out-Null
    Remove-Item $StdOutFile -Force -ErrorAction SilentlyContinue
    Remove-Item $StdErrFile -Force -ErrorAction SilentlyContinue

    $argStr = ($Arguments | ForEach-Object {
        if ($_ -match '\s') { '"' + $_.Replace('"','\"') + '"' } else { $_ }
    }) -join ' '

    try {
        $p = Start-Process -FilePath $WingetPath -ArgumentList $argStr -NoNewWindow -PassThru `
            -RedirectStandardOutput $StdOutFile -RedirectStandardError $StdErrFile
    } catch {
        Write-Log "Start-Process failed: $($_.Exception.Message)" "ERROR"
        return 9998
    }

    $exited = $p.WaitForExit($TimeoutSeconds * 1000)
    if (-not $exited) {
        try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch { }
        return 1460
    }

    try { $p.Refresh() } catch { }
    $ec = $p.ExitCode

    if ($null -eq $ec) {
        $merged = Get-CleanedWingetOutput -StdOutFile $StdOutFile -StdErrFile $StdErrFile
        if ($merged -match '(?i)\bSuccessfully installed\b') { return 0 }
        if ($merged -match '(?i)\bNo available upgrade found\b|No newer package versions\b|No applicable update found\b|No updates available\b') {
            return 0
        }
        return 9997
    }
    return [int]$ec
}

function Get-InstalledWingetVersion {
    param(
        [Parameter(Mandatory)][string]$WingetPath,
        [Parameter(Mandatory)][string]$Id
    )
    $tok = Convert-ToSafeFileToken $Id
    $tmpOut = Join-Path $RootDir "list-$tok.out.txt"
    $tmpErr = Join-Path $RootDir "list-$tok.err.txt"
    $null = Invoke-WingetToFiles -WingetPath $WingetPath -Arguments @(
        "list","--id",$Id,"--exact","--source","winget",
        "--accept-source-agreements","--disable-interactivity"
    ) -StdOutFile $tmpOut -StdErrFile $tmpErr -TimeoutSeconds 120
    if (-not (Test-Path $tmpOut)) { return $null }
    $lines = Get-Content $tmpOut -Encoding UTF8 -ErrorAction SilentlyContinue |
        ForEach-Object { Remove-Ansi "$_" } | Where-Object { $_ -and $_.Trim() }
    $hit = $lines | Where-Object { $_ -match [regex]::Escape($Id) } | Select-Object -First 1
    if (-not $hit) { return $null }
    $ver = ([regex]::Matches($hit, '\d+(\.\d+){1,3}') | Select-Object -First 1).Value
    if ([string]::IsNullOrWhiteSpace($ver)) { return $null }
    return $ver
}

function Test-WingetAsSystem {
    param([Parameter(Mandatory)][string]$WingetPath)
    $out = ""; $err = ""; $ec = $null
    try {
        $out = & $WingetPath --version 2>$null
        $ec = $LASTEXITCODE
    } catch {
        $err = $_.Exception.Message
        $ec = 9999
    }
    if ($null -eq $out) { $out = "" } else { $out = ([string]$out).Trim() }
    if ($null -eq $err) { $err = "" } else { $err = ([string]$err).Trim() }
    $hv = ($out -match '^\s*\d+(\.\d+){1,3}\s*$')
    $ob = 0; $eb = 0
    try {
        if ($out) { $ob = [Text.Encoding]::UTF8.GetByteCount($out) }
        if ($err) { $eb = [Text.Encoding]::UTF8.GetByteCount($err) }
    } catch { }
    return [pscustomobject]@{
        ExitCode   = $ec
        StdOut     = $out
        StdErr     = $err
        HasVersion = $hv
        Ok         = ($hv -or ($ec -eq 0))
        OutBytes   = $ob
        ErrBytes   = $eb
    }
}

function Initialize-WingetForSystem {
    param([string]$CacheRoot = "$env:ProgramData\Winomator\WingetRuntime")

    function Expand-ZipLike {
        param(
            [Parameter(Mandatory)][string]$Path,
            [Parameter(Mandatory)][string]$Destination
        )
        if (-not (Test-Path -LiteralPath $Path)) { throw "Archive not found: $Path" }
        if (Test-Path -LiteralPath $Destination) {
            Remove-Item -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue
        }
        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
        $tar = Join-Path $env:SystemRoot "System32\tar.exe"
        if (-not (Test-Path -LiteralPath $tar)) { throw "tar.exe not found." }
        $tarOut = & $tar -xf $Path -C $Destination 2>&1
        if ($LASTEXITCODE -ne 0) {
            $msg = ($tarOut | Select-Object -First 12) -join " | "
            throw "Extraction failed (tar exit $LASTEXITCODE): $msg"
        }
    }

    $installed = Get-WingetPath
    if ($installed) {
        $t = Test-WingetAsSystem -WingetPath $installed
        if ($t.Ok) { return $installed }
        if ($t.ExitCode -ne -1073741515) { return $installed }
    }

    New-Item -ItemType Directory -Path $CacheRoot -Force | Out-Null
    $bundle = Join-Path $CacheRoot "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
    Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile $bundle -UseBasicParsing

    $fi = Get-Item -LiteralPath $bundle -ErrorAction SilentlyContinue
    if (-not $fi -or $fi.Length -lt 2000000) {
        throw "Downloaded winget bundle too small. SizeBytes=$($fi.Length)"
    }

    $sig = Get-AuthenticodeSignature -FilePath $bundle -ErrorAction SilentlyContinue
    if ($sig -and $sig.Status -ne 'Valid') {
        Write-Log "Winget bundle signature invalid: Status=$($sig.Status)" "WARN"
    }

    $bundleDir = Join-Path $CacheRoot "bundle"
    Expand-ZipLike -Path $bundle -Destination $bundleDir

    $payload = Get-ChildItem $bundleDir -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '(?i)x64' -and $_.Extension -match '(?i)\.(msix|appx)$' } |
        Sort-Object Length -Descending | Select-Object -First 1
    if (-not $payload) {
        $payload = Get-ChildItem $bundleDir -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '(?i)\.(msix|appx)$' } |
            Sort-Object Length -Descending | Select-Object -First 1
    }
    if (-not $payload) { throw "No .msix/.appx found in msixbundle." }

    $msixDir = Join-Path $CacheRoot "msix"
    Expand-ZipLike -Path $payload.FullName -Destination $msixDir

    $winget = Get-ChildItem $msixDir -Recurse -Filter "winget.exe" -File -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName
    if (-not $winget) { throw "winget.exe not found in extracted payload." }

    $sigExe = Get-AuthenticodeSignature -FilePath $winget -ErrorAction SilentlyContinue
    if ($sigExe -and $sigExe.Status -ne 'Valid') {
        Write-Log "Extracted winget.exe signature invalid: Status=$($sigExe.Status)" "WARN"
    }

    $t2 = Test-WingetAsSystem -WingetPath $winget
    if (-not $t2.Ok) {
        throw "Extracted winget not runnable. ExitCode=$($t2.ExitCode)"
    }
    return $winget
}
#endregion

#region Logged-on User
function Get-LoggedOnUser {
    try {
        $u = (Get-CimInstance Win32_ComputerSystem -ErrorAction Stop).UserName
        if ([string]::IsNullOrWhiteSpace($u)) { return $null }
        return $u
    } catch {
        return $null
    }
}
#endregion

#region BurntToast
function Install-BurntToastModule {
    try {
        if (Get-Module -ListAvailable -Name BurntToast -ErrorAction SilentlyContinue) { return $true }
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -Force -Scope AllUsers | Out-Null
        }
        try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch { }
        Install-Module -Name BurntToast -Scope AllUsers -Force -AllowClobber -ErrorAction Stop | Out-Null
        return $true
    } catch {
        Write-Log "BurntToast install failed: $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Resolve-ToastImagePath {
    param(
        [ValidateSet("Info","Success","Warning","Error")][string]$Type = "Info",
        [string]$ExplicitPath
    )
    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath) -and (Test-Path -LiteralPath $ExplicitPath)) {
        return $ExplicitPath
    }
    if ($ToastImageMap.ContainsKey($Type)) {
        $p = [string]$ToastImageMap[$Type]
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    if ($ToastBuiltinCandidates.ContainsKey($Type)) {
        foreach ($c in $ToastBuiltinCandidates[$Type]) {
            if ($c -and (Test-Path -LiteralPath $c)) { return $c }
        }
    }
    return $null
}
#endregion

#region Scheduled Tasks
function Import-ScheduledTasksModule {
    try {
        Import-Module ScheduledTasks -ErrorAction Stop
        return $true
    } catch {
        Write-Log "Failed to import ScheduledTasks: $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Register-InteractiveUserTask {
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][string]$UserName,
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$ArgName,
        [Parameter(Mandatory)][string]$ArgValue
    )
    if (-not (Import-ScheduledTasksModule)) { return $false }
    $ch = Join-Path $env:SystemRoot "System32\conhost.exe"
    $ps = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path -LiteralPath $ch) -or
        -not (Test-Path -LiteralPath $ps) -or
        -not (Test-Path -LiteralPath $ScriptPath)) {
        Write-Log "Missing binaries/scripts for task '$TaskName'" "WARN"
        return $false
    }
    $taskArgs = "--headless `"$ps`" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" $ArgName `"$ArgValue`""
    try {
        $action    = New-ScheduledTaskAction -Execute $ch -Argument $taskArgs
        $principal = New-ScheduledTaskPrincipal -UserId $UserName -LogonType Interactive -RunLevel Limited
        $trigger   = New-ScheduledTaskTrigger -Once -At ((Get-Date).Date.AddYears(10).AddHours(23).AddMinutes(59))
        $settings  = New-ScheduledTaskSettingsSet -Compatibility Win8 `
            -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 10)
        $task = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger -Settings $settings
        Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
        return $true
    } catch {
        Write-Log "Register-ScheduledTask '$TaskName' failed: $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Start-UserTask {
    param([Parameter(Mandatory)][string]$TaskName)
    if (-not (Import-ScheduledTasksModule)) { return 9999 }
    try {
        $null = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    } catch {
        Write-Log "Task '$TaskName' not found." "WARN"
        return 2
    }
    try {
        Start-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        return 0
    } catch {
        Write-Log "Start-ScheduledTask '$TaskName' failed: $($_.Exception.Message)" "WARN"
        return 3
    }
}

function Unregister-WinomatorTasks {
    if (-not (Import-ScheduledTasksModule)) { return }
    foreach ($n in @("WinomatorToast","WinomatorPopup")) {
        try {
            if (Get-ScheduledTask -TaskName $n -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $n -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log "Unregistered task: $n"
            }
        } catch {
            Write-Log "Failed to unregister '$n': $($_.Exception.Message)" "WARN"
        }
    }
}
#endregion

#region Toast Script
function Initialize-ToastScript {
    New-Item -ItemType Directory -Path $ToastDir -Force | Out-Null
    $sp = Join-Path $ToastDir "Show-Toast.ps1"
    $scriptContent = @'
param([Parameter(Mandatory)][string]$StatusFile)
$ErrorActionPreference = 'Stop'
$diagRan = Join-Path (Split-Path -Parent $StatusFile) 'toast-ran.txt'
$diagErr = Join-Path (Split-Path -Parent $StatusFile) 'toast-error.txt'
try {
    "$(Get-Date -Format o) Task ran as: $(whoami)" | Out-File $diagRan -Encoding UTF8 -Append
    if (-not (Test-Path -LiteralPath $StatusFile)) {
        "Missing: $StatusFile" | Out-File $diagErr -Encoding UTF8 -Append
        exit 0
    }
    $j = Get-Content -LiteralPath $StatusFile -Raw | ConvertFrom-Json
    $title = if ($j.Title) { [string]$j.Title } else { "Software Update" }
    $body  = if ($j.Body)  { [string]$j.Body  } else { "Update status." }
    $img   = [string]$j.ImagePath
    Import-Module BurntToast -ErrorAction Stop
    if ($img -and (Test-Path -LiteralPath $img)) {
        New-BurntToastNotification -Text $title, $body -AppLogo $img | Out-Null
    } else {
        New-BurntToastNotification -Text $title, $body | Out-Null
    }
} catch {
    "$(Get-Date -Format o) ERROR: $($_.Exception.Message)" | Out-File $diagErr -Encoding UTF8 -Append
    exit 0
}
'@
    Set-Content -Path $sp -Value $scriptContent -Encoding UTF8 -Force
    return $sp
}

function Send-WinomatorToast {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Body,
        [ValidateSet("Info","Success","Warning","Error")][string]$Type = "Info",
        [string]$ImagePath
    )
    if (-not $EnableToasts) { return }
    $user = Get-LoggedOnUser
    if (-not $user) { Write-Log "No interactive user; skipping toast: $Title"; return }
    if (-not (Install-BurntToastModule)) { Write-Log "BurntToast unavailable" "WARN"; return }

    $sp = Initialize-ToastScript
    $sf = Join-Path $ToastDir "status.json"
    if (-not (Register-InteractiveUserTask -TaskName "WinomatorToast" -UserName $user `
        -ScriptPath $sp -ArgName "-StatusFile" -ArgValue $sf)) {
        Write-Log "Toast task unavailable" "WARN"
        return
    }

    $resolved = Resolve-ToastImagePath -Type $Type -ExplicitPath $ImagePath
    $payload = [pscustomobject]@{
        Title     = $Title
        Body      = $Body
        Type      = $Type
        ImagePath = $resolved
        Time      = (Get-Date).ToString("o")
    }
    Set-ContentAtomic -Path $sf -Content ($payload | ConvertTo-Json -Depth 4)
    $ec = Start-UserTask -TaskName "WinomatorToast"
    Write-Log "Toast requested. Code=$ec; User=$user; Type=$Type; Title='$Title'"
}
#endregion

#region Popup Script
function Initialize-PopupScript {
    New-Item -ItemType Directory -Path $ToastDir -Force | Out-Null
    $sp = Join-Path $ToastDir "Show-Popup.ps1"
    # Using a here-string for the popup script content
    # Title and Body are set programmatically via .Text property to prevent XAML injection
    $scriptContent = @'
param([Parameter(Mandatory)][string]$RequestFile)
$ErrorActionPreference = 'Stop'
try {
    if (-not (Test-Path -LiteralPath $RequestFile)) { exit 0 }
    $req = Get-Content -LiteralPath $RequestFile -Raw | ConvertFrom-Json
    $title = [string]$req.Title
    $body  = [string]$req.Body
    $resp  = [string]$req.ResponseFile
    $opts  = @()
    try { $opts = @($req.SnoozeOptions) } catch { $opts = @(15,30,60,120) }
    if (-not $title) { $title = 'Software Update - Application is open' }
    if (-not $body)  { $body  = 'An update is ready, but the application is currently open.' }

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    [xml]$xamlDoc = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Software Update" WindowStartupLocation="CenterScreen"
        Width="560" MinHeight="320" SizeToContent="Height"
        ResizeMode="NoResize" WindowStyle="None" AllowsTransparency="True"
        Background="Transparent" Topmost="True" ShowInTaskbar="True">
  <Border CornerRadius="16" Background="#FF1F1F1F" BorderBrush="#FF2D2D2D"
          BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>
      <DockPanel Grid.Row="0" LastChildFill="True">
        <TextBlock x:Name="HdrTitle" FontFamily="Segoe UI" FontSize="16"
                   FontWeight="SemiBold" Foreground="White" Text="..."/>
        <Button x:Name="BtnX" DockPanel.Dock="Right" Width="36" Height="28"
                HorizontalAlignment="Right" VerticalAlignment="Top"
                FontFamily="Segoe MDL2 Assets" FontSize="10"
                Content="&#xE8BB;" Cursor="Hand">
          <Button.Style>
            <Style TargetType="Button">
              <Setter Property="Foreground" Value="White"/>
              <Setter Property="Background" Value="Red"/>
              <Setter Property="BorderBrush" Value="#FF111123"/>
              <Setter Property="BorderThickness" Value="1"/>
              <Setter Property="Padding" Value="0"/>
              <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                  <Setter Property="Background" Value="DarkRed"/>
                </Trigger>
                <Trigger Property="IsPressed" Value="True">
                  <Setter Property="Background" Value="DarkRed"/>
                </Trigger>
              </Style.Triggers>
            </Style>
          </Button.Style>
        </Button>
      </DockPanel>
      <ScrollViewer Grid.Row="1" Margin="0,14,0,0"
                    VerticalScrollBarVisibility="Auto"
                    HorizontalScrollBarVisibility="Disabled">
        <StackPanel>
          <TextBlock x:Name="BodyText" FontFamily="Segoe UI" FontSize="13"
                     Foreground="#FFDADADA" TextWrapping="Wrap" Text="..."/>
          <StackPanel Orientation="Horizontal" Margin="0,14,0,0"
                      VerticalAlignment="Center">
            <TextBlock FontFamily="Segoe UI" FontSize="12" Foreground="#FFBDBDBD"
                       Margin="0,0,10,0" Text="Snooze:"
                       VerticalAlignment="Center"/>
            <ComboBox x:Name="SnoozeCombo" Width="180" MinHeight="34"
                      Padding="10,4,10,4" VerticalContentAlignment="Center"
                      FontFamily="Segoe UI" FontSize="12"
                      Background="#FF2A2A2A" Foreground="#FFFFFFFF"
                      BorderBrush="#FF3A3A3A" BorderThickness="1"
                      IsEditable="True" IsReadOnly="True"
                      StaysOpenOnEdit="True">
              <ComboBox.Resources>
                <SolidColorBrush x:Key="{x:Static SystemColors.WindowBrushKey}" Color="#FF2A2A2A"/>
                <SolidColorBrush x:Key="{x:Static SystemColors.WindowTextBrushKey}" Color="#FFFFFFFF"/>
                <SolidColorBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" Color="#FF3A3A3A"/>
                <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="#FFFFFFFF"/>
              </ComboBox.Resources>
              <ComboBox.ItemContainerStyle>
                <Style TargetType="{x:Type ComboBoxItem}">
                  <Setter Property="MinHeight" Value="34"/>
                  <Setter Property="Padding" Value="10,6"/>
                  <Setter Property="Foreground" Value="#FFFFFFFF"/>
                  <Setter Property="Background" Value="#FF2A2A2A"/>
                </Style>
              </ComboBox.ItemContainerStyle>
              <ComboBox.Template>
                <ControlTemplate TargetType="{x:Type ComboBox}">
                  <Grid>
                    <Border x:Name="Bd"
                            Background="{TemplateBinding Background}"
                            BorderBrush="{TemplateBinding BorderBrush}"
                            BorderThickness="{TemplateBinding BorderThickness}"
                            CornerRadius="8" SnapsToDevicePixels="True"/>
                    <DockPanel Margin="8,0,8,0">
                      <ToggleButton x:Name="ToggleButton" DockPanel.Dock="Right"
                                    Focusable="False"
                                    IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                    Background="Transparent"
                                    BorderBrush="Transparent" Width="28">
                        <TextBlock Text="&#xE70D;"
                                   FontFamily="Segoe MDL2 Assets"
                                   Foreground="#FFBDBDBD" FontSize="10"
                                   VerticalAlignment="Center"
                                   HorizontalAlignment="Center"/>
                      </ToggleButton>
                      <TextBox x:Name="PART_EditableTextBox"
                               Background="Transparent"
                               Foreground="{TemplateBinding Foreground}"
                               BorderThickness="0" IsReadOnly="True"
                               VerticalContentAlignment="Center"
                               Padding="0,2,0,0"/>
                    </DockPanel>
                    <Popup x:Name="Popup" Placement="Bottom"
                           IsOpen="{TemplateBinding IsDropDownOpen}"
                           AllowsTransparency="True" Focusable="False"
                           PopupAnimation="Slide">
                      <Border Background="#FF2A2A2A" BorderBrush="#FF3A3A3A"
                              BorderThickness="1" CornerRadius="10"
                              Padding="4" SnapsToDevicePixels="True">
                        <ScrollViewer Margin="0" SnapsToDevicePixels="True">
                          <StackPanel IsItemsHost="True"/>
                        </ScrollViewer>
                      </Border>
                    </Popup>
                  </Grid>
                  <ControlTemplate.Triggers>
                    <Trigger Property="IsMouseOver" Value="True">
                      <Setter TargetName="Bd" Property="BorderBrush" Value="#FF4A4A4A"/>
                    </Trigger>
                    <Trigger Property="IsKeyboardFocusWithin" Value="True">
                      <Setter TargetName="Bd" Property="BorderBrush" Value="#FF60CDFF"/>
                    </Trigger>
                    <Trigger Property="IsEnabled" Value="False">
                      <Setter Property="Opacity" Value="0.55"/>
                    </Trigger>
                  </ControlTemplate.Triggers>
                </ControlTemplate>
              </ComboBox.Template>
            </ComboBox>
          </StackPanel>
        </StackPanel>
      </ScrollViewer>
      <StackPanel Grid.Row="2" Orientation="Horizontal"
                  HorizontalAlignment="Right" Margin="0,18,0,0">
        <Button x:Name="BtnProceed" Width="110" Height="34" Margin="0,0,10,0"
                FontFamily="Segoe UI" FontSize="12" FontWeight="SemiBold"
                Foreground="Black" Background="#FF60CDFF"
                BorderBrush="#00000000" Cursor="Hand" Content="Proceed"/>
        <Button x:Name="BtnSnooze" Width="110" Height="34" Margin="0,0,10,0"
                FontFamily="Segoe UI" FontSize="12" Foreground="White"
                Background="#FF2A2A2A" BorderBrush="#FF3A3A3A"
                Cursor="Hand" Content="Snooze"/>
        <Button x:Name="BtnClose" Width="90" Height="34"
                FontFamily="Segoe UI" FontSize="12" Foreground="#FFDADADA"
                Background="#00000000" BorderBrush="#FF3A3A3A"
                Cursor="Hand" Content="Close"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xamlDoc
    $win = [Windows.Markup.XamlReader]::Load($reader)

    $HdrTitle    = $win.FindName("HdrTitle")
    $BodyText    = $win.FindName("BodyText")
    $SnoozeCombo = $win.FindName("SnoozeCombo")
    $BtnProceed  = $win.FindName("BtnProceed")
    $BtnSnooze   = $win.FindName("BtnSnooze")
    $BtnClose    = $win.FindName("BtnClose")
    $BtnX        = $win.FindName("BtnX")

    # Set text programmatically to prevent XAML injection
    $HdrTitle.Text = $title
    $BodyText.Text = $body

    foreach ($m in $opts) {
        [void]$SnoozeCombo.Items.Add("Snooze $([int]$m) min")
    }
    $di = 0
    for ($i = 0; $i -lt $SnoozeCombo.Items.Count; $i++) {
        if ([string]$SnoozeCombo.Items[$i] -match '60') { $di = $i; break }
    }
    if ($SnoozeCombo.Items.Count -gt 0) { $SnoozeCombo.SelectedIndex = $di }

    $win.Add_MouseLeftButtonDown({ $win.DragMove() })

    function Write-Response([string]$Action, [int]$Minutes) {
        $r = [ordered]@{
            Time          = (Get-Date).ToString('o')
            Action        = $Action
            SnoozeMinutes = $Minutes
        }
        $j = ($r | ConvertTo-Json -Depth 4)
        $tmp = "$resp.tmp.$PID"
        $j | Set-Content -LiteralPath $tmp -Encoding UTF8 -Force
        Move-Item -LiteralPath $tmp -Destination $resp -Force
    }

    $ds = 60

    $BtnProceed.Add_Click({
        Write-Response "Proceed" 0
        $win.Close()
    })

    $BtnSnooze.Add_Click({
        $sel = [string]$SnoozeCombo.SelectedItem
        $m = $ds
        if ($sel -match '(\d+)\s*min') { $m = [int]$matches[1] }
        if ($m -lt 1) { $m = $ds }
        Write-Response "Snooze" $m
        $win.Close()
    })

    $BtnClose.Add_Click({
        Write-Response "Snooze" $ds
        $win.Close()
    })

    $BtnX.Add_Click({
        Write-Response "Snooze" $ds
        $win.Close()
    })

    $win.Add_Closing({
        if (-not (Test-Path -LiteralPath $resp)) {
            Write-Response "Snooze" $ds
        }
    })

    [void]$win.ShowDialog()
} catch {
    exit 0
}
'@
    Set-Content -Path $sp -Value $scriptContent -Encoding UTF8 -Force
    return $sp
}
#endregion

#region Snooze Helpers
function Get-SnoozeFileForApp {
    param([Parameter(Mandatory)][string]$WingetId)
    New-Item -ItemType Directory -Path $SnoozeDir -Force | Out-Null
    return Join-Path $SnoozeDir "snooze-$(Convert-ToSafeFileToken $WingetId).txt"
}

function Get-SnoozeUntil {
    param([Parameter(Mandatory)][string]$WingetId)
    $f = Get-SnoozeFileForApp -WingetId $WingetId
    if (-not (Test-Path -LiteralPath $f)) { return $null }
    try {
        return [datetime]::Parse((Get-Content -LiteralPath $f -Raw).Trim())
    } catch {
        return $null
    }
}

function Set-SnoozeUntil {
    param(
        [Parameter(Mandatory)][string]$WingetId,
        [Parameter(Mandatory)][int]$Minutes
    )
    $until = (Get-Date).AddMinutes($Minutes)
    Set-ContentAtomic -Path (Get-SnoozeFileForApp -WingetId $WingetId) -Content ($until.ToString("o"))
    return $until
}

function Remove-ExpiredSnoozeFiles {
    if (-not (Test-Path -LiteralPath $SnoozeDir)) { return }
    $now = Get-Date
    Get-ChildItem -Path $SnoozeDir -Filter "snooze-*.txt" -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            try {
                $raw = (Get-Content -LiteralPath $_.FullName -Raw).Trim()
                if (-not $raw -or $now -gt [datetime]::Parse($raw)) {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log "Cleaned expired snooze: $($_.Name)"
                }
            } catch {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        }
}
#endregion

#region Install-Directory Process Detection

# Directories to NEVER scan (Windows, system runtimes, broad framework dirs).
# Processes running from these paths are not "app" processes we should close.
$script:InstallDirExclusions = @(
    $env:SystemRoot                                          # C:\Windows
    "$env:SystemRoot\System32"
    "$env:SystemRoot\SysWOW64"
    $env:ProgramFiles                                        # Bare "C:\Program Files" (no subfolder)
    ${env:ProgramFiles(x86)}                                 # Bare "C:\Program Files (x86)"
    "$env:ProgramData"                                       # Bare ProgramData
    "$env:ProgramFiles\WindowsApps"                          # Store apps
    "$env:ProgramFiles\dotnet"                               # .NET runtime
    "$env:ProgramFiles\Common Files"
    "${env:ProgramFiles(x86)}\Common Files"
) | Where-Object { $_ } | ForEach-Object { $_.TrimEnd('\') }

function Get-AppInstallLocation {
    <#
    .SYNOPSIS
      Looks up the install directory for a winget package ID by searching
      the Uninstall registry keys. Returns the first valid path found, or $null.
    .DESCRIPTION
      Searches both native and WOW6432Node uninstall keys. Matches on:
        1. Registry key name contains the winget ID
        2. DisplayName loosely matches the winget app name (if provided)
      Uses InstallLocation first, then falls back to parsing the UninstallString
      or DisplayIcon path to derive the directory.
    #>
    param(
        [Parameter(Mandatory)][string]$WingetId,
        [string]$AppName = ""
    )

    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    # Build search tokens from the winget ID (e.g., "Microsoft.VisualStudioCode" -> "Microsoft", "VisualStudioCode")
    $idTokens = $WingetId -split '\.' | Where-Object { $_.Length -gt 2 }
    # The most specific token is usually the last segment
    $primaryToken = ($idTokens | Select-Object -Last 1)

    $candidates = @()

    foreach ($regPath in $uninstallPaths) {
        $entries = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue
        if (-not $entries) { continue }

        foreach ($entry in $entries) {
            $matched = $false

            # Match 1: Registry key name contains the winget ID or primary token
            $keyName = $entry.PSChildName
            if ($keyName -and (
                $keyName -match [regex]::Escape($WingetId) -or
                $keyName -match [regex]::Escape($primaryToken)
            )) {
                $matched = $true
            }

            # Match 2: DisplayName contains the primary token
            if (-not $matched -and $entry.DisplayName -and $primaryToken) {
                if ($entry.DisplayName -match [regex]::Escape($primaryToken)) {
                    $matched = $true
                }
            }

            # Match 3: If we have the friendly app name from winget, try that
            if (-not $matched -and $AppName -and $entry.DisplayName) {
                # Require at least a significant substring match (first 10+ chars)
                $cmpName = $AppName.Trim()
                if ($cmpName.Length -ge 6 -and $entry.DisplayName -match [regex]::Escape($cmpName)) {
                    $matched = $true
                }
            }

            if (-not $matched) { continue }

            # Extract install path (try multiple sources)
            $dir = $null

            # Source 1: InstallLocation (most reliable)
            if ($entry.InstallLocation -and (Test-Path -LiteralPath $entry.InstallLocation -PathType Container)) {
                $dir = $entry.InstallLocation.TrimEnd('\')
            }

            # Source 2: Parse UninstallString to get directory
            if (-not $dir -and $entry.UninstallString) {
                $uStr = $entry.UninstallString -replace '^"([^"]+)".*', '$1' -replace '^\s*', ''
                if ($uStr -and (Test-Path -LiteralPath $uStr -ErrorAction SilentlyContinue)) {
                    $dir = (Split-Path -Parent $uStr).TrimEnd('\')
                }
            }

            # Source 3: Parse DisplayIcon path
            if (-not $dir -and $entry.DisplayIcon) {
                $iconPath = $entry.DisplayIcon -replace ',\d+$', '' -replace '^"([^"]+)".*', '$1'
                if ($iconPath -and (Test-Path -LiteralPath $iconPath -ErrorAction SilentlyContinue)) {
                    $dir = (Split-Path -Parent $iconPath).TrimEnd('\')
                }
            }

            if ($dir) { $candidates += $dir }
        }
    }

    if (-not $candidates) { return $null }

    # Deduplicate and pick the most specific (longest) path
    $candidates = @($candidates | Select-Object -Unique | Sort-Object Length -Descending)

    # Filter out excluded/too-broad directories
    foreach ($c in $candidates) {
        $dominated = $false
        foreach ($excl in $script:InstallDirExclusions) {
            if ($c -eq $excl -or $c -like "$excl\*") {
                $dominated = $true
                break
            }
        }
        # Also reject paths that are just a drive root (e.g., "C:")
        if (-not $dominated -and $c.Length -gt 3) {
            return $c
        }
    }

    return $null
}

function Get-RunningProcessesInInstallDir {
    <#
    .SYNOPSIS
      Returns unique process names for any process whose executable path
      is under the given install directory.
    .DESCRIPTION
      Uses Get-Process and filters by the .Path property. Only returns
      processes with a valid path (filters out system/idle/etc).
      Excludes common Windows host processes that happen to load DLLs
      from app directories (e.g., svchost, dllhost, conhost).
    #>
    param(
        [Parameter(Mandatory)][string]$InstallDir
    )

    # Normalize for comparison
    $normDir = $InstallDir.TrimEnd('\') + '\'

    # Host processes that may legitimately have handles into app dirs
    # but should never be "closed" as part of an app update
    $hostExclusions = @(
        'svchost', 'dllhost', 'conhost', 'csrss', 'services',
        'lsass', 'smss', 'wininit', 'winlogon', 'explorer',
        'SearchHost', 'RuntimeBroker', 'sihost', 'dwm',
        'taskhostw', 'fontdrvhost', 'WmiPrvSE', 'msiexec',
        'TiWorker', 'TrustedInstaller', 'winget'
    )

    $procs = Get-Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Path -and
            $_.Path.StartsWith($normDir, [StringComparison]::OrdinalIgnoreCase) -and
            $_.Name -notin $hostExclusions
        }

    if (-not $procs) { return @() }

    return @($procs | Select-Object -ExpandProperty Name -Unique)
}

#endregion


#region Open-App Detection + Prompt
function Get-RunningProcessNamesForApp {
    <#
    .SYNOPSIS
      Detects running processes that belong to a given winget package.
      Now handles truncated winget IDs when looking up $AppProcessMap.
    #>
    param(
        [Parameter(Mandatory)][string]$WingetId,
        [string]$AppName = ""
    )

    $running = @()

    # ── Layer 1: Hardcoded process map (truncation-aware lookup) ──────
    $mapKey = $null
    if ($AppProcessMap.ContainsKey($WingetId)) {
        # Exact match (non-truncated ID)
        $mapKey = $WingetId
    } else {
        # Try truncation-aware match against all map keys
        foreach ($k in $AppProcessMap.Keys) {
            if (Test-WingetIdMatch -ParsedId $WingetId -KnownId $k) {
                $mapKey = $k
                break
            }
        }
    }

    if ($mapKey) {
        $names = @($AppProcessMap[$mapKey] | Where-Object { $_ -and $_.Trim() })
        foreach ($n in $names) {
            if (Get-Process -Name $n -ErrorAction SilentlyContinue) { $running += $n }
        }
        # Map entry exists but none running — trust the map, skip Layer 2
        if ($running.Count -eq 0) { return @() }
    }

    # ── Layer 2: Install-directory scan ───────────────────────────────
    $installDir = Get-AppInstallLocation -WingetId $WingetId -AppName $AppName
    if ($installDir) {
        $dirProcs = Get-RunningProcessesInInstallDir -InstallDir $installDir
        if ($dirProcs.Count -gt 0) {
            Write-Log "Install-dir scan for $WingetId ($installDir): found running: $($dirProcs -join ', ')"
            $running = @(($running + $dirProcs) | Select-Object -Unique)
        }
    }

    return @($running | Select-Object -Unique)
}

function Close-AppProcessesBestEffort {
    param([Parameter(Mandatory)][string[]]$ProcessNames)
    foreach ($n in $ProcessNames) {
        if ([string]::IsNullOrWhiteSpace($n)) { continue }
        $procs = @(Get-Process -Name $n -ErrorAction SilentlyContinue)
        foreach ($p in $procs) {
            try { $null = $p.CloseMainWindow() } catch { }
        }
        $deadline = (Get-Date).AddSeconds(5)
        while ((Get-Date) -lt $deadline) {
            if (-not (Get-Process -Name $n -ErrorAction SilentlyContinue)) { break }
            Start-Sleep -Milliseconds 300
        }
        $still = @(Get-Process -Name $n -ErrorAction SilentlyContinue)
        if ($still.Count -gt 0) {
            Write-Log "Force-closing '$n'" "WARN"
            $still | Stop-Process -Force -ErrorAction SilentlyContinue
        } else {
            Write-Log "Closed '$n' gracefully."
        }
    }
}

function Invoke-UserPromptForOpenApp {
    param(
        [Parameter(Mandatory)][pscustomobject]$App,
        [Parameter(Mandatory)][string[]]$ProcessNames
    )
    $user = Get-LoggedOnUser
    if (-not $user -or -not $EnablePopups) {
        return [pscustomobject]@{ Action = "Snooze"; SnoozeMinutes = $DefaultSnoozeMinutes }
    }

    $sp = Initialize-PopupScript
    $reqFile  = Join-Path $ToastDir "popup-request.json"
    $respFile = Join-Path $ToastDir "popup-response.json"

    if (-not (Register-InteractiveUserTask -TaskName "WinomatorPopup" -UserName $user `
        -ScriptPath $sp -ArgName "-RequestFile" -ArgValue $reqFile)) {
        Write-Log "Popup task unavailable; defaulting snooze for $($App.Id)" "WARN"
        return [pscustomobject]@{ Action = "Snooze"; SnoozeMinutes = $DefaultSnoozeMinutes }
    }

    New-Item -ItemType Directory -Path $ToastDir -Force | Out-Null
    Remove-Item -LiteralPath $respFile -Force -ErrorAction SilentlyContinue

    $procList = ($ProcessNames -join ", ")
    $popupTitle = "Software update - '$($App.Name)' is currently open"
    $popupBody = @"
While trying to update your software, we detected that '$($App.Name)' is currently open.

Your computer has a pending update for this application. To safely install the update, the app needs to be closed.

What would you like to do?
- Proceed: $OrganizationName will automatically close the app and install the update right now
- Snooze: Remind me later (you can choose the time before we try again)

Open application(s): $procList
"@

    $payload = [pscustomobject]@{
        Title         = $popupTitle
        Body          = $popupBody
        ResponseFile  = $respFile
        SnoozeOptions = $PopupSnoozeOptionsMinutes
        Time          = (Get-Date).ToString("o")
    }
    Set-ContentAtomic -Path $reqFile -Content ($payload | ConvertTo-Json -Depth 6)

    $startEc = Start-UserTask -TaskName "WinomatorPopup"
    Write-Log "Popup task started. Code=$startEc"

    $deadline = (Get-Date).AddSeconds($PopupResponseTimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-Path -LiteralPath $respFile) { break }
        Start-Sleep -Milliseconds 400
    }

    if (-not (Test-Path -LiteralPath $respFile)) {
        Write-Log "Popup timed out. Snoozing $($App.Id) for $DefaultSnoozeMinutes min." "WARN"
        return [pscustomobject]@{ Action = "Snooze"; SnoozeMinutes = $DefaultSnoozeMinutes }
    }

    try {
        $r = Get-Content -LiteralPath $respFile -Raw | ConvertFrom-Json
        if ([string]$r.Action -eq "Proceed") {
            return [pscustomobject]@{ Action = "Proceed"; SnoozeMinutes = 0 }
        }
        $mins = $DefaultSnoozeMinutes
        try { $mins = [int]$r.SnoozeMinutes } catch { }
        if ($mins -lt 1) { $mins = $DefaultSnoozeMinutes }
        return [pscustomobject]@{ Action = "Snooze"; SnoozeMinutes = $mins }
    } catch {
        Write-Log "Failed to parse popup response. Snoozing $($App.Id)." "WARN"
        return [pscustomobject]@{ Action = "Snooze"; SnoozeMinutes = $DefaultSnoozeMinutes }
    }
}
#endregion


#region Parse winget update output
function Get-WingetAvailableUpdates {
    param([Parameter(Mandatory)][string]$WingetPath)

    New-Item -ItemType Directory -Path $RootDir -Force | Out-Null
    $outF = Join-Path $RootDir "winget-update.out.txt"
    $errF = Join-Path $RootDir "winget-update.err.txt"

    $ec = Invoke-WingetToFiles -WingetPath $WingetPath -Arguments @(
        "update","--source","winget","--accept-source-agreements","--disable-interactivity"
    ) -StdOutFile $outF -StdErrFile $errF

    if ($ec -eq 1460) {
        $prev = Get-MergedLogPreview -StdOutFile $outF -StdErrFile $errF -MaxLines 25
        throw "winget update timed out. Preview: $prev"
    }

    $lines = @()
    if (Test-Path $outF) { $lines += Get-Content $outF -Encoding UTF8 -ErrorAction SilentlyContinue }
    if (Test-Path $errF) { $lines += Get-Content $errF -Encoding UTF8 -ErrorAction SilentlyContinue }
    $lines = $lines | ForEach-Object { Remove-Ansi "$_" }

    $clean = $lines | Where-Object {
        $_ -and
        $_ -notmatch '^\s*(Windows Package Manager|Copyright|Install the latest PowerShell|See https?://)' -and
        $_.Trim()
    } | Where-Object {
        $t = $_.Trim()
        $t -notmatch '^[\-\|\\/]+$' -and
        $t -notmatch '^\s*\d+(\.\d+)?\s*(KB|MB|GB)\s*/' -and
        $t -notmatch '\b(KB|MB|GB)\s*/' -and
        $t -notmatch '^(Downloading|Installing)\b' -and
        $t -notmatch 'msstore|Microsoft Store' -and
        $t -notmatch 'The\s+msstore\s+source\s+requires' -and
        $t -notmatch 'Do you agree to all the source agreements' -and
        $t -notmatch '^\[Y\]\s*Yes\s+\[N\]\s*No'
    }

    $joined = $clean -join "`n"
    if ($joined -match '(No applicable update found|No updates available)') { return @() }

    $headerIndex = -1
    for ($i = 0; $i -lt $clean.Count; $i++) {
        if ($clean[$i] -match '^\s*Name\s+Id\s+Version\s+Available(\s+Source)?\s*$') {
            $headerIndex = $i
            break
        }
    }

    if ($clean.Count -eq 0) {
        $prev = Get-MergedLogPreview -StdOutFile $outF -StdErrFile $errF -MaxLines 30
        throw "winget returned no usable lines. Preview: $prev"
    }
    if ($headerIndex -lt 0) {
        $prev = Get-MergedLogPreview -StdOutFile $outF -StdErrFile $errF -MaxLines 30
        throw "Could not find update header row. Preview: $prev"
    }

    $header  = $clean[$headerIndex]
    $posName = $header.IndexOf("Name")
    $posId   = $header.IndexOf("Id")
    $posVer  = $header.IndexOf("Version")
    $posAvl  = $header.IndexOf("Available")
    $posSrc  = $header.IndexOf("Source")

    if ($posName -lt 0 -or $posId -lt 0 -or $posVer -lt 0 -or $posAvl -lt 0) {
        throw "Header column positions not found."
    }

    $dataStart = $headerIndex + 1
    for ($j = $headerIndex + 1; $j -lt [Math]::Min($clean.Count, $headerIndex + 6); $j++) {
        if ($clean[$j] -match '^-{3,}') { $dataStart = $j + 1; break }
    }
    if ($dataStart -ge $clean.Count) { return @() }

    $results = @()
    foreach ($line in ($clean | Select-Object -Skip $dataStart)) {
        if (-not $line.Trim() -or $line -match '^-{3,}' -or $line -match '^\d+\s+upgrades?\s+available') {
            continue
        }
        $padTo = if ($posSrc -gt 0) { $posSrc + 10 } else { $posAvl + 10 }
        $pad = $line.PadRight([Math]::Max($line.Length, $padTo))
        $name = $pad.Substring($posName, $posId - $posName).Trim()
        $id   = $pad.Substring($posId, $posVer - $posId).Trim()
        $ver  = $pad.Substring($posVer, $posAvl - $posVer).Trim()
        if ($posSrc -gt 0) {
            $avl = $pad.Substring($posAvl, $posSrc - $posAvl).Trim()
            $src = $pad.Substring($posSrc).Trim()
        } else {
            $avl = $pad.Substring($posAvl).Trim()
            $src = ""
        }
        if ($id) {
            $results += [PSCustomObject]@{
                Name = $name; Id = $id; Version = $ver; Available = $avl; Source = $src
            }
        }
    }
    return $results
}

function Test-Excluded {
    param([Parameter(Mandatory)][pscustomobject]$App)
    foreach ($exc in $ExcludedWingetIds) {
        if (Test-WingetIdMatch -ParsedId $App.Id -KnownId $exc) { return $true }
    }
    foreach ($pat in $ExcludedNamePatterns) {
        if ($pat -and $App.Name -match $pat) { return $true }
    }
    return $false
}
#endregion

#region Product Conflict Detection
<#
.SYNOPSIS
  Pre-flight checks to detect when a winget upgrade target conflicts with
  an already-installed superset product (e.g., Acrobat Reader vs Acrobat DC).

.DESCRIPTION
  Maintains a conflict map of known "subset → superset" relationships.
  Before upgrading, checks whether the superset is present via:
    1. Registry detection (Uninstall keys)
    2. File-path detection (known exe paths)
    3. Winget list detection (already-installed packages)

  Returns a skip reason string if a conflict is found, or $null if clear.
#>

# ── Conflict Map ──────────────────────────────────────────────────────
# Each entry: the winget ID being upgraded → array of superset definitions.
# A superset definition contains:
#   - SupersetName:  human-readable name for logging
#   - SupersetIds:   winget IDs that would make this upgrade pointless
#   - RegistryHints: registry paths whose existence signals the superset
#   - FileHints:     file paths whose existence signals the superset
#   - MSIProductCodes: MSI product code prefixes to check in Uninstall keys

$script:ProductConflictMap = @{

    # Adobe Reader is blocked if full Acrobat (DC/Pro/Standard) is installed
    "Adobe.Acrobat.Reader.32-bit" = @(
        @{
            SupersetName    = "Adobe Acrobat DC"
            SupersetIds     = @("Adobe.Acrobat.Pro", "Adobe.Acrobat.Standard")
            RegistryHints   = @(
                "HKLM:\SOFTWARE\Adobe\Acrobat DC\Acrobat"
                "HKLM:\SOFTWARE\WOW6432Node\Adobe\Acrobat DC\Acrobat"
            )
            FileHints       = @(
                "$env:ProgramFiles\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
                "${env:ProgramFiles(x86)}\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
            )
            MSIProductCodes = @()
        }
    )
    "Adobe.Acrobat.Reader.64-bit" = @(
        @{
            SupersetName    = "Adobe Acrobat DC"
            SupersetIds     = @("Adobe.Acrobat.Pro", "Adobe.Acrobat.Standard")
            RegistryHints   = @(
                "HKLM:\SOFTWARE\Adobe\Acrobat DC\Acrobat"
            )
            FileHints       = @(
                "$env:ProgramFiles\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
            )
            MSIProductCodes = @()
        }
    )

    # Example: 7-Zip blocked if NanaZip (fork/superset) is installed
    # "7zip.7zip" = @(
    #     @{
    #         SupersetName  = "NanaZip"
    #         SupersetIds   = @("M2Team.NanaZip")
    #         RegistryHints = @()
    #         FileHints     = @()
    #         MSIProductCodes = @()
    #     }
    # )

    # Template for adding more:
    # "Subset.WingetId" = @(
    #     @{
    #         SupersetName    = "Full Product Name"
    #         SupersetIds     = @("Superset.WingetId")
    #         RegistryHints   = @("HKLM:\SOFTWARE\...")
    #         FileHints       = @("C:\Program Files\...\app.exe")
    #         MSIProductCodes = @("{XXXXXXXX-")  # prefix match
    #     }
    # )
}


# ── Detection Functions ───────────────────────────────────────────────

function Test-SupersetViaRegistry {
    param([Parameter(Mandatory)][string[]]$RegistryHints)
    foreach ($path in $RegistryHints) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }
    return $null
}

function Test-SupersetViaFile {
    param([Parameter(Mandatory)][string[]]$FileHints)
    foreach ($path in $FileHints) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (Test-Path -LiteralPath $path) {
            return $path
        }
    }
    return $null
}

function Test-SupersetViaWinget {
    param(
        [Parameter(Mandatory)][string]$WingetPath,
        [Parameter(Mandatory)][string[]]$SupersetIds
    )
    foreach ($sid in $SupersetIds) {
        if ([string]::IsNullOrWhiteSpace($sid)) { continue }
        $tok = $sid -replace '[^A-Za-z0-9\.\-_]', '_'
        $tmpOut = Join-Path $RootDir "conflict-check-$tok.out.txt"
        $tmpErr = Join-Path $RootDir "conflict-check-$tok.err.txt"
        $null = Invoke-WingetToFiles -WingetPath $WingetPath -Arguments @(
            "list", "--id", $sid, "--exact", "--source", "winget",
            "--accept-source-agreements", "--disable-interactivity"
        ) -StdOutFile $tmpOut -StdErrFile $tmpErr -TimeoutSeconds 60

        if (Test-Path $tmpOut) {
            $content = (Get-Content $tmpOut -Encoding UTF8 -ErrorAction SilentlyContinue |
                ForEach-Object { Remove-Ansi "$_" }) -join "`n"
            # If the output contains the superset ID and does NOT say "No installed package"
            if ($content -match [regex]::Escape($sid) -and
                $content -notmatch '(?i)No installed package found') {
                return $sid
            }
        }
    }
    return $null
}

function Test-SupersetViaMSI {
    param([Parameter(Mandatory)][string[]]$MSIProductCodes)
    $uninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    foreach ($prefix in $MSIProductCodes) {
        if ([string]::IsNullOrWhiteSpace($prefix)) { continue }
        foreach ($regPath in $uninstallPaths) {
            $hits = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue |
                Where-Object { $_.PSChildName -like "$prefix*" -and $_.DisplayName }
            if ($hits) {
                $first = $hits | Select-Object -First 1
                return "$($first.DisplayName) ($($first.PSChildName))"
            }
        }
    }
    return $null
}


# ── Main Entry Point ─────────────────────────────────────────────────

function Test-ProductConflict {
    <#
    .SYNOPSIS
      Checks whether upgrading $WingetId would conflict with an installed superset.
      Now handles truncated winget IDs when looking up $ProductConflictMap.
    .OUTPUTS
      Returns a skip-reason string if conflicted, or $null if safe to proceed.
    #>
    param(
        [Parameter(Mandatory)][string]$WingetId,
        [Parameter(Mandatory)][string]$WingetPath
    )

    # Truncation-aware lookup into conflict map
    $mapKey = $null
    if ($script:ProductConflictMap.ContainsKey($WingetId)) {
        $mapKey = $WingetId
    } else {
        foreach ($k in $script:ProductConflictMap.Keys) {
            if (Test-WingetIdMatch -ParsedId $WingetId -KnownId $k) {
                $mapKey = $k
                break
            }
        }
    }

    if (-not $mapKey) { return $null }

    $conflicts = @($script:ProductConflictMap[$mapKey])

    foreach ($c in $conflicts) {
        $superName = $c.SupersetName

        # 1. Registry check (fastest)
        if ($c.RegistryHints -and $c.RegistryHints.Count -gt 0) {
            $regHit = Test-SupersetViaRegistry -RegistryHints $c.RegistryHints
            if ($regHit) {
                return "Superset '$superName' detected via registry: $regHit"
            }
        }

        # 2. File check (fast, no external calls)
        if ($c.FileHints -and $c.FileHints.Count -gt 0) {
            $fileHit = Test-SupersetViaFile -FileHints $c.FileHints
            if ($fileHit) {
                return "Superset '$superName' detected via file: $fileHit"
            }
        }

        # 3. MSI product code check
        if ($c.MSIProductCodes -and $c.MSIProductCodes.Count -gt 0) {
            $msiHit = Test-SupersetViaMSI -MSIProductCodes $c.MSIProductCodes
            if ($msiHit) {
                return "Superset '$superName' detected via MSI: $msiHit"
            }
        }

        # 4. Winget list check (slowest — only if other checks inconclusive)
        if ($c.SupersetIds -and $c.SupersetIds.Count -gt 0) {
            $wingetHit = Test-SupersetViaWinget -WingetPath $WingetPath -SupersetIds $c.SupersetIds
            if ($wingetHit) {
                return "Superset '$superName' detected via winget list: $wingetHit"
            }
        }
    }

    return $null
}
#endregion

#region Failure Cooldown Cache
<#
.SYNOPSIS
  Remembers recent upgrade failures and suppresses retries during a cooldown window.
  Prevents Winomator from re-downloading 745 MB installers every run for packages
  with broken upstream URLs (404), network errors, or other transient failures.

.DESCRIPTION
  After a failed upgrade, Record-FailedUpgrade writes a JSON file to the FailCache dir.
  Before each upgrade attempt, Test-FailureCooldown checks whether a recent failure
  exists within the cooldown window. If so, the upgrade is skipped.

  Cooldown durations vary by failure category:
    - Download failures (404, network): 48 hours (upstream must fix manifest)
    - Installer failures (1603, 1605):   24 hours (may resolve after reboot)
    - Unknown failures:                  12 hours (conservative retry)

  Files auto-expire and are cleaned up each run.
#>

$FailCacheDir = Join-Path $RootDir "FailCache"

# Cooldown durations by failure category (hours)
$script:FailureCooldownHours = @{
    Download     = 48   # 404, network, TLS — upstream/manifest problem
    Installer    = 24   # MSI 1603, 1605, 1618 — may need reboot or config change
    AccessDenied = 48   # permission issues unlikely to self-resolve
    Unknown      = 12   # conservative default
}

function Get-FailCacheFile {
    param([Parameter(Mandatory)][string]$WingetId)
    New-Item -ItemType Directory -Path $FailCacheDir -Force | Out-Null
    return Join-Path $FailCacheDir "fail-$(Convert-ToSafeFileToken $WingetId).json"
}

function Get-FailureCategory {
    <#
    .SYNOPSIS
      Classifies a failure detail string into a cooldown category.
    #>
    param([Parameter(Mandatory)][string]$Detail)

    # Download / network failures
    if ($Detail -match '(?i)(404|Not found|0x80190194)') { return "Download" }
    if ($Detail -match '(?i)(0x80072EE7|0x80072EFD|0x80072F8F|network|DNS|TLS)') { return "Download" }
    if ($Detail -match '(?i)(Download request status|download failed)') { return "Download" }

    # Access denied
    if ($Detail -match '(?i)(Access is denied|0x80070005|access denied)') { return "AccessDenied" }

    # Installer failures
    if ($Detail -match '(?i)(1603|1605|1618|1619|1638|Installer failed|Install failed)') { return "Installer" }
    if ($Detail -match '(?i)(more functional product|Setup will now terminate)') { return "Installer" }

    return "Unknown"
}

function Record-FailedUpgrade {
    <#
    .SYNOPSIS
      Records a failed upgrade so future runs can skip it during the cooldown window.
    #>
    param(
        [Parameter(Mandatory)][string]$WingetId,
        [Parameter(Mandatory)][string]$VersionAvailable,
        [Parameter(Mandatory)][string]$Detail
    )
    $category = Get-FailureCategory -Detail $Detail
    $cooldownHrs = $script:FailureCooldownHours[$category]
    if (-not $cooldownHrs) { $cooldownHrs = $script:FailureCooldownHours["Unknown"] }

    $record = [pscustomobject]@{
        WingetId         = $WingetId
        VersionAttempted = $VersionAvailable
        FailedAt         = (Get-Date).ToString("o")
        CooldownUntil    = (Get-Date).AddHours($cooldownHrs).ToString("o")
        Category         = $category
        CooldownHours    = $cooldownHrs
        Detail           = $Detail
        FailCount        = 1
    }

    # If there's an existing record for this package, increment fail count
    $file = Get-FailCacheFile -WingetId $WingetId
    if (Test-Path -LiteralPath $file) {
        try {
            $existing = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
            $record.FailCount = ([int]$existing.FailCount) + 1
        } catch { }
    }

    Set-ContentAtomic -Path $file -Content ($record | ConvertTo-Json -Depth 4)
    Write-Log "Recorded failure for $WingetId (category=$category, cooldown=${cooldownHrs}h, failCount=$($record.FailCount))"
}

function Test-FailureCooldown {
    <#
    .SYNOPSIS
      Checks if a package is in failure cooldown. Returns a skip-reason string or $null.
    .DESCRIPTION
      Also checks whether the available version has changed — if winget now offers a
      newer version than what previously failed, the cooldown is cleared so the new
      version gets a fresh attempt.
    #>
    param(
        [Parameter(Mandatory)][string]$WingetId,
        [string]$VersionAvailable = ""
    )
    $file = Get-FailCacheFile -WingetId $WingetId
    if (-not (Test-Path -LiteralPath $file)) { return $null }

    try {
        $record = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
    } catch {
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
        return $null
    }

    # If a new version is available, clear the cooldown — fresh attempt
    if ($VersionAvailable -and $record.VersionAttempted -and
        $VersionAvailable -ne $record.VersionAttempted) {
        Write-Log "New version available for $WingetId ($($record.VersionAttempted) -> $VersionAvailable). Clearing cooldown."
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
        return $null
    }

    # Check if still within cooldown window
    try {
        $until = [datetime]::Parse($record.CooldownUntil)
    } catch {
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
        return $null
    }

    if ((Get-Date) -lt $until) {
        $remaining = [math]::Round(($until - (Get-Date)).TotalHours, 1)
        return "In failure cooldown ($($record.Category), attempt #$($record.FailCount)). " +
               "Last failed: $($record.Detail). Retry in ${remaining}h."
    }

    # Cooldown expired
    Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
    return $null
}

function Remove-ExpiredFailCacheFiles {
    <#
    .SYNOPSIS
      Cleans up expired cooldown files. Call alongside Remove-ExpiredSnoozeFiles.
    #>
    if (-not (Test-Path -LiteralPath $FailCacheDir)) { return }
    $now = Get-Date
    Get-ChildItem -Path $FailCacheDir -Filter "fail-*.json" -File -ErrorAction SilentlyContinue |
        ForEach-Object {
            try {
                $r = Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json
                $until = [datetime]::Parse($r.CooldownUntil)
                if ($now -gt $until) {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log "Cleaned expired fail cache: $($_.Name)"
                }
            } catch {
                Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            }
        }
}

function Clear-FailCacheForPackage {
    <#
    .SYNOPSIS
      Manually clears failure cooldown for a specific package.
      Useful after fixing an issue and wanting to force a retry.
    #>
    param([Parameter(Mandatory)][string]$WingetId)
    $file = Get-FailCacheFile -WingetId $WingetId
    if (Test-Path -LiteralPath $file) {
        Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
        Write-Log "Cleared fail cache for $WingetId"
    }
}
#endregion

#region Single Package Upgrade
function Invoke-SingleUpgrade {
    param(
        [Parameter(Mandatory)][pscustomobject]$App,
        [Parameter(Mandatory)][string]$WingetPath
    )
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $token = Convert-ToSafeFileToken $App.Id
    $outF  = Join-Path $RootDir "upgrade-$token.out.txt"
    $errF  = Join-Path $RootDir "upgrade-$token.err.txt"
    $outF2 = Join-Path $RootDir "upgrade-$token.retry.out.txt"
    $errF2 = Join-Path $RootDir "upgrade-$token.retry.err.txt"

    $baseArgs = @("upgrade","--id",$App.Id,"--exact","--source","winget",
        "--accept-source-agreements","--accept-package-agreements","--disable-interactivity")

    Write-Log "Upgrading $($App.Name) ($($App.Id)) $($App.Version) -> $($App.Available)"

    # Silent attempt
    $code = Invoke-WingetToFiles -WingetPath $WingetPath -Arguments ($baseArgs + @("--silent")) `
        -StdOutFile $outF -StdErrFile $errF
    $preview = Get-CleanedWingetOutput -StdOutFile $outF -StdErrFile $errF

    # Success on exit code 0
    if ($code -eq 0) {
        $verAfter = Get-InstalledWingetVersion -WingetPath $WingetPath -Id $App.Id
        $sw.Stop()
        Clear-FailCacheForPackage -WingetId $App.Id
        Add-RunResult -Name $App.Name -Id $App.Id -VersionBefore $App.Version `
            -VersionAfter $verAfter -VersionAvailable $App.Available `
            -Outcome "Success" -Detail "Silent install OK" -DurationSeconds $sw.Elapsed.TotalSeconds
        Write-Log "Success: $($App.Name) (now $verAfter)"
        return "Success"
    }

    # Success indicated by output text
    if (Test-WingetOutputIndicatesSuccess -T $preview) {
        $verAfter = Get-InstalledWingetVersion -WingetPath $WingetPath -Id $App.Id
        $sw.Stop()
        Clear-FailCacheForPackage -WingetId $App.Id
        Add-RunResult -Name $App.Name -Id $App.Id -VersionBefore $App.Version `
            -VersionAfter $verAfter -VersionAvailable $App.Available `
            -Outcome "Success" -Detail "Output indicates completion" -DurationSeconds $sw.Elapsed.TotalSeconds
        Write-Log "Success (output): $($App.Name)"
        return "Success"
    }

    # Unknown exit code - try version verification
    if ($code -eq 9997) {
        $instNow = Get-InstalledWingetVersion -WingetPath $WingetPath -Id $App.Id
        $verified = $false
        if ($instNow) {
            if ($App.Available -and ($instNow -eq $App.Available -or $instNow.StartsWith($App.Available))) {
                $verified = $true
            } elseif ($App.Version -and $instNow -ne $App.Version) {
                $verified = $true
            }
        }
        if ($verified) {
            $sw.Stop()
            Clear-FailCacheForPackage -WingetId $App.Id
            Add-RunResult -Name $App.Name -Id $App.Id -VersionBefore $App.Version `
                -VersionAfter $instNow -VersionAvailable $App.Available `
                -Outcome "Success" -Detail "Verified via version check" -DurationSeconds $sw.Elapsed.TotalSeconds
            Write-Log "Success (verified): $($App.Name) now at $instNow"
            return "Success"
        }
        Write-Log "Unknown exit (9997), not verifiable. Retrying. Preview: $preview" "WARN"
    }

    $embedded = Get-WingetEmbeddedExitCode -Text $preview
    if ($null -ne $embedded) { $reported = $embedded } else { $reported = $code }
    $noRetry  = Get-WingetNoRetryReason -Text $preview
    $hardFail = Test-WingetOutputIndicatesHardFailure -T $preview

    Write-Log "FAILED ($reported) silent: $($App.Name). Output: $preview" "WARN"

    if ($noRetry) {
        $sw.Stop()
        Add-RunResult -Name $App.Name -Id $App.Id -VersionBefore $App.Version `
            -VersionAfter "" -VersionAvailable $App.Available `
            -Outcome "Failed" -Detail "No retry: $noRetry" -DurationSeconds $sw.Elapsed.TotalSeconds
        Record-FailedUpgrade -WingetId $App.Id -VersionAvailable $App.Available -Detail "No retry: $noRetry"
        return "Failed"
    }
    if ($hardFail -and $null -ne $embedded -and $embedded -ne 1618) {
        $sw.Stop()
        Add-RunResult -Name $App.Name -Id $App.Id -VersionBefore $App.Version `
            -VersionAfter "" -VersionAvailable $App.Available `
            -Outcome "Failed" -Detail "Hard failure (code=$embedded)" -DurationSeconds $sw.Elapsed.TotalSeconds
        Record-FailedUpgrade -WingetId $App.Id -VersionAvailable $App.Available -Detail "Hard failure (code=$embedded)"
        return "Failed"
    }

    # Retry without --silent
    $code2 = Invoke-WingetToFiles -WingetPath $WingetPath -Arguments $baseArgs `
        -StdOutFile $outF2 -StdErrFile $errF2
    $preview2 = Get-CleanedWingetOutput -StdOutFile $outF2 -StdErrFile $errF2

    if ($code2 -eq 0 -or (Test-WingetOutputIndicatesSuccess -T $preview2)) {
        $verAfter = Get-InstalledWingetVersion -WingetPath $WingetPath -Id $App.Id
        $sw.Stop()
        Clear-FailCacheForPackage -WingetId $App.Id
        Add-RunResult -Name $App.Name -Id $App.Id -VersionBefore $App.Version `
            -VersionAfter $verAfter -VersionAvailable $App.Available `
            -Outcome "Success" -Detail "Retry succeeded" -DurationSeconds $sw.Elapsed.TotalSeconds
        Write-Log "Success on retry: $($App.Name)"
        return "Success"
    }

    $embedded2 = Get-WingetEmbeddedExitCode -Text $preview2
    if ($null -ne $embedded2) { $reported2 = $embedded2 } else { $reported2 = $code2 }
    $sw.Stop()
    Add-RunResult -Name $App.Name -Id $App.Id -VersionBefore $App.Version `
        -VersionAfter "" -VersionAvailable $App.Available `
        -Outcome "Failed" -Detail "Retry failed (code=$reported2)" -DurationSeconds $sw.Elapsed.TotalSeconds
    Record-FailedUpgrade -WingetId $App.Id -VersionAvailable $App.Available -Detail "Retry failed (code=$reported2)"
    Write-Log "FAILED ($reported2) retry: $($App.Name). Output: $preview2" "WARN"

    if ($preview2 -match '(?i)exit\s*code:\s*1605') {
        Write-Log "MSI 1605 = per-user install. Consider excluding '$($App.Id)'." "WARN"
    }
    return "Failed"
}
#endregion

# ═══════════════════════════════════════
# MAIN
# ═══════════════════════════════════════

Invoke-LogRotation

if (-not (Enter-WinomatorLock)) {
    Write-Log "Another Winomator instance is running. Exiting." "WARN"
    exit 3
}

try {

Write-Log "Starting Winomator run (PID=$PID)"
Remove-ExpiredSnoozeFiles
Remove-ExpiredFailCacheFiles

$winget = Initialize-WingetForSystem
if (-not $winget) { Write-Log "winget.exe not found." "ERROR"; exit 1 }
Write-Log "Using winget: $winget"

$wt = Test-WingetAsSystem -WingetPath $winget
Write-Log "winget self-test: Ok=$($wt.Ok) ExitCode=$($wt.ExitCode) Out='$($wt.StdOut)'"
if (-not $wt.Ok) { Write-Log "winget self-test failed." "ERROR"; exit 2 }

try {
    $updates = @(Get-WingetAvailableUpdates -WingetPath $winget)
} catch {
    Write-Log "Failed to query updates: $($_.Exception.Message)" "ERROR"
    Send-WinomatorToast -Title "Software Update" -Body "Update scan failed. See logs." -Type Error
    exit 2
}

if (-not $updates -or $updates.Count -eq 0) {
    Write-Log "No applicable updates found."
    Send-WinomatorToast -Title "Software Update" -Body "No updates found." -Type Info
    Export-RunSummary -SuccessCount 0 -FailedCount 0 -TotalAvailable 0 -RebootPending (Test-RebootPending)
    exit 0
}

Write-Log "Found $($updates.Count) updates."
foreach ($u in $updates) {
    Write-Log "  $($u.Name) ($($u.Id)) $($u.Version) -> $($u.Available)"
}

<#
foreach ($u in $updates) {
    $id = $u.Id
    $hasEllipsis = $id.EndsWith([string][char]0x2026)
    $prefix = $id.TrimEnd([char]0x2026)
    $matchResults = @()
    foreach ($exc in $ExcludedWingetIds) {
        $exact = ($id -eq $exc)
        $starts = $exc.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)
        $fnResult = Test-WingetIdMatch -ParsedId $id -KnownId $exc
        if ($exact -or $starts -or $fnResult) {
            $matchResults += "exc=[$exc] exact=$exact starts=$starts fn=$fnResult"
        }
    }
    $excResult = Test-Excluded -App $u
    Write-Log "DEBUG: Id=[$id] Ellipsis=$hasEllipsis Prefix=[$prefix] Excluded=$excResult Matches=[$($matchResults -join ' | ')]"
}
#>

$targets = @($updates | Where-Object { -not (Test-Excluded -App $_) })
if (-not $targets -or $targets.Count -eq 0) {
    Write-Log "All updates excluded."
    Export-RunSummary -SuccessCount 0 -FailedCount 0 -TotalAvailable $updates.Count `
        -RebootPending (Test-RebootPending)
    exit 0
}

# Resolve truncated winget IDs to their full form
$allKnownIds = @($ExcludedWingetIds) +
               @($AppProcessMap.Keys) +
               @($script:ProductConflictMap.Keys) |
               Select-Object -Unique

foreach ($t in $targets) {
    $resolved = Resolve-WingetId -ParsedId $t.Id -KnownIds $allKnownIds
    if ($resolved -ne $t.Id) {
        Write-Log "Resolved truncated ID: $($t.Id) -> $resolved"
        $t.Id = $resolved
    }
}

Write-Log "Upgrading $($targets.Count) packages..."
Send-WinomatorToast -Title "Software updates starting" `
    -Body "Installing $($targets.Count) update(s) in the background." -Type Info

$success   = 0
$failed    = 0
$failedIds = @()

# Separate interactive (running processes) from non-interactive
$interactive    = [System.Collections.ArrayList]::new()
$nonInteractive = [System.Collections.ArrayList]::new()

foreach ($app in $targets) {
    if ([string]::IsNullOrWhiteSpace($app.Id)) {
        Write-Log "Skipping row with empty Id" "WARN"
        continue
    }
    $until = Get-SnoozeUntil -WingetId $app.Id
    if ($until -and (Get-Date) -lt $until) {
        Write-Log "Snoozed: $($app.Id) until $($until.ToString('o'))"
        Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
            -VersionAfter "" -VersionAvailable $app.Available `
            -Outcome "Snoozed" -Detail "Until $($until.ToString('o'))"
        continue
    }
    $running = Get-RunningProcessNamesForApp -WingetId $app.Id -AppName $app.Name
    if ($running.Count -gt 0) {
        [void]$interactive.Add([pscustomobject]@{ App = $app; Running = $running })
    } else {
        [void]$nonInteractive.Add($app)
    }
}

# WhatIf mode
if ($WhatIfPreference) {
    foreach ($app in $nonInteractive) {
        Write-Log "[WhatIf] Would upgrade: $($app.Name) ($($app.Id))"
        Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
            -VersionAfter "" -VersionAvailable $app.Available `
            -Outcome "WhatIf" -Detail "Would upgrade"
    }
    foreach ($item in $interactive) {
        $a = $item.App
        Write-Log "[WhatIf] Would prompt for: $($a.Name) (running: $($item.Running -join ', '))"
        Add-RunResult -Name $a.Name -Id $a.Id -VersionBefore $a.Version `
            -VersionAfter "" -VersionAvailable $a.Available `
            -Outcome "WhatIf" -Detail "Would prompt"
    }
    Export-RunSummary -SuccessCount 0 -FailedCount 0 -TotalAvailable $updates.Count `
        -RebootPending (Test-RebootPending)
    Write-Log "WhatIf complete. No changes made."
    exit 0
}

# Parallel upgrades for non-interactive packages
if ($nonInteractive.Count -gt 0 -and $MaxConcurrentUpgrades -gt 1) {
    Write-Log "Running up to $MaxConcurrentUpgrades parallel upgrades for $($nonInteractive.Count) packages..."
    $jobItems = [System.Collections.ArrayList]::new()

    foreach ($app in $nonInteractive) {
        if (-not $PSCmdlet.ShouldProcess("$($app.Name) ($($app.Id))", "Upgrade")) {
            Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
                -VersionAfter "" -VersionAvailable $app.Available `
                -Outcome "Skipped" -Detail "ShouldProcess declined"
            continue
        }
        
        $conflictReason = Test-ProductConflict -WingetId $app.Id -WingetPath $winget
        if ($conflictReason) {
            Write-Log "SKIPPED $($app.Id): $conflictReason"
            Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
                -VersionAfter "" -VersionAvailable $app.Available `
                -Outcome "Skipped" -Detail $conflictReason
            continue
        }

        $cooldownReason = Test-FailureCooldown -WingetId $app.Id -VersionAvailable $app.Available
        if ($cooldownReason) {
            Write-Log "COOLDOWN $($app.Id): $cooldownReason"
            Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
                -VersionAfter "" -VersionAvailable $app.Available `
                -Outcome "Skipped" -Detail $cooldownReason
            continue
        }

        # Wait if at concurrency limit
        while (($jobItems | Where-Object { $_.Job.State -eq 'Running' }).Count -ge $MaxConcurrentUpgrades) {
            Start-Sleep -Milliseconds 500
        }

        $job = Start-Job -ScriptBlock {
            param($AppId, $WingetExe, $WorkDir)
            $ErrorActionPreference = 'Stop'
            $token = $AppId -replace '[^A-Za-z0-9\.\-_]', '_'
            $outF = Join-Path $WorkDir "upgrade-$token.out.txt"
            $errF = Join-Path $WorkDir "upgrade-$token.err.txt"
            Remove-Item $outF -Force -ErrorAction SilentlyContinue
            Remove-Item $errF -Force -ErrorAction SilentlyContinue

            $wArgs = @("upgrade","--id",$AppId,"--exact","--source","winget",
                "--accept-source-agreements","--accept-package-agreements",
                "--disable-interactivity","--silent")
            $argStr = ($wArgs | ForEach-Object {
                if ($_ -match '\s') { '"' + $_.Replace('"','\"') + '"' } else { $_ }
            }) -join ' '

            $p = Start-Process -FilePath $WingetExe -ArgumentList $argStr `
                -NoNewWindow -PassThru `
                -RedirectStandardOutput $outF -RedirectStandardError $errF
            $exited = $p.WaitForExit(3600000)
            if (-not $exited) {
                try { Stop-Process -Id $p.Id -Force } catch { }
                return "Failed:Timeout"
            }
            try { $p.Refresh() } catch { }
            $ec = $p.ExitCode

            if ($ec -eq 0) { return "Success" }

            # Read ALL output and strip junk before checking for success text
            $lines = @()
            if (Test-Path $outF) { $lines += Get-Content $outF -Encoding UTF8 -ErrorAction SilentlyContinue }
            if (Test-Path $errF) { $lines += Get-Content $errF -Encoding UTF8 -ErrorAction SilentlyContinue }
            $cleaned = $lines | ForEach-Object {
                $t = $_ -replace "`e\[[0-?]*[ -/]*[@-~]", ''
                $t = $t -replace '[\u2580-\u259F]', ''
                $t = $t -replace '(?:â–.)+', ''
                if ($t.Trim() -match '^[-\\|/\s]+$') { return }
                if ($t.Trim()) { $t.Trim() }
            } | Where-Object { $_ }
            $merged = $cleaned -join " | "

            if ($merged -match '(?i)\bSuccessfully installed\b' -or
                $merged -match '(?i)\bNo available upgrade found\b' -or
                $merged -match '(?i)\bNo newer package versions\b' -or
                $merged -match '(?i)\bNo applicable update found\b') {
                return "Success"
            }

            if ($null -eq $ec) { return "Failed:null" }
            return "Failed:$ec"
        } -ArgumentList $app.Id, $winget, $RootDir

        [void]$jobItems.Add([pscustomobject]@{ Job = $job; App = $app })
    }

    # Collect results
    foreach ($item in $jobItems) {
        $result = Receive-Job -Job $item.Job -Wait -AutoRemoveJob
        $a = $item.App
        if ($result -and "$result" -match '^Success') {
            $verAfter = Get-InstalledWingetVersion -WingetPath $winget -Id $a.Id
            $success++
            Clear-FailCacheForPackage -WingetId $a.Id
            Add-RunResult -Name $a.Name -Id $a.Id -VersionBefore $a.Version `
                -VersionAfter $verAfter -VersionAvailable $a.Available `
                -Outcome "Success" -Detail "Parallel OK" -DurationSeconds 0
            Write-Log "Success (parallel): $($a.Name) -> $verAfter"
        } else {
            $failed++
            $failedIds += $a.Id
            Add-RunResult -Name $a.Name -Id $a.Id -VersionBefore $a.Version `
                -VersionAfter "" -VersionAvailable $a.Available `
                -Outcome "Failed" -Detail "Parallel: $result" -DurationSeconds 0
            Write-Log "FAILED (parallel): $($a.Name). Result=$result" "WARN"
            Record-FailedUpgrade -WingetId $a.Id -VersionAvailable $a.Available -Detail "Parallel: $result"
        }
    }
} elseif ($nonInteractive.Count -gt 0) {
    # Sequential fallback (MaxConcurrentUpgrades = 1)
    foreach ($app in $nonInteractive) {
        if (-not $PSCmdlet.ShouldProcess("$($app.Name) ($($app.Id))", "Upgrade")) {
            Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
                -VersionAfter "" -VersionAvailable $app.Available `
                -Outcome "Skipped" -Detail "ShouldProcess declined"
            continue
        }
        $conflictReason = Test-ProductConflict -WingetId $app.Id -WingetPath $winget
        if ($conflictReason) {
            Write-Log "SKIPPED $($app.Id): $conflictReason"
            Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
                -VersionAfter "" -VersionAvailable $app.Available `
                -Outcome "Skipped" -Detail $conflictReason
            continue
        }
        $cooldownReason = Test-FailureCooldown -WingetId $app.Id -VersionAvailable $app.Available
        if ($cooldownReason) {
            Write-Log "COOLDOWN $($app.Id): $cooldownReason"
            Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
                -VersionAfter "" -VersionAvailable $app.Available `
                -Outcome "Skipped" -Detail $cooldownReason
            continue
        }
        $outcome = Invoke-SingleUpgrade -App $app -WingetPath $winget
        if ($outcome -eq "Success") { $success++ } else { $failed++; $failedIds += $app.Id }
    }
}

# Interactive packages (sequential - needs user prompts)
foreach ($item in $interactive) {
    $app     = $item.App
    $running = $item.Running

    $conflictReason = Test-ProductConflict -WingetId $app.Id -WingetPath $winget
    if ($conflictReason) {
        Write-Log "SKIPPED $($app.Id): $conflictReason"
        Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
            -VersionAfter "" -VersionAvailable $app.Available `
            -Outcome "Skipped" -Detail $conflictReason
        continue
    }
    $cooldownReason = Test-FailureCooldown -WingetId $app.Id -VersionAvailable $app.Available
    if ($cooldownReason) {
        Write-Log "COOLDOWN $($app.Id): $cooldownReason"
        Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
            -VersionAfter "" -VersionAvailable $app.Available `
            -Outcome "Skipped" -Detail $cooldownReason
        continue
    }
    Write-Log "Detected running app for $($app.Id): $($running -join ', ')"
    $resp = Invoke-UserPromptForOpenApp -App $app -ProcessNames $running

    if ($resp.Action -eq "Snooze") {
        $until2 = Set-SnoozeUntil -WingetId $app.Id -Minutes ([int]$resp.SnoozeMinutes)
        Write-Log "User snoozed $($app.Id) for $($resp.SnoozeMinutes) min"
        Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
            -VersionAfter "" -VersionAvailable $app.Available `
            -Outcome "Snoozed" -Detail "User chose snooze ($($resp.SnoozeMinutes) min)"
        Send-WinomatorToast -Title "Software Update" `
            -Body "Snoozed '$($app.Name)' for $($resp.SnoozeMinutes) minutes." -Type Warning
        continue
    }

    if (-not $PSCmdlet.ShouldProcess("$($app.Name) ($($app.Id))", "Close and upgrade")) {
        Add-RunResult -Name $app.Name -Id $app.Id -VersionBefore $app.Version `
            -VersionAfter "" -VersionAvailable $app.Available `
            -Outcome "Skipped" -Detail "ShouldProcess declined"
        continue
    }

    Write-Log "User chose Proceed; closing processes for $($app.Id)..."
    Close-AppProcessesBestEffort -ProcessNames $running

    $outcome = Invoke-SingleUpgrade -App $app -WingetPath $winget
    if ($outcome -eq "Success") { $success++ } else { $failed++; $failedIds += $app.Id }
}

# Reboot check
$rebootPending = Test-RebootPending
if ($rebootPending) {
    Write-Log "Reboot pending after upgrades." "WARN"
    Send-WinomatorToast -Title "Reboot Required" `
        -Body "One or more updates require a restart." -Type Warning
}

# Summary
Write-Log "Winomator complete. Success=$success Failed=$failed"
Export-RunSummary -SuccessCount $success -FailedCount $failed `
    -TotalAvailable $updates.Count -RebootPending $rebootPending

if ($failed -eq 0) {
    Send-WinomatorToast -Title "Software updates completed" `
        -Body "Installed $success update(s). No failures." -Type Success
} else {
    $uniq = @($failedIds | Select-Object -Unique)
    $failList = ($uniq | Select-Object -First 5) -join ", "
    $more = ""
    if ($uniq.Count -gt 5) { $more = " (more...)" }
    Send-WinomatorToast -Title "Software updates completed" `
        -Body "Installed $success update(s). $failed failed: $failList$more" -Type Error
}

} finally {
    Unregister-WinomatorTasks
    Exit-WinomatorLock
}

exit 0
