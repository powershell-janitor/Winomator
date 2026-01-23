<#
Winomator – Winget Bulk Updater + BurntToast User Notifications
Runs as: SYSTEM (NinjaOne weekly)

Adds:
- Interactive popups if an app needing updates is currently open
  - User can Snooze or Proceed
  - Proceed closes the app (best-effort) and continues the update
  - Snooze writes a per-app snooze-until file in %ProgramData%\Winomator\Snooze\

Key fixes retained:
- Prevent msstore agreement prompts by forcing enumeration to source=winget and disabling interactivity
- Harden parsing so progress/spinner noise can't break header detection (ASCII-only filters)
- If enumeration fails, still attempt a toast ("scan failed") so notifications don't "disappear"
#>

# ===============================
# Configuration
# ===============================
$ErrorActionPreference = 'Stop'

$LogRoot    = "$env:ProgramData\Winomator\Logs"
$LogFile    = Join-Path $LogRoot "Winomator-Weekly.log"
$RootDir    = Join-Path $env:ProgramData "Winomator"
$ToastDir   = Join-Path $RootDir "Toast"
$SnoozeDir  = Join-Path $RootDir "Snooze"

$ExcludedWingetIds = @(
    "VMware.HorizonClient"
)

$ExcludedNamePatterns = @()
$EnableToasts = $true

# Enable/disable interactive prompts
$EnablePopups = $true

# How long SYSTEM waits for the user to respond to the popup before defaulting
$PopupResponseTimeoutSeconds = 120

# Default snooze if user closes the window or times out
$DefaultSnoozeMinutes = 60

# Allowed snooze options shown in the UI
$PopupSnoozeOptionsMinutes = @(15, 30, 60, 120)

# Map Winget IDs -> process names (WITHOUT ".exe") to detect "app is open"
# Add to this list over time as you see real-world installs.
$AppProcessMap = @{
    "Cisco.Webex"          = @("webex","ciscocollabhost","ptoneclk","wbxreport")
    "Microsoft.OneDrive"   = @("OneDrive")
    "Citrix.Workspace"     = @("SelfService","Receiver","wfcrun32","wfica32")
    "VMware.HorizonClient" = @("vmware-view","vmware-horizon-client")
    "Microsoft.Teams"      = @("ms-teams","Teams","TeamsUpdater")
}

$ToastImageMap = @{
    Info    = (Join-Path $ToastDir "info.png")
    Success = (Join-Path $ToastDir "success.png")
    Warning = (Join-Path $ToastDir "warning.png")
    Error   = (Join-Path $ToastDir "error.png")
}

$ToastBuiltinCandidates = @{
    Info    = @(
        "$env:windir\System32\SecurityAndMaintenance.png"
    )
    Success = @(
        "$env:windir\System32\SecurityAndMaintenance.png"
    )
    Warning = @(
        "$env:windir\System32\SecurityAndMaintenance_Alert.png"
    )
    Error   = @(
        "$env:windir\System32\SecurityAndMaintenance_Error.png"
    )
}

# ===============================
# Logging
# ===============================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO"
    )
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "$ts [$Level] $Message"
    New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    Write-Output $line
}

# ===============================
# winget discovery
# ===============================

function Test-WingetOutputIndicatesSuccess {
    param([Parameter(Mandatory)][string]$PreviewOrText)

    if ($PreviewOrText -match '(?i)\bNo installed package found\b') { return $false }

    # Treat these as success states (even if ExitCode is missing/odd)
    return ($PreviewOrText -match '(?i)\bSuccessfully installed\b' -or
            $PreviewOrText -match '(?i)\bNo available upgrade found\b' -or
            $PreviewOrText -match '(?i)\bNo upgrade available\b' -or
            $PreviewOrText -match '(?i)\bNo newer package versions\b' -or
            $PreviewOrText -match '(?i)\bNo applicable update found\b' -or
            $PreviewOrText -match '(?i)\bNo updates available\b')
}

function Test-WingetOutputIndicatesHardFailure {
    param([Parameter(Mandatory)][string]$PreviewOrText)

    return (
        $PreviewOrText -match '(?i)\bInstaller failed with exit code:\s*\d+\b' -or
        $PreviewOrText -match '(?i)\bUninstall failed with exit code:\s*\d+\b' -or
        $PreviewOrText -match '(?i)\bInstall failed with exit code:\s*\d+\b' -or
        $PreviewOrText -match '(?i)\bError:\b'
    )
}

function Convert-ToSafeFileToken {
    param([Parameter(Mandatory)][string]$Text)
    return ($Text -replace '[^A-Za-z0-9\.\-_]', '_')
}

function Remove-ProgressJunk {
    param([string]$Text)
    if ($null -eq $Text) { return $Text }

    # Strip common progress/spinner glyphs and mojibake block chars
    $t = $Text

    # Remove unicode block elements + their mojibake representations
    $t = $t -replace '[\u2580-\u259F]', ''          # ▀-▟
    $t = $t -replace '(?:â–.)+', ''

    # Remove simple spinners like "- \ | /" when they are the only content
    if ($t.Trim() -match '^[-\\|/]+$') { return '' }

    return $t
}

function Get-WingetPath {
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }

    $match = Get-Item "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*\winget.exe" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1

    if ($match) { return $match.FullName }
    return $null
}

function Get-WingetEmbeddedExitCode {
    param([string]$Text)
    if ($Text -match '(?i)\bInstaller failed with exit code:\s*(\d+)\b') { return [int]$matches[1] }
    if ($Text -match '(?i)\bUninstall failed with exit code:\s*(\d+)\b') { return [int]$matches[1] }
    if ($Text -match '(?i)\bInstall failed with exit code:\s*(\d+)\b')   { return [int]$matches[1] }  # add this
    return $null
}

function Get-WingetNoRetryReason {
    param([Parameter(Mandatory)][string]$Text)

    if ($Text -match '(?i)\bNo installed package found\b') { return "Not installed / not visible to SYSTEM" }
    if ($Text -match '(?i)\bNo package found matching input criteria\b') { return "Not found in configured sources" }

    # MSI busy - retry later, not immediately
    if ($Text -match '(?i)\bAnother installation is already in progress\b' -or $Text -match '(?i)\b1618\b') { 
        return "Windows Installer busy (1618)" 
    }

    # Common “access denied / policy” style failures
    if ($Text -match '(?i)\bAccess is denied\b|\b0x80070005\b') { return "Access denied (0x80070005)" }

    # Winget/network issues (common under SYSTEM if proxy/DNS weird)
    if ($Text -match '(?i)\b0x80072EE7\b|\b0x80072EFD\b|\b0x80072F8F\b') { return "Network/DNS/TLS failure" }

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
    if (Test-Path $StdOutFile) { Remove-Item $StdOutFile -Force -ErrorAction SilentlyContinue }
    if (Test-Path $StdErrFile) { Remove-Item $StdErrFile -Force -ErrorAction SilentlyContinue }

    $argString = ($Arguments | ForEach-Object {
        if ($_ -match '\s') { '"' + $_.Replace('"','\"') + '"' } else { $_ }
    }) -join ' '

    try {
        $p = Start-Process -FilePath $WingetPath `
            -ArgumentList $argString `
            -NoNewWindow `
            -PassThru `
            -RedirectStandardOutput $StdOutFile `
            -RedirectStandardError  $StdErrFile
    } catch {
        Write-Log "Start-Process failed launching winget: $($_.Exception.Message)" "ERROR"
        return 9998
    }

    $exited = $p.WaitForExit($TimeoutSeconds * 1000)
    if (-not $exited) {
        try { Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue } catch {}
        return 1460
    }

    try { $p.Refresh() } catch {}
    $ec = $p.ExitCode

    if ($null -eq $ec) {
        # ExitCode is missing — infer from output
        $mergedPreview = Get-MergedLogPreview -StdOutFile $StdOutFile -StdErrFile $StdErrFile -MaxLines 80

        if ($mergedPreview -match '(?i)\bSuccessfully installed\b') {
            return 0
        }
        if ($mergedPreview -match '(?i)\bNo available upgrade found\b|\bNo newer package versions\b|\bNo applicable update found\b|\bNo updates available\b') {
            # Treat "nothing to do" as success for upgrade/update commands
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

    $tmpOut = Join-Path $RootDir "list-$((Convert-ToSafeFileToken $Id)).out.txt"
    $tmpErr = Join-Path $RootDir "list-$((Convert-ToSafeFileToken $Id)).err.txt"

    $ec = Invoke-WingetToFiles -WingetPath $WingetPath -Arguments @(
        "list",
        "--id", $Id,
        "--exact",
        "--source", "winget",
        "--accept-source-agreements",
        "--disable-interactivity"
    ) -StdOutFile $tmpOut -StdErrFile $tmpErr -TimeoutSeconds 120

    if (-not (Test-Path $tmpOut)) { return $null }

    if ((Test-Path $tmpErr) -and ((Get-Content $tmpErr -Raw -ErrorAction SilentlyContinue) -match 'agree to all the source agreements')) {
        Write-Log "winget list hit msstore agreement prompt unexpectedly for Id=$Id (forcing source should prevent this)." "WARN"
    }
    
    $lines = Get-Content $tmpOut -Encoding UTF8 -ErrorAction SilentlyContinue
    $lines = $lines | ForEach-Object { Remove-Ansi "$_" } | Where-Object { $_ -and $_.Trim() }

    # Crude but effective: find the line that contains the Id and pull a version-ish token
    $hit = $lines | Where-Object { $_ -match [regex]::Escape($Id) } | Select-Object -First 1
    if (-not $hit) { return $null }

    # Many winget list outputs are fixed-width; easiest is to grab last "version looking" token
    $ver = ([regex]::Matches($hit, '\d+(\.\d+){1,3}') | Select-Object -First 1).Value
    if ([string]::IsNullOrWhiteSpace($ver)) { return $null }
    return $ver
}

function Ensure-WingetForSystem {
    param(
        [string]$CacheRoot = "$env:ProgramData\Winomator\WingetRuntime"
    )

    # 1) First try the installed packaged winget
    $installed = Get-WingetPath
    if ($installed) {
        $t = Test-WingetAsSystem -WingetPath $installed
        if ($t.Ok) { return $installed }

        # If it's NOT the known "missing dependency" failure, return installed so caller can log it
        if ($t.ExitCode -ne -1073741515) { return $installed }
    }

    # Helper: extract zip container regardless of extension using tar.exe
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
        if (-not (Test-Path -LiteralPath $tar)) { throw "tar.exe not found; cannot extract msixbundle/msix/appx." }

        # capture tar output for diagnostics
        $tarOut = & $tar -xf $Path -C $Destination 2>&1
        $ec = $LASTEXITCODE

        if ($ec -ne 0) {
            $msg = ($tarOut | Select-Object -First 12) -join " | "
            throw "Extraction failed for '$Path' (tar exit code $ec). Output: $msg"
        }
    }


    # 2) Build an extracted “runtime” copy in ProgramData
    New-Item -ItemType Directory -Path $CacheRoot -Force | Out-Null
    $bundle = Join-Path $CacheRoot "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"

    Invoke-WebRequest -Uri "https://aka.ms/getwinget" -OutFile $bundle -UseBasicParsing

    # quick sanity check
    $fi = Get-Item -LiteralPath $bundle -ErrorAction SilentlyContinue
    if (-not $fi -or $fi.Length -lt 2000000) {  # < ~2MB is suspicious for the msixbundle
        throw "Downloaded winget bundle looks wrong (missing/small). Path=$bundle SizeBytes=$($fi.Length)"
    }

    $bundleExtract = Join-Path $CacheRoot "bundle"
    Expand-ZipLike -Path $bundle -Destination $bundleExtract

    # Pick x64 payload inside (msix/appx)
    $payload = Get-ChildItem $bundleExtract -Recurse -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '(?i)x64' -and $_.Extension -match '(?i)\.(msix|appx)$' } |
        Sort-Object Length -Descending |
        Select-Object -First 1

    if (-not $payload) {
        $payload = Get-ChildItem $bundleExtract -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '(?i)\.(msix|appx)$' } |
            Sort-Object Length -Descending |
            Select-Object -First 1
    }
    if (-not $payload) { throw "Could not find an .msix or .appx inside the msixbundle." }

    $msixExtract = Join-Path $CacheRoot "msix"
    Expand-ZipLike -Path $payload.FullName -Destination $msixExtract

    $winget = Get-ChildItem $msixExtract -Recurse -Filter "winget.exe" -File -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty FullName

    if (-not $winget) { throw "Extracted App Installer payload but could not locate winget.exe." }

    $t2 = Test-WingetAsSystem -WingetPath $winget
    if (-not $t2.Ok) {
        throw "Extracted winget still not runnable under SYSTEM. ExitCode=$($t2.ExitCode) StdOut='$($t2.StdOut)' StdErr='$($t2.StdErr)'"
    }

    return $winget
}

# ===============================
# ANSI/VT stripping
# ===============================

function Remove-Ansi {
    param([string]$Text)
    if ($null -eq $Text) { return $Text }
    $ansiPattern = "`e\[[0-?]*[ -/]*[@-~]"
    return ($Text -replace $ansiPattern, "")
}

function Test-WingetAsSystem {
    param([Parameter(Mandatory)][string]$WingetPath)

    $out = ""
    $err = ""
    $ec  = $null

    try {
        $out = & $WingetPath --version 2>$null
        $ec = $LASTEXITCODE
    } catch {
        $err = $_.Exception.Message
        $ec = 9999
    }

    $out = if ($null -eq $out) { "" } else { ([string]$out).Trim() }
    $err = if ($null -eq $err) { "" } else { ([string]$err).Trim() }

    $hasVersion = ($out -match '^\s*\d+(\.\d+){1,3}\s*$')

    # Compute byte counts BEFORE the object (no "if" expressions)
    $outBytes = 0
    $errBytes = 0
    try {
        if (-not [string]::IsNullOrWhiteSpace($out)) {
            $outBytes = [Text.Encoding]::UTF8.GetByteCount($out)
        }
        if (-not [string]::IsNullOrWhiteSpace($err)) {
            $errBytes = [Text.Encoding]::UTF8.GetByteCount($err)
        }
    } catch { }

    [pscustomobject]@{
        ExitCode   = $ec
        StdOut     = $out
        StdErr     = $err
        HasVersion = $hasVersion
        Ok         = ($hasVersion -or ($ec -eq 0))
        OutBytes   = $outBytes
        ErrBytes   = $errBytes
    }
}

function Get-MergedLogPreview {
    param(
        [Parameter(Mandatory)][string]$StdOutFile,
        [Parameter(Mandatory)][string]$StdErrFile,
        [int]$MaxLines = 14
    )

    $lines = @()
    if (Test-Path $StdOutFile) { $lines += Get-Content -Path $StdOutFile -Encoding UTF8 -ErrorAction SilentlyContinue }
    if (Test-Path $StdErrFile) { $lines += Get-Content -Path $StdErrFile -Encoding UTF8 -ErrorAction SilentlyContinue }
    $lines = $lines | ForEach-Object {
        Remove-ProgressJunk (Remove-Ansi "$_")
    } | Where-Object { $_ -and $_.Trim() }

    return (($lines | Select-Object -First $MaxLines) -join " | ")
}

# ===============================
# Logged-on user discovery
# ===============================

function Get-LoggedOnUser {
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $u = $cs.UserName
        if ([string]::IsNullOrWhiteSpace($u)) { return $null }
        return $u
    } catch {
        return $null
    }
}

# ===============================
# BurntToast install (AllUsers)
# ===============================

function Ensure-BurntToastInstalled {
    try {
        $existing = Get-Module -ListAvailable -Name BurntToast -ErrorAction SilentlyContinue
        if ($existing) { return $true }

        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -Force -Scope AllUsers | Out-Null
        }

        try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}

        Install-Module -Name BurntToast -Scope AllUsers -Force -AllowClobber -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        Write-Log "BurntToast install failed: $($_.Exception.Message)" "WARN"
        return $false
    }
}

# ===============================
# Toast image selection helpers
# ===============================

function Resolve-ToastImagePath {
    param(
        [ValidateSet("Info","Success","Warning","Error")]
        [string]$Type = "Info",
        [string]$ExplicitPath
    )

    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        if (Test-Path -LiteralPath $ExplicitPath) { return $ExplicitPath }
    }

    if ($ToastImageMap.ContainsKey($Type)) {
        $p = [string]$ToastImageMap[$Type]
        if (-not [string]::IsNullOrWhiteSpace($p) -and (Test-Path -LiteralPath $p)) { return $p }
    }

    if ($ToastBuiltinCandidates.ContainsKey($Type)) {
        foreach ($cand in $ToastBuiltinCandidates[$Type]) {
            if (-not [string]::IsNullOrWhiteSpace($cand) -and (Test-Path -LiteralPath $cand)) { return $cand }
        }
    }

    return $null
}

# ===============================
# Toast script + wrapper (quoting-safe)
# ===============================

function Ensure-ToastScript {
    New-Item -ItemType Directory -Path $ToastDir -Force | Out-Null
    $scriptPath = Join-Path $ToastDir "Show-Toast.ps1"

@"
param(
  [Parameter(Mandatory)][string]`$StatusFile
)

`$ErrorActionPreference = 'Stop'

`$diagRan   = Join-Path (Split-Path -Parent `$StatusFile) 'toast-ran.txt'
`$diagError = Join-Path (Split-Path -Parent `$StatusFile) 'toast-error.txt'

try {
  "`$(Get-Date -Format o) Task ran as: `$(whoami)" | Out-File -FilePath `$diagRan -Encoding UTF8 -Append

  if (-not (Test-Path -LiteralPath `$StatusFile)) {
    "Status file not found: `$StatusFile" | Out-File -FilePath `$diagError -Encoding UTF8 -Append
    exit 0
  }

  `$json = Get-Content -LiteralPath `$StatusFile -Raw | ConvertFrom-Json

  `$title = [string]`$json.Title
  `$body  = [string]`$json.Body
  `$img   = [string]`$json.ImagePath

  if ([string]::IsNullOrWhiteSpace(`$title)) { `$title = "Software Update" }
  if ([string]::IsNullOrWhiteSpace(`$body))  { `$body  = "Update status." }

  Import-Module BurntToast -ErrorAction Stop

  if (-not [string]::IsNullOrWhiteSpace(`$img) -and (Test-Path -LiteralPath `$img)) {
    New-BurntToastNotification -Text `$title, `$body -AppLogo `$img | Out-Null
  } else {
    New-BurntToastNotification -Text `$title, `$body | Out-Null
  }
}
catch {
  "`$(Get-Date -Format o) ERROR: `$($_.Exception.Message)" | Out-File -FilePath `$diagError -Encoding UTF8 -Append
  exit 0
}
"@ | Set-Content -Path $scriptPath -Encoding UTF8 -Force

    return $scriptPath
}

function Ensure-ToastTaskForUser {
    param([Parameter(Mandatory)][string]$UserName)

    $taskName = "WinomatorToast"
    [void](Ensure-ToastScript)

    $scriptPath = Join-Path $ToastDir "Show-Toast.ps1"
    $statusFile = Join-Path $ToastDir "status.json"

    return (Ensure-InteractiveUserTask -TaskName $taskName -UserName $UserName -ScriptPath $scriptPath -ArgName "-StatusFile" -ArgValue $statusFile)
}

function Run-ToastTask {
    param([string]$TaskName = "WinomatorToast")
    return (Start-UserTask -TaskName $TaskName)
}

function Send-WinomatorToast {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$Body,
        [ValidateSet("Info","Success","Warning","Error")]
        [string]$Type = "Info",
        [string]$ImagePath
    )

    if (-not $EnableToasts) { return }

    $user = Get-LoggedOnUser
    if ([string]::IsNullOrWhiteSpace($user)) {
        Write-Log "No interactive user session detected; skipping toast: $Title" "INFO"
        return
    }

    if (-not (Ensure-BurntToastInstalled)) {
        Write-Log "BurntToast unavailable; skipping toast: $Title" "WARN"
        return
    }

    if (-not (Ensure-ToastTaskForUser -UserName $user)) {
        Write-Log "Toast task unavailable; skipping toast: $Title" "WARN"
        return
    }

    New-Item -ItemType Directory -Path $ToastDir -Force | Out-Null
    $statusFile = Join-Path $ToastDir "status.json"

    $resolved = Resolve-ToastImagePath -Type $Type -ExplicitPath $ImagePath

    $payload = [pscustomobject]@{
        Title     = $Title
        Body      = $Body
        Type      = $Type
        ImagePath = $resolved
        Time      = (Get-Date).ToString("o")
    }
    $payload | ConvertTo-Json -Depth 4 | Set-Content -Path $statusFile -Encoding UTF8 -Force

    $ec = Run-ToastTask -TaskName "WinomatorToast"

    $diagRan   = Join-Path $ToastDir "toast-ran.txt"
    $diagError = Join-Path $ToastDir "toast-error.txt"

    Write-Log "Toast run requested. TaskStartCode=$ec; User=$user; Type=$Type; Title='$Title'; Img='$resolved'; Diag='$diagRan' Err='$diagError'" "INFO"

}

# ===============================
# Popup prompt script + scheduled task (user context)
# ===============================

function Ensure-PopupScript {
    New-Item -ItemType Directory -Path $ToastDir -Force | Out-Null
    $scriptPath = Join-Path $ToastDir "Show-Popup.ps1"

@"
param(
  [Parameter(Mandatory)][string]`$RequestFile
)

`$ErrorActionPreference = 'Stop'

try {
  if (-not (Test-Path -LiteralPath `$RequestFile)) { exit 0 }
  `$req = Get-Content -LiteralPath `$RequestFile -Raw | ConvertFrom-Json

  `$title = [string]`$req.Title
  `$body  = [string]`$req.Body
  `$resp  = [string]`$req.ResponseFile
  `$opts  = @()
  try { `$opts = @(`$req.SnoozeOptions) } catch { `$opts = @(15,30,60,120) }

  if ([string]::IsNullOrWhiteSpace(`$title)) { `$title = 'Software Update - Application is open' }
  if ([string]::IsNullOrWhiteSpace(`$body))  { `$body  = 'An update is ready, but the application is currently open.' }

  Add-Type -AssemblyName PresentationFramework
  Add-Type -AssemblyName PresentationCore
  Add-Type -AssemblyName WindowsBase

  # Build XAML (Win11-ish styling: rounded corners, clean spacing)
  `$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Software Update"
        WindowStartupLocation="CenterScreen"
        Width="560"
        MinHeight="320"
        SizeToContent="Height"
        ResizeMode="NoResize"
        WindowStyle="None"
        AllowsTransparency="True"
        Background="Transparent"
        Topmost="True"
        ShowInTaskbar="True">
  <Border CornerRadius="16" Background="#FF1F1F1F" BorderBrush="#FF2D2D2D" BorderThickness="1" Padding="18">
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <!-- Header -->
      <DockPanel Grid.Row="0" LastChildFill="True">
        <TextBlock x:Name="HdrTitle"
                   FontFamily="Segoe UI"
                   FontSize="16"
                   FontWeight="SemiBold"
                   Foreground="White"
                   Text="Software Update"/>
            <Button x:Name="BtnX"
                    DockPanel.Dock="Right"
                    Width="36"
                    Height="28"
                    HorizontalAlignment="Right"
                    VerticalAlignment="Top"
                    FontFamily="Segoe MDL2 Assets"
                    FontSize="10"
                    Content="&#xE8BB;"
                    Cursor="Hand">
                <Button.Style>
                    <Style TargetType="Button">
                        <!-- Default (idle) -->
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="Background" Value="Red"/>
                        <Setter Property="BorderBrush" Value="#FF111123"/>
                        <Setter Property="BorderThickness" Value="1"/>
                        <Setter Property="Padding" Value="0"/>

                        <Style.Triggers>
                            <!-- Hover: Windows 11 close red -->
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="DarkRed"/>
                                <Setter Property="BorderBrush" Value="#FF111123"/>
                                <Setter Property="Foreground" Value="White"/>
                            </Trigger>

                            <!-- Pressed: darker red -->
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="DarkRed"/>
                                <Setter Property="BorderBrush" Value="#FF111123"/>
                            </Trigger>

                            <!-- Disabled -->
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </Style.Triggers>
                    </Style>
                </Button.Style>
            </Button>
      </DockPanel>

      <!-- Body -->
      <ScrollViewer Grid.Row="1"
                    Margin="0,14,0,0"
                    VerticalScrollBarVisibility="Auto"
                    HorizontalScrollBarVisibility="Disabled">
        <StackPanel>
          <TextBlock x:Name="BodyText"
                     FontFamily="Segoe UI"
                     FontSize="13"
                     Foreground="#FFDADADA"
                     TextWrapping="Wrap"
                     Text="Message goes here"/>

          <StackPanel Orientation="Horizontal" Margin="0,14,0,0" VerticalAlignment="Center">
            <TextBlock FontFamily="Segoe UI" FontSize="12" Foreground="#FFBDBDBD"
                       Margin="0,0,10,0" Text="Snooze:" VerticalAlignment="Center"/>

            <ComboBox x:Name="SnoozeCombo"
                      Width="180"
                      MinHeight="34"
                      Padding="10,4,10,4"
                      VerticalContentAlignment="Center"
                      FontFamily="Segoe UI"
                      FontSize="12"
                      Background="#FF2A2A2A"
                      Foreground="#FFFFFFFF"
                      BorderBrush="#FF3A3A3A"
                      BorderThickness="1"
                      IsEditable="True"
                      IsReadOnly="True"
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
                            CornerRadius="8"
                            SnapsToDevicePixels="True"/>
                    <DockPanel Margin="8,0,8,0">
                      <ToggleButton x:Name="ToggleButton"
                                    DockPanel.Dock="Right"
                                    Focusable="False"
                                    IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                    Background="Transparent"
                                    BorderBrush="Transparent"
                                    Width="28">
                        <TextBlock Text="&#xE70D;"
                                   FontFamily="Segoe MDL2 Assets"
                                   Foreground="#FFBDBDBD"
                                   FontSize="10"
                                   VerticalAlignment="Center"
                                   HorizontalAlignment="Center"/>
                      </ToggleButton>

                      <TextBox x:Name="PART_EditableTextBox"
                               Background="Transparent"
                               Foreground="{TemplateBinding Foreground}"
                               BorderThickness="0"
                               IsReadOnly="True"
                               VerticalContentAlignment="Center"
                               Padding="0,2,0,0"/>
                    </DockPanel>

                    <Popup x:Name="Popup"
                           Placement="Bottom"
                           IsOpen="{TemplateBinding IsDropDownOpen}"
                           AllowsTransparency="True"
                           Focusable="False"
                           PopupAnimation="Slide">
                      <Border Background="#FF2A2A2A"
                              BorderBrush="#FF3A3A3A"
                              BorderThickness="1"
                              CornerRadius="10"
                              Padding="4"
                              SnapsToDevicePixels="True">
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

      <!-- Buttons -->
      <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,18,0,0">
        <Button x:Name="BtnProceed"
                Width="110" Height="34"
                Margin="0,0,10,0"
                FontFamily="Segoe UI"
                FontSize="12"
                FontWeight="SemiBold"
                Foreground="Black"
                Background="#FF60CDFF"
                BorderBrush="#00000000"
                Cursor="Hand"
                Content="Proceed"/>
        <Button x:Name="BtnSnooze"
                Width="110" Height="34"
                Margin="0,0,10,0"
                FontFamily="Segoe UI"
                FontSize="12"
                Foreground="White"
                Background="#FF2A2A2A"
                BorderBrush="#FF3A3A3A"
                Cursor="Hand"
                Content="Snooze"/>
        <Button x:Name="BtnClose"
                Width="90" Height="34"
                FontFamily="Segoe UI"
                FontSize="12"
                Foreground="#FFDADADA"
                Background="#00000000"
                BorderBrush="#FF3A3A3A"
                Cursor="Hand"
                Content="Close"/>
      </StackPanel>
    </Grid>
  </Border>
</Window>
'@

  `$reader = New-Object System.Xml.XmlNodeReader ([xml]`$xaml)
  `$win = [Windows.Markup.XamlReader]::Load(`$reader)

  # Bind controls
  `$HdrTitle   = `$win.FindName("HdrTitle")
  `$BodyText   = `$win.FindName("BodyText")
  `$SnoozeCombo= `$win.FindName("SnoozeCombo")
  `$BtnProceed = `$win.FindName("BtnProceed")
  `$BtnSnooze  = `$win.FindName("BtnSnooze")
  `$BtnClose   = `$win.FindName("BtnClose")
  `$BtnX       = `$win.FindName("BtnX")

  `$HdrTitle.Text = `$title
  `$BodyText.Text = `$body

  foreach (`$m in `$opts) { [void]`$SnoozeCombo.Items.Add("Snooze `$( [int]`$m ) min") }

  # Prefer 60 min if present
  `$defaultIndex = 0
  for (`$i=0; `$i -lt `$SnoozeCombo.Items.Count; `$i++) {
    if ([string]`$SnoozeCombo.Items[`$i] -match '60') { `$defaultIndex = `$i; break }
  }
  if (`$SnoozeCombo.Items.Count -gt 0) { `$SnoozeCombo.SelectedIndex = `$defaultIndex }

  # Allow dragging window by clicking anywhere on the border
  `$win.Add_MouseLeftButtonDown({ `$win.DragMove() })

  function Write-Response([string]`$Action,[int]`$Minutes) {
    `$result = [ordered]@{
      Time = (Get-Date).ToString('o')
      Action = `$Action
      SnoozeMinutes = `$Minutes
    }
    (`$result | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath `$resp -Encoding UTF8 -Force
  }

  # default snooze minutes if close is used
  `$defaultSnooze = 60

  `$BtnProceed.Add_Click({
    Write-Response -Action "Proceed" -Minutes 0
    `$win.Close()
  })

  `$BtnSnooze.Add_Click({
    `$sel = [string]`$SnoozeCombo.SelectedItem
    `$mins = `$defaultSnooze
    if (`$sel -match '(\d+)\s*min') { `$mins = [int]`$matches[1] }
    if (`$mins -lt 1) { `$mins = `$defaultSnooze }
    Write-Response -Action "Snooze" -Minutes `$mins
    `$win.Close()
  })

  `$BtnClose.Add_Click({
    Write-Response -Action "Snooze" -Minutes `$defaultSnooze
    `$win.Close()
  })

  `$BtnX.Add_Click({
    Write-Response -Action "Snooze" -Minutes `$defaultSnooze
    `$win.Close()
  })

  # If user alt-f4s etc, still write a response (snooze)
  `$win.Add_Closing({
    if (-not (Test-Path -LiteralPath `$resp)) {
      Write-Response -Action "Snooze" -Minutes `$defaultSnooze
    }
  })

  [void]`$win.ShowDialog()
}
catch {
  exit 0
}
"@ | Set-Content -Path $scriptPath -Encoding UTF8 -Force

    return $scriptPath
}

# ===============================
# Scheduled Task creation (InteractiveToken) - no password required
# ===============================

function Ensure-ScheduledTasksModule {
    try {
        Import-Module ScheduledTasks -ErrorAction Stop
        return $true
    } catch {
        Write-Log "Failed to import ScheduledTasks module: $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Ensure-InteractiveUserTask {
    param(
        [Parameter(Mandatory)][string]$TaskName,
        [Parameter(Mandatory)][string]$UserName,   # DOMAIN\User from Win32_ComputerSystem.UserName
        [Parameter(Mandatory)][string]$ScriptPath,
        [Parameter(Mandatory)][string]$ArgName,    # -StatusFile OR -RequestFile
        [Parameter(Mandatory)][string]$ArgValue    # path to json file
    )

    if (-not (Ensure-ScheduledTasksModule)) { return $false }

    $conhost = Join-Path $env:SystemRoot "System32\conhost.exe"
    $ps      = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

    if (-not (Test-Path -LiteralPath $conhost)) {
        Write-Log "conhost.exe not found at expected path: $conhost" "WARN"
        return $false
    }
    if (-not (Test-Path -LiteralPath $ps)) {
        Write-Log "powershell.exe not found at expected path: $ps" "WARN"
        return $false
    }
    if (-not (Test-Path -LiteralPath $ScriptPath)) {
        Write-Log "Task script not found: $ScriptPath" "WARN"
        return $false
    }

    # conhost --headless "powershell.exe" -File "script.ps1" -ArgName "file.json"
    $args = "--headless `"$ps`" -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" $ArgName `"$ArgValue`""

    try {
        $action = New-ScheduledTaskAction -Execute $conhost -Argument $args

        # IMPORTANT: your box supports: LogonType Interactive; RunLevel Limited/Highest
        $principal = New-ScheduledTaskPrincipal -UserId $UserName -LogonType Interactive -RunLevel Limited

        # One-time far-future trigger (we run the task on-demand)
        $triggerAt = (Get-Date).Date.AddYears(10).AddHours(23).AddMinutes(59)
        $trigger   = New-ScheduledTaskTrigger -Once -At $triggerAt

        $settings  = New-ScheduledTaskSettingsSet `
            -Compatibility Win8 `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

        $task = New-ScheduledTask -Action $action -Principal $principal -Trigger $trigger -Settings $settings

        Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
        return $true
    }
    catch {
        Write-Log "Register-ScheduledTask failed for '$TaskName' (User='$UserName'): $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Start-UserTask {
    param([Parameter(Mandatory)][string]$TaskName)

    if (-not (Ensure-ScheduledTasksModule)) { return 9999 }

    # If task doesn't exist, return a non-zero code (and log)
    try {
        $null = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    } catch {
        Write-Log "Task '$TaskName' does not exist; cannot start." "WARN"
        return 2
    }

    try {
        Start-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        return 0
    } catch {
        Write-Log "Start-ScheduledTask failed for '$TaskName': $($_.Exception.Message)" "WARN"
        return 3
    }
}


function Ensure-PopupTaskForUser {
    param([Parameter(Mandatory)][string]$UserName)

    $taskName = "WinomatorPopup"
    [void](Ensure-PopupScript)

    $scriptPath = Join-Path $ToastDir "Show-Popup.ps1"
    $reqFile    = Join-Path $ToastDir "popup-request.json"

    return (Ensure-InteractiveUserTask -TaskName $taskName -UserName $UserName -ScriptPath $scriptPath -ArgName "-RequestFile" -ArgValue $reqFile)
}

function Run-PopupTask {
    param([string]$TaskName = "WinomatorPopup")
    return (Start-UserTask -TaskName $TaskName)
}

# ===============================
# Snooze helpers
# ===============================

function Get-SnoozeFileForApp {
    param([Parameter(Mandatory)][string]$WingetId)
    New-Item -ItemType Directory -Path $SnoozeDir -Force | Out-Null
    $token = Convert-ToSafeFileToken -Text $WingetId
    return Join-Path $SnoozeDir "snooze-$token.txt"
}

function Get-SnoozeUntil {
    param([Parameter(Mandatory)][string]$WingetId)
    $f = Get-SnoozeFileForApp -WingetId $WingetId
    if (-not (Test-Path -LiteralPath $f)) { return $null }
    try {
        $t = (Get-Content -LiteralPath $f -Raw).Trim()
        if ([string]::IsNullOrWhiteSpace($t)) { return $null }
        return [datetime]::Parse($t)
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
    $f = Get-SnoozeFileForApp -WingetId $WingetId
    $until.ToString("o") | Set-Content -LiteralPath $f -Encoding UTF8 -Force
    return $until
}

# ===============================
# Open-app detection + prompt
# ===============================

function Get-RunningProcessNamesForApp {
    param([Parameter(Mandatory)][string]$WingetId)

    if (-not $AppProcessMap.ContainsKey($WingetId)) { return @() }
    $names = @($AppProcessMap[$WingetId] | Where-Object { $_ -and $_.Trim() })
    if (-not $names -or $names.Count -eq 0) { return @() }

    $running = @()
    foreach ($n in $names) {
        $p = Get-Process -Name $n -ErrorAction SilentlyContinue
        if ($p) { $running += $n }
    }

    return @($running | Select-Object -Unique)
}

function Close-AppProcessesBestEffort {
    param([Parameter(Mandatory)][string[]]$ProcessNames)

    foreach ($n in $ProcessNames) {
        if ([string]::IsNullOrWhiteSpace($n)) { continue }

        # First: try graceful close for up to ~5 seconds
        $procs = @(Get-Process -Name $n -ErrorAction SilentlyContinue)
        foreach ($p in $procs) {
            try { $null = $p.CloseMainWindow() } catch {}
        }

        $deadline = (Get-Date).AddSeconds(5)
        while ((Get-Date) -lt $deadline) {
            if (-not (Get-Process -Name $n -ErrorAction SilentlyContinue)) { break }
            Start-Sleep -Milliseconds 300
        }

        # Then: force close any remaining
        $still = @(Get-Process -Name $n -ErrorAction SilentlyContinue)
        if ($still.Count -gt 0) {
            Write-Log "Forcing close for process '$n' (still running after graceful attempt)." "WARN"
            $still | Stop-Process -Force -ErrorAction SilentlyContinue
        } else {
            Write-Log "Closed process '$n' gracefully." "INFO"
        }
    }
}

function Prompt-UserForOpenApp {
    param(
        [Parameter(Mandatory)][pscustomobject]$App,
        [Parameter(Mandatory)][string[]]$ProcessNames
    )

    # returns [pscustomobject]@{ Action="Proceed|Snooze"; SnoozeMinutes=int }

    $user = Get-LoggedOnUser
    if ([string]::IsNullOrWhiteSpace($user) -or (-not $EnablePopups)) {
        # No user session -> can't prompt. Default snooze to avoid surprise close.
        return [pscustomobject]@{ Action="Snooze"; SnoozeMinutes=$DefaultSnoozeMinutes }
    }

    if (-not (Ensure-PopupTaskForUser -UserName $user)) {
        Write-Log "Popup task unavailable; defaulting to snooze $DefaultSnoozeMinutes minutes for $($App.Id)" "WARN"
        return [pscustomobject]@{ Action="Snooze"; SnoozeMinutes=$DefaultSnoozeMinutes }
    }

    New-Item -ItemType Directory -Path $ToastDir -Force | Out-Null
    $reqFile  = Join-Path $ToastDir "popup-request.json"
    $respFile = Join-Path $ToastDir "popup-response.json"
    if (Test-Path -LiteralPath $respFile) { Remove-Item -LiteralPath $respFile -Force -ErrorAction SilentlyContinue }

$procList = ($ProcessNames -join ", ")

$title = "Software update - '$($App.Name)' is currently open"

$body = @"
While trying to update your software, we detected that '$($App.Name)' is currently open.

Your computer has a pending update for this application. To safely install the update, the app needs to be closed.

What would you like to do?
- Proceed: Reveleer's IT tool will automatically close the app and install the update right now
- Snooze: Remind me later (you can choose the time before we try again)

Open application(s): $procList
"@


    $payload = [pscustomobject]@{
        Title         = $title
        Body          = $body
        ResponseFile  = $respFile
        SnoozeOptions = $PopupSnoozeOptionsMinutes
        Time          = (Get-Date).ToString("o")
    }
    $payload | ConvertTo-Json -Depth 6 | Set-Content -Path $reqFile -Encoding UTF8 -Force

    $startEc = Run-PopupTask -TaskName "WinomatorPopup"
    Write-Log "Popup task start requested. ExitCode=$startEc" "INFO"

    # Wait for response
    $deadline = (Get-Date).AddSeconds($PopupResponseTimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if (Test-Path -LiteralPath $respFile) { break }
        Start-Sleep -Milliseconds 400
    }

    if (-not (Test-Path -LiteralPath $respFile)) {
        Write-Log "User did not respond to popup in time. Snoozing $($App.Id) for $DefaultSnoozeMinutes minutes." "WARN"
        return [pscustomobject]@{ Action="Snooze"; SnoozeMinutes=$DefaultSnoozeMinutes }
    }

    try {
        $resp = Get-Content -LiteralPath $respFile -Raw | ConvertFrom-Json
        $action = [string]$resp.Action
        $mins   = 0
        try { $mins = [int]$resp.SnoozeMinutes } catch { $mins = $DefaultSnoozeMinutes }

        if ($action -eq "Proceed") {
            return [pscustomobject]@{ Action="Proceed"; SnoozeMinutes=0 }
        }

        if ($mins -lt 1) { $mins = $DefaultSnoozeMinutes }
        return [pscustomobject]@{ Action="Snooze"; SnoozeMinutes=$mins }
    }
    catch {
        Write-Log "Failed to parse popup response. Snoozing $($App.Id) for $DefaultSnoozeMinutes minutes." "WARN"
        return [pscustomobject]@{ Action="Snooze"; SnoozeMinutes=$DefaultSnoozeMinutes }
    }
}

# ===============================
# Parse winget update output (fixed-width)
# ===============================

function Get-WingetAvailableUpdates {
    param([Parameter(Mandatory)][string]$WingetPath)

    New-Item -ItemType Directory -Path $RootDir -Force | Out-Null
    $outFile = Join-Path $RootDir "winget-update.out.txt"
    $errFile = Join-Path $RootDir "winget-update.err.txt"

    $ec = Invoke-WingetToFiles -WingetPath $WingetPath -Arguments @(
        "update",
        "--source","winget",
        "--accept-source-agreements",
        "--disable-interactivity"
    ) -StdOutFile $outFile -StdErrFile $errFile

    if ($ec -eq 1460) {
        $prev = Get-MergedLogPreview -StdOutFile $outFile -StdErrFile $errFile -MaxLines 25
        throw "winget update timed out (possible prompt/hang). Preview: $prev"
    }

    $outLines = @()
    $errLines = @()
    if (Test-Path $outFile) { $outLines = Get-Content -Path $outFile -Encoding UTF8 -ErrorAction SilentlyContinue }
    if (Test-Path $errFile) { $errLines = Get-Content -Path $errFile -Encoding UTF8 -ErrorAction SilentlyContinue }

    $lines = @()
    $lines += $outLines
    $lines += $errLines
    $lines = $lines | ForEach-Object { Remove-Ansi "$_" }

    $clean = $lines | Where-Object {
        $_ -and
        ($_ -notmatch '^\s*Windows Package Manager') -and
        ($_ -notmatch '^\s*Copyright') -and
        ($_ -notmatch '^\s*Install the latest PowerShell') -and
        ($_ -notmatch '^\s*See https?:\/\/') -and
        ($_ -notmatch '^\s*$')
    }

    $clean = $clean | Where-Object {
        $t = $_.Trim()
        if ($t -match '^[\-\|\\/]+$') { return $false }
        if ($t -match '^\s*\d+(\.\d+)?\s*(KB|MB|GB)\s*/\s*\d+(\.\d+)?\s*(KB|MB|GB)\s*$') { return $false }
        if ($t -match '\b(KB|MB|GB)\s*/\s*\d') { return $false }
        if ($t -match '^(Downloading|Installing)\b') { return $false }
        if ($t -match 'msstore' -or $t -match 'Microsoft Store') { return $false }
        if ($t -match 'The\s+msstore\s+source\s+requires' ) { return $false }
        if ($t -match 'Do you agree to all the source agreements terms\?' ) { return $false }
        if ($t -match '^\[Y\]\s*Yes\s+\[N\]\s*No' ) { return $false }
        return $true
    }

    $joined = $clean -join "`n"
    if ($joined -match '(No applicable update found|No updates available)') { return @() }

    $headerIndex = -1
    for ($i = 0; $i -lt $clean.Count; $i++) {
        if ($clean[$i] -match '^\s*Name\s+Id\s+Version\s+Available(\s+Source)?\s*$') { $headerIndex = $i; break }
    }

    if ($clean.Count -eq 0) {
        $prev = Get-MergedLogPreview -StdOutFile $outFile -StdErrFile $errFile -MaxLines 30
        throw "winget returned no usable lines after filtering. Preview: $prev"
    }

    if ($headerIndex -lt 0) {
        $take = [Math]::Min(12, $clean.Count)
        $firstLines = ($clean | Select-Object -First $take) -join ' | '
        $prev = Get-MergedLogPreview -StdOutFile $outFile -StdErrFile $errFile -MaxLines 30
        throw "Could not find update header row. CleanFirstLines: $firstLines ; RawPreview: $prev"
    }


    $header = $clean[$headerIndex]
    $posName      = $header.IndexOf("Name")
    $posId        = $header.IndexOf("Id")
    $posVersion   = $header.IndexOf("Version")
    $posAvailable = $header.IndexOf("Available")
    $posSource    = $header.IndexOf("Source")

    if ($posName -lt 0 -or $posId -lt 0 -or $posVersion -lt 0 -or $posAvailable -lt 0) {
        throw "Header column positions not found."
    }

    $dataStart = $headerIndex + 1
    for ($j = $headerIndex + 1; $j -lt [Math]::Min($clean.Count, $headerIndex + 6); $j++) {
        if ($clean[$j] -match '^-{3,}') { $dataStart = $j + 1; break }
    }
    if ($dataStart -ge $clean.Count) { return @() }

    $rows = $clean | Select-Object -Skip $dataStart

    $results = @()

    foreach ($line in $rows) {
        if (-not $line.Trim()) { continue }
        if ($line -match '^-{3,}') { continue }
        if ($line -match '^\d+\s+upgrades?\s+available\.?$') { continue }

        $padTo = if ($posSource -gt 0) { $posSource + 10 } else { $posAvailable + 10 }
        $padded = $line.PadRight([Math]::Max($line.Length, $padTo))

        $name    = $padded.Substring($posName,    $posId - $posName).Trim()
        $id      = $padded.Substring($posId,      $posVersion - $posId).Trim()
        $version = $padded.Substring($posVersion, $posAvailable - $posVersion).Trim()

        $available = ""
        $source    = ""
        if ($posSource -gt 0) {
            $available = $padded.Substring($posAvailable, $posSource - $posAvailable).Trim()
            $source    = $padded.Substring($posSource).Trim()
        } else {
            $available = $padded.Substring($posAvailable).Trim()
        }

        if ([string]::IsNullOrWhiteSpace($id)) { continue }

        $results += [PSCustomObject]@{
            Name      = $name
            Id        = $id
            Version   = $version
            Available = $available
            Source    = $source
        }
    }

    return $results
}

function Is-Excluded {
    param([Parameter(Mandatory)][pscustomobject]$App)

    if ($ExcludedWingetIds -contains $App.Id) { return $true }

    foreach ($pat in $ExcludedNamePatterns) {
        if ([string]::IsNullOrWhiteSpace($pat)) { continue }
        if ($App.Name -match $pat) { return $true }
    }

    return $false
}

# ===============================
# Main
# ===============================

Write-Log "Starting weekly Winomator run (SYSTEM)"

$winget = Ensure-WingetForSystem
if (-not $winget) {
    Write-Log "winget.exe not found. App Installer missing or not available." "ERROR"
    exit 1
}
Write-Log "Using winget at: $winget"

$wt = Test-WingetAsSystem -WingetPath $winget
Write-Log "winget self-test: Ok=$($wt.Ok) ExitCode=$($wt.ExitCode) Out='$($wt.StdOut)'" "INFO"

if (-not $wt.Ok) {
    $hex = ""
    try { $hex = ('{0:X8}' -f ([uint32]($LASTEXITCODE -band 0xFFFFFFFF))) } catch {}
    Write-Log "winget self-test failed under SYSTEM. ExitCode=$($wt.ExitCode) $hex; OutBytes=$($wt.OutBytes); ErrBytes=$($wt.ErrBytes)" "ERROR"
    exit 2
}

try {
    $updates = @(Get-WingetAvailableUpdates -WingetPath $winget)
}
catch {
    $msg = $_.Exception.Message
    Write-Log "Failed to query winget updates: $msg" "ERROR"
    Send-WinomatorToast -Title "Software Update" -Body "Update scan failed. See logs for details." -Type Error
    exit 2
}

if (-not $updates -or $updates.Count -eq 0) {
    Write-Log "No applicable updates found."
    Send-WinomatorToast -Title "Software Update" -Body "No application updates were found this run." -Type Info
    exit 0
}

Write-Log "Found $($updates.Count) updates available."
foreach ($u in $updates) {
    Write-Log "Update detected: Name='$($u.Name)' Id='$($u.Id)' $($u.Version) -> $($u.Available) Source=$($u.Source)"
}

$targets = @($updates | Where-Object { -not (Is-Excluded -App $_) })

if (-not $targets -or $targets.Count -eq 0) {
    Write-Log "All available updates are excluded. Nothing to do."
    #Send-WinomatorToast -Title "Software Update" -Body "Updates were found, but all were excluded by policy." -Type Info
    exit 0
}

Write-Log "Upgrading $($targets.Count) packages after exclusions..."
Send-WinomatorToast -Title "Software updates starting" -Body "Installing $($targets.Count) application update(s) in the background. You may notice apps close briefly." -Type Info

$success = 0
$failed  = 0
$failedIds = @()

foreach ($app in $targets) {
    if ([string]::IsNullOrWhiteSpace($app.Id)) {
        Write-Log "Skipping row with empty Id (Name='$($app.Name)')" "WARN"
        continue
    }

    # Snooze check
    $until = Get-SnoozeUntil -WingetId $app.Id
    if ($until -and (Get-Date) -lt $until) {
        Write-Log "Skipping $($app.Id) due to snooze until $($until.ToString('o'))" "INFO"
        continue
    }

    # Open-app detection -> prompt
    $running = Get-RunningProcessNamesForApp -WingetId $app.Id
    if ($running.Count -gt 0) {
        Write-Log "Detected running app for $($app.Id): $($running -join ', ')" "INFO"
        $resp = Prompt-UserForOpenApp -App $app -ProcessNames $running

        if ($resp.Action -eq "Snooze") {
            $until2 = Set-SnoozeUntil -WingetId $app.Id -Minutes ([int]$resp.SnoozeMinutes)
            Write-Log "User snoozed $($app.Id) for $($resp.SnoozeMinutes) minutes (until $($until2.ToString('o')))" "INFO"
            Send-WinomatorToast -Title "Software Update" -Body "Snoozed '$($app.Name)' update for $($resp.SnoozeMinutes) minutes." -Type Warning
            continue
        }

        Write-Log "User chose Proceed; closing processes for $($app.Id)..." "INFO"
        Close-AppProcessesBestEffort -ProcessNames $running
    }

    Write-Log "Upgrading $($app.Name) ($($app.Id)) $($app.Version) -> $($app.Available)"

    $token = Convert-ToSafeFileToken -Text $app.Id
    $outFile  = Join-Path $RootDir "upgrade-$token.out.txt"
    $errFile  = Join-Path $RootDir "upgrade-$token.err.txt"
    $outFile2 = Join-Path $RootDir "upgrade-$token.retry.out.txt"
    $errFile2 = Join-Path $RootDir "upgrade-$token.retry.err.txt"

    $baseArgs = @(
        "upgrade",
        "--id", $app.Id,
        "--exact",
        "--source", "winget",
        "--accept-source-agreements",
        "--accept-package-agreements",
        "--disable-interactivity"
    )

    $code = Invoke-WingetToFiles -WingetPath $winget -Arguments ($baseArgs + @("--silent")) -StdOutFile $outFile -StdErrFile $errFile
    $preview = Get-MergedLogPreview -StdOutFile $outFile -StdErrFile $errFile -MaxLines 40

    # 1) Normal success
    if ($code -eq 0) {
        $success++
        Write-Log "Success: $($app.Name)"
        continue
    }

    # 2) Winget says there is nothing to do -> treat as success and DO NOT retry
    # This is the "first run succeeded but exit code handling/version detection is flaky" situation.
    if (Test-WingetOutputIndicatesSuccess -PreviewOrText $preview) {
        $success++
        Write-Log "Success (winget output indicates completion/no-op): $($app.Name). Output: $preview" "INFO"
        continue
    }

    # 3) If 9997 AND we still didn't see a success/no-op marker, attempt version verify
    if ($code -eq 9997) {
        $installedNow = Get-InstalledWingetVersion -WingetPath $winget -Id $app.Id
        $available    = [string]$app.Available
        $installedWas = [string]$app.Version

        $verified = $false
        if (-not [string]::IsNullOrWhiteSpace($installedNow)) {
            if (-not [string]::IsNullOrWhiteSpace($available)) {
                if ($installedNow -eq $available -or $installedNow.StartsWith($available)) { $verified = $true }
            } elseif (-not [string]::IsNullOrWhiteSpace($installedWas) -and $installedNow -ne $installedWas) {
                $verified = $true
            }
        }

        if ($verified) {
            $success++
            Write-Log "Success (verified after unknown exit code): $($app.Name) now at $installedNow" "INFO"
            continue
        }

        Write-Log "Unknown exit code (9997) and not verifiable; will consider retry. Preview: $preview" "WARN"
    }

    # If the output clearly indicates a "hard failure", capture the real embedded code for reporting
    $embedded = Get-WingetEmbeddedExitCode -Text $preview
    $reported = if ($embedded -ne $null) { $embedded } else { $code }

    # Decide if we should skip retry (known no-win scenarios or "retry later" scenarios)
    $noRetryReason = Get-WingetNoRetryReason -Text $preview
    $hardFail      = Test-WingetOutputIndicatesHardFailure -PreviewOrText $preview

    Write-Log "FAILED ($reported) on silent attempt: $($app.Name). Output: $preview" "WARN"

    if ($noRetryReason) {
        $failed++
        $failedIds += $app.Id
        Write-Log "Skipping retry for $($app.Id): $noRetryReason" "WARN"
        continue
    }

    # If it's a hard failure with a real embedded installer code like 1603, retrying immediately usually won't help.
    # You *can* still keep your single retry for safety, but this makes it configurable logic.
    if ($hardFail -and $embedded -ne $null -and $embedded -ne 1618) {
        # Optional: comment this out if you still want the retry always.
        $failed++
        $failedIds += $app.Id
        Write-Log "Skipping retry for $($app.Id): hard installer failure (embedded=$embedded)" "WARN"
        continue
    }

    # Otherwise: do your existing retry (non-silent)
    $code2 = Invoke-WingetToFiles -WingetPath $winget -Arguments $baseArgs -StdOutFile $outFile2 -StdErrFile $errFile2
    $preview2 = Get-MergedLogPreview -StdOutFile $outFile2 -StdErrFile $errFile2 -MaxLines 60

    # If retry output indicates success or no-op, treat as success (and stop)
    if ($code2 -eq 0 -or (Test-WingetOutputIndicatesSuccess -PreviewOrText $preview2)) {
        $success++
        Write-Log "Success on retry (output indicates completion/no-op): $($app.Name). Output: $preview2" "INFO"
        continue
    }

    $failed++
    $failedIds += $app.Id
    $embedded2 = Get-WingetEmbeddedExitCode -Text $preview2
    $reported2 = if ($embedded2 -ne $null) { $embedded2 } else { $code2 }

    Write-Log "FAILED ($reported2) on retry: $($app.Name). Output: $preview2" "WARN"
    Write-Log "Saved detailed logs: $outFile2 and $errFile2" "WARN"

    if ($preview2 -match '(?i)exit\s*code:\s*1605') {
        Write-Log "MSI 1605 suggests per-user install / not visible to SYSTEM. Consider excluding '$($app.Id)' or deploying a machine-wide installer." "WARN"
    }

}

Write-Log "Winomator run completed. Success=$success Failed=$failed"

if ($failed -eq 0) {
    Send-WinomatorToast -Title "Software updates completed" -Body "Installed $success update(s). No failures." -Type Success
} else {
    $uniq = @($failedIds | Select-Object -Unique)
    $failList = ($uniq | Select-Object -First 5) -join ", "
    $more = ""
    if ($uniq.Count -gt 5) { $more = " (more...)" }
    Send-WinomatorToast -Title "Software updates completed" -Body "Installed $success update(s). $failed failed: $failList$more" -Type Error
}

exit 0