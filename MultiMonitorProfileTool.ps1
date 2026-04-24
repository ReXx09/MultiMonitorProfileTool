param(
    [string]$ConfigPath = "$PSScriptRoot\monitor-profiles.json"
)

# Echter Dot-Source nur wenn InvocationName '.' ist UND PSCommandPath nicht gesetzt
# (VS Code Debug-Adapter setzt PSCommandPath, obwohl es wie Dot-Source aussieht)
$script:IsDotSourced = ($MyInvocation.InvocationName -eq '.') -and [string]::IsNullOrWhiteSpace($PSCommandPath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not $script:IsDotSourced -and [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Host "Neustart in STA-Modus..."
    $argsList = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-STA',
        '-File', "`"$PSCommandPath`"",
        '-ConfigPath', "`"$ConfigPath`""
    )
    Start-Process -FilePath 'powershell.exe' -ArgumentList $argsList | Out-Null
    return
}

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$signature = @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public static class MultiMonitorWinApi {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct DISPLAY_DEVICE {
        public int cb;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string DeviceName;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceString;

        public int StateFlags;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceID;

        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public string DeviceKey;
    }

    [DllImport("user32.dll")]
    public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    public static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", SetLastError = true)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    [DllImport("user32.dll")]
    public static extern uint GetDpiForWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern IntPtr SetProcessDpiAwarenessContext(IntPtr dpiContext);

    [DllImport("shcore.dll")]
    public static extern int SetProcessDpiAwareness(int awareness);

    [DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();

    [DllImport("user32.dll")]
    public static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern bool EnumDisplayDevices(string lpDevice, uint iDevNum, ref DISPLAY_DEVICE lpDisplayDevice, uint dwFlags);

    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
}
"@
if (-not ('MultiMonitorWinApi' -as [type])) {
    Add-Type -TypeDefinition $signature -Language CSharp
}

function Enable-PerMonitorDpiAwareness {
    # Prefer Per-Monitor V2; fallback to older APIs for compatibility.
    $applied = $false

    try {
        $dpiAwarenessContextPerMonitorV2 = [IntPtr]::new(-4)
        $ctxResult = [MultiMonitorWinApi]::SetProcessDpiAwarenessContext($dpiAwarenessContextPerMonitorV2)
        if ($ctxResult -ne [IntPtr]::Zero) {
            $applied = $true
        }
    }
    catch {
    }

    if (-not $applied) {
        try {
            $hr = [MultiMonitorWinApi]::SetProcessDpiAwareness(2)
            if ($hr -eq 0 -or $hr -eq -2147023649) {
                $applied = $true
            }
        }
        catch {
        }
    }

    if (-not $applied) {
        try {
            [void][MultiMonitorWinApi]::SetProcessDPIAware()
        }
        catch {
        }
    }
}

Enable-PerMonitorDpiAwareness

$script:monitorVendorMap = @{
    ACI = 'ASUS'
    ACR = 'Acer'
    AOC = 'AOC'
    AUS = 'ASUS'
    BNQ = 'BenQ'
    DEL = 'Dell'
    DEN = 'Denon'
    GSM = 'LG'
    HWP = 'HP'
    IVM = 'Iiyama'
    LEN = 'Lenovo'
    LPL = 'LG Philips'
    MSI = 'MSI'
    PHL = 'Philips'
    SAM = 'Samsung'
    SEC = 'Samsung'
    SNY = 'Sony'
    VSC = 'ViewSonic'
}

$script:monitorMetadataCache = $null

function Convert-UInt16ArrayToText {
    param([object]$Value)

    if (-not $Value) { return '' }

    $chars = foreach ($item in $Value) {
        $number = [int]$item
        if ($number -gt 0) { [char]$number }
    }

    return (-join $chars).Trim()
}

function Resolve-MonitorVendorName {
    param([string]$VendorCode)

    if ([string]::IsNullOrWhiteSpace($VendorCode)) { return '' }
    $upper = $VendorCode.ToUpperInvariant()
    if ($script:monitorVendorMap.ContainsKey($upper)) {
        return $script:monitorVendorMap[$upper]
    }
    return $upper
}

function Get-MonitorMetadataCache {
    if ($script:monitorMetadataCache) {
        return $script:monitorMetadataCache
    }

    $cache = @{}
    try {
        $items = @(Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorID -ErrorAction Stop)
    }
    catch {
        $items = @()
    }

    foreach ($item in $items) {
        $instanceName = [string]$item.InstanceName
        $monitorCode = ''
        if ($instanceName -match '^DISPLAY\\([^\\]+)\\') {
            $monitorCode = [string]$matches[1]
        }

        $manufacturerCode = Convert-UInt16ArrayToText -Value $item.ManufacturerName
        if ([string]::IsNullOrWhiteSpace($manufacturerCode) -and $monitorCode.Length -ge 3) {
            $manufacturerCode = $monitorCode.Substring(0, 3)
        }

        $entry = [pscustomobject]@{
            MonitorCode = $monitorCode
            ManufacturerCode = $manufacturerCode
            ManufacturerName = Resolve-MonitorVendorName -VendorCode $manufacturerCode
            ModelName = Convert-UInt16ArrayToText -Value $item.UserFriendlyName
            SerialNumber = Convert-UInt16ArrayToText -Value $item.SerialNumberID
        }

        if (-not [string]::IsNullOrWhiteSpace($monitorCode)) {
            $cache[$monitorCode] = $entry
        }
    }

    $script:monitorMetadataCache = $cache
    return $cache
}

function Get-DisplayDeviceInfo {
    param([Parameter(Mandatory)][string]$ScreenDeviceName)

    $monitorDevice = $null
    for ($index = 0; $index -lt 10; $index++) {
        $dd = New-Object MultiMonitorWinApi+DISPLAY_DEVICE
        $dd.cb = [Runtime.InteropServices.Marshal]::SizeOf([type][MultiMonitorWinApi+DISPLAY_DEVICE])

        if (-not [MultiMonitorWinApi]::EnumDisplayDevices($ScreenDeviceName, [uint32]$index, [ref]$dd, [uint32]0)) {
            break
        }

        if (-not [string]::IsNullOrWhiteSpace([string]$dd.DeviceID) -and [string]$dd.DeviceID -match '^MONITOR\\') {
            $monitorDevice = $dd
            break
        }
    }

    if (-not $monitorDevice) {
        return [pscustomobject]@{
            MonitorCode = ''
            DeviceString = ''
            DeviceId = ''
        }
    }

    $monitorCode = ''
    if ([string]$monitorDevice.DeviceID -match '^MONITOR\\([^\\]+)\\') {
        $monitorCode = [string]$matches[1]
    }

    return [pscustomobject]@{
        MonitorCode = $monitorCode
        DeviceString = [string]$monitorDevice.DeviceString
        DeviceId = [string]$monitorDevice.DeviceID
    }
}

function Get-FriendlyMonitorInfo {
    param([Parameter(Mandatory)]$Screen)

    $displayInfo = Get-DisplayDeviceInfo -ScreenDeviceName $Screen.DeviceName
    $metadataCache = Get-MonitorMetadataCache
    $meta = $null
    if (-not [string]::IsNullOrWhiteSpace([string]$displayInfo.MonitorCode) -and $metadataCache.ContainsKey([string]$displayInfo.MonitorCode)) {
        $meta = $metadataCache[[string]$displayInfo.MonitorCode]
    }

    $manufacturer = ''
    $model = ''

    if ($meta) {
        $manufacturer = [string]$meta.ManufacturerName
        $model = [string]$meta.ModelName
    }

    if ([string]::IsNullOrWhiteSpace($model)) {
        $deviceString = [string]$displayInfo.DeviceString
        if (-not [string]::IsNullOrWhiteSpace($deviceString) -and $deviceString -notmatch '^Generic PnP') {
            $model = $deviceString
        }
    }

    if ([string]::IsNullOrWhiteSpace($manufacturer) -and -not [string]::IsNullOrWhiteSpace([string]$displayInfo.MonitorCode) -and $displayInfo.MonitorCode.Length -ge 3) {
        $manufacturer = Resolve-MonitorVendorName -VendorCode $displayInfo.MonitorCode.Substring(0, 3)
    }

    $friendly = ''
    if (-not [string]::IsNullOrWhiteSpace($manufacturer) -and -not [string]::IsNullOrWhiteSpace($model)) {
        if ($model.ToUpperInvariant().StartsWith($manufacturer.ToUpperInvariant())) {
            $friendly = $model
        }
        else {
            $friendly = "$manufacturer $model"
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($model)) {
        $friendly = $model
    }
    elseif (-not [string]::IsNullOrWhiteSpace($manufacturer)) {
        $friendly = $manufacturer
    }
    else {
        $friendly = $Screen.DeviceName
    }

    $roleName = 'Monitor'
    $compareText = "$friendly $manufacturer $model".ToLowerInvariant()
    if ($compareText -match 'denon|projector|beamer|avr') {
        $roleName = 'Beamer/AV'
    }
    elseif ($Screen.Primary) {
        $roleName = 'Hauptmonitor'
    }

    $displaySlot = ($Screen.DeviceName -replace '^\\\\\.\\', '')

    return [pscustomobject]@{
        ManufacturerName = $manufacturer
        ModelName = $model
        FriendlyName = $friendly
        DisplaySlot = $displaySlot
        RoleName = $roleName
        MonitorCode = [string]$displayInfo.MonitorCode
    }
}

function New-DefaultConfig {
    [ordered]@{
        Profiles = @(
            [ordered]@{
                Name = 'Alltag'
                DisplayMode = 'internal'
                WindowLayouts = @()
            },
            [ordered]@{
                Name = 'StreamingGaming'
                DisplayMode = 'extend'
                WindowLayouts = @()
            }
        )
        Settings = [ordered]@{
            RestoreAfterSwitch = $true
            SwitchDelayMs = 2500
            AutoLaunchMissingWindows = $true
            LaunchDelayMs = 1800
            UiLanguage = 'de'
            DebugLoggingEnabled = $false
            DebugLogPath = "$PSScriptRoot\monitor-debug.log"
            MaxWindowsPerCapture = 80
            ExcludedProcesses = @('dwm','explorer','ShellExperienceHost','StartMenuExperienceHost','ApplicationFrameHost','SearchHost','TextInputHost','LockApp')
        }
    }
}

function Load-Config {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        $cfg = New-DefaultConfig
        Save-Config -Config $cfg -Path $Path
        return $cfg
    }

    try {
        $raw = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        if ([string]::IsNullOrWhiteSpace($raw)) {
            $cfg = New-DefaultConfig
            Save-Config -Config $cfg -Path $Path
            return $cfg
        }

        $cfg = $raw | ConvertFrom-Json

        if (-not $cfg.Profiles -or -not $cfg.Settings) {
            throw 'Konfiguration unvollstaendig.'
        }

        return $cfg
    }
    catch {
        [System.Windows.MessageBox]::Show("Fehler beim Laden der Konfiguration:`n$($_.Exception.Message)", 'Konfigurationsfehler', 'OK', 'Error') | Out-Null
        $backup = "$Path.corrupt.$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
        try { Copy-Item -LiteralPath $Path -Destination $backup -ErrorAction SilentlyContinue } catch {}
        $cfg = New-DefaultConfig
        Save-Config -Config $cfg -Path $Path
        return $cfg
    }
}

function Save-Config {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Path
    )

    $json = $Config | ConvertTo-Json -Depth 8
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    Set-Content -LiteralPath $Path -Value $json -Encoding UTF8
}

function Get-DebugLogPath {
    param([Parameter(Mandatory)]$Config)

    $configuredPath = [string](Get-SafePropertyValue -Object $Config.Settings -Name 'DebugLogPath' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($configuredPath)) {
        return (Join-Path $PSScriptRoot 'monitor-debug.log')
    }

    return [Environment]::ExpandEnvironmentVariables($configuredPath)
}

function Write-DebugLog {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Message,
        [string]$Scope = 'General'
    )

    $enabled = [bool](Get-SafePropertyValue -Object $Config.Settings -Name 'DebugLoggingEnabled' -DefaultValue $false)
    if (-not $enabled) { return }

    $path = Get-DebugLogPath -Config $Config
    if ([string]::IsNullOrWhiteSpace($path)) { return }

    try {
        $dir = Split-Path -Path $path -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -Path $dir -ItemType Directory -Force | Out-Null
        }

        $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
        Add-Content -LiteralPath $path -Value ("[{0}] [{1}] {2}" -f $timestamp, $Scope, $Message) -Encoding UTF8
    }
    catch {
    }
}

function Resolve-ScreenByRect {
    param(
        [Parameter(Mandatory)][int]$Left,
        [Parameter(Mandatory)][int]$Top,
        [Parameter(Mandatory)][int]$Width,
        [Parameter(Mandatory)][int]$Height
    )

    $screens = @(Get-Screens)
    if ($screens.Count -eq 0) { return $null }

    $right = $Left + $Width
    $bottom = $Top + $Height
    $centerX = $Left + [Math]::Floor($Width / 2)
    $centerY = $Top + [Math]::Floor($Height / 2)

    $centerMatch = $screens |
        Where-Object {
            $centerX -ge $_.BoundsX -and
            $centerX -lt ($_.BoundsX + $_.BoundsWidth) -and
            $centerY -ge $_.BoundsY -and
            $centerY -lt ($_.BoundsY + $_.BoundsHeight)
        } |
        Select-Object -First 1
    if ($centerMatch) { return $centerMatch }

    $best = $null
    $bestArea = -1
    foreach ($screen in $screens) {
        $sLeft = [int]$screen.BoundsX
        $sTop = [int]$screen.BoundsY
        $sRight = $sLeft + [int]$screen.BoundsWidth
        $sBottom = $sTop + [int]$screen.BoundsHeight

        $ix = [Math]::Max($Left, $sLeft)
        $iy = [Math]::Max($Top, $sTop)
        $ax = [Math]::Min($right, $sRight)
        $ay = [Math]::Min($bottom, $sBottom)

        $iw = [Math]::Max(0, $ax - $ix)
        $ih = [Math]::Max(0, $ay - $iy)
        $area = $iw * $ih

        if ($area -gt $bestArea) {
            $bestArea = $area
            $best = $screen
        }
    }

    return $best
}

function Get-WindowRectSnapshot {
    param([Parameter(Mandatory)][long]$Handle)

    $rect = New-Object MultiMonitorWinApi+RECT
    if (-not [MultiMonitorWinApi]::GetWindowRect([IntPtr]::new([int64]$Handle), [ref]$rect)) {
        return $null
    }

    $width = $rect.Right - $rect.Left
    $height = $rect.Bottom - $rect.Top

    $dpi = 0
    try {
        $dpi = [int][MultiMonitorWinApi]::GetDpiForWindow([IntPtr]::new([int64]$Handle))
    }
    catch {
        $dpi = 0
    }

    $screen = Resolve-ScreenByRect -Left $rect.Left -Top $rect.Top -Width $width -Height $height

    return [pscustomobject]@{
        Left = $rect.Left
        Top = $rect.Top
        Width = $width
        Height = $height
        Dpi = $dpi
        ScreenDeviceName = if ($screen) { [string]$screen.DeviceName } else { '' }
        ScreenDisplayName = if ($screen) { [string]$screen.DisplayName } else { '' }
    }
}

function Get-RelativeRectOnScreen {
    param(
        [Parameter(Mandatory)]$ScreenItem,
        [Parameter(Mandatory)][int]$Left,
        [Parameter(Mandatory)][int]$Top,
        [Parameter(Mandatory)][int]$Width,
        [Parameter(Mandatory)][int]$Height
    )

    $baseX = [double]$ScreenItem.WorkingAreaX
    $baseY = [double]$ScreenItem.WorkingAreaY
    $baseW = [double]$ScreenItem.WorkingAreaWidth
    $baseH = [double]$ScreenItem.WorkingAreaHeight
    if ($baseW -le 0 -or $baseH -le 0) {
        return $null
    }

    return [pscustomobject]@{
        RelativeLeft = ([double]$Left - $baseX) / $baseW
        RelativeTop = ([double]$Top - $baseY) / $baseH
        RelativeWidth = [double]$Width / $baseW
        RelativeHeight = [double]$Height / $baseH
    }
}

function Get-AbsoluteRectFromRelative {
    param(
        [Parameter(Mandatory)]$ScreenItem,
        [Parameter(Mandatory)][double]$RelativeLeft,
        [Parameter(Mandatory)][double]$RelativeTop,
        [Parameter(Mandatory)][double]$RelativeWidth,
        [Parameter(Mandatory)][double]$RelativeHeight
    )

    $baseX = [double]$ScreenItem.WorkingAreaX
    $baseY = [double]$ScreenItem.WorkingAreaY
    $baseW = [double]$ScreenItem.WorkingAreaWidth
    $baseH = [double]$ScreenItem.WorkingAreaHeight
    if ($baseW -le 0 -or $baseH -le 0) {
        return $null
    }

    $left = [int][Math]::Round($baseX + ($RelativeLeft * $baseW))
    $top = [int][Math]::Round($baseY + ($RelativeTop * $baseH))
    $width = [int][Math]::Round($RelativeWidth * $baseW)
    $height = [int][Math]::Round($RelativeHeight * $baseH)

    if ($width -lt 120) { $width = 120 }
    if ($height -lt 80) { $height = 80 }

    return [ordered]@{
        Left = $left
        Top = $top
        Width = $width
        Height = $height
    }
}

function Get-Screens {
    $screens = [System.Windows.Forms.Screen]::AllScreens
    foreach ($s in $screens) {
        $friendlyInfo = Get-FriendlyMonitorInfo -Screen $s
        [pscustomobject]@{
            DeviceName = $s.DeviceName
            DisplayName = [string]$friendlyInfo.FriendlyName
            DisplayLabel = "{0} ({1})" -f $friendlyInfo.FriendlyName, $friendlyInfo.DisplaySlot
            ManufacturerName = [string]$friendlyInfo.ManufacturerName
            ModelName = [string]$friendlyInfo.ModelName
            RoleName = [string]$friendlyInfo.RoleName
            DisplaySlot = [string]$friendlyInfo.DisplaySlot
            MonitorCode = [string]$friendlyInfo.MonitorCode
            Primary = $s.Primary
            Bounds = $s.Bounds
            WorkingArea = $s.WorkingArea
            BoundsX = $s.Bounds.X
            BoundsY = $s.Bounds.Y
            BoundsWidth = $s.Bounds.Width
            BoundsHeight = $s.Bounds.Height
            WorkingAreaX = $s.WorkingArea.X
            WorkingAreaY = $s.WorkingArea.Y
            WorkingAreaWidth = $s.WorkingArea.Width
            WorkingAreaHeight = $s.WorkingArea.Height
        }
    }
}

function Split-CommandLineExecutableAndArguments {
    param([string]$CommandLine)

    $line = [string]$CommandLine
    if ([string]::IsNullOrWhiteSpace($line)) {
        return [pscustomobject]@{ ExecutablePath = ''; Arguments = '' }
    }

    $line = $line.Trim()
    if ($line.StartsWith('"')) {
        $closingQuote = $line.IndexOf('"', 1)
        if ($closingQuote -gt 1) {
            $exe = $line.Substring(1, $closingQuote - 1)
            $args = $line.Substring($closingQuote + 1).Trim()
            return [pscustomobject]@{ ExecutablePath = $exe; Arguments = $args }
        }
    }

    $firstSpace = $line.IndexOf(' ')
    if ($firstSpace -gt 0) {
        return [pscustomobject]@{
            ExecutablePath = $line.Substring(0, $firstSpace).Trim()
            Arguments = $line.Substring($firstSpace + 1).Trim()
        }
    }

    return [pscustomobject]@{ ExecutablePath = $line; Arguments = '' }
}

function Resolve-VSCodeWorkspacePathFromTitle {
    param([string]$Title)

    $windowTitle = [string]$Title
    if ([string]::IsNullOrWhiteSpace($windowTitle)) { return '' }

    $match = [regex]::Match($windowTitle, ' - (.+) - Visual Studio Code$')
    if (-not $match.Success) { return '' }

    $folderName = [string]$match.Groups[1].Value.Trim()
    if ([string]::IsNullOrWhiteSpace($folderName)) { return '' }

    if (Test-Path -LiteralPath $folderName -PathType Container) {
        return (Resolve-Path -LiteralPath $folderName).Path
    }

    $roots = @()
    if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $roots += $PSScriptRoot
        $parent = Split-Path -Path $PSScriptRoot -Parent
        if (-not [string]::IsNullOrWhiteSpace($parent)) {
            $roots += $parent
        }
    }

    $desktop = [Environment]::GetFolderPath('Desktop')
    if (-not [string]::IsNullOrWhiteSpace($desktop)) { $roots += $desktop }
    $documents = [Environment]::GetFolderPath('MyDocuments')
    if (-not [string]::IsNullOrWhiteSpace($documents)) { $roots += $documents }

    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($root in @($roots | Select-Object -Unique)) {
        if ([string]::IsNullOrWhiteSpace([string]$root)) { continue }
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }

        $direct = Join-Path -Path $root -ChildPath $folderName
        if (Test-Path -LiteralPath $direct -PathType Container) {
            $candidates.Add((Resolve-Path -LiteralPath $direct).Path)
        }

        foreach ($child in @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue)) {
            $nested = Join-Path -Path $child.FullName -ChildPath $folderName
            if (Test-Path -LiteralPath $nested -PathType Container) {
                $candidates.Add((Resolve-Path -LiteralPath $nested).Path)
            }
        }

        foreach ($hit in @(Get-ChildItem -LiteralPath $root -Directory -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $folderName } | Select-Object -First 6)) {
            $candidates.Add($hit.FullName)
        }
    }

    $unique = @($candidates | Select-Object -Unique)
    if ($unique.Count -ge 1) {
        return [string]$unique[0]
    }

    $titleParts = [regex]::Match($windowTitle, '^(.+?) - (.+) - Visual Studio Code$')
    if ($titleParts.Success) {
        $fileName = [string]$titleParts.Groups[1].Value.Trim()
        $workspaceHint = [string]$titleParts.Groups[2].Value.Trim()

        if (-not [string]::IsNullOrWhiteSpace($fileName) -and $fileName.Contains('.')) {
            $fileCandidates = New-Object System.Collections.Generic.List[string]
            foreach ($root in @($roots | Select-Object -Unique)) {
                if ([string]::IsNullOrWhiteSpace([string]$root)) { continue }
                if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }

                foreach ($fileHit in @(Get-ChildItem -LiteralPath $root -File -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $fileName } | Select-Object -First 8)) {
                    $parentDir = Split-Path -Path $fileHit.FullName -Parent
                    if (-not [string]::IsNullOrWhiteSpace($workspaceHint) -and $parentDir -match [regex]::Escape($workspaceHint)) {
                        $fileCandidates.Add($parentDir)
                    }
                    else {
                        $fileCandidates.Add($parentDir)
                    }
                }
            }

            $fileUnique = @($fileCandidates | Select-Object -Unique)
            if ($fileUnique.Count -ge 1) {
                return [string]$fileUnique[0]
            }
        }
    }

    return ''
}

function Get-VisibleTopLevelWindows {
    param([string[]]$ExcludedProcesses)

    $list = New-Object System.Collections.Generic.List[object]
    $commandLineCache = @{}

    $callback = [MultiMonitorWinApi+EnumWindowsProc]{
        param([IntPtr]$hWnd, [IntPtr]$lParam)

        if (-not [MultiMonitorWinApi]::IsWindowVisible($hWnd)) { return $true }
        if ([MultiMonitorWinApi]::IsIconic($hWnd)) { return $true }

        $sb = New-Object System.Text.StringBuilder 512
        [void][MultiMonitorWinApi]::GetWindowText($hWnd, $sb, $sb.Capacity)
        $title = $sb.ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($title)) { return $true }

        [uint32]$windowProcessId = 0
        [void][MultiMonitorWinApi]::GetWindowThreadProcessId($hWnd, [ref]$windowProcessId)

        try {
            $proc = Get-Process -Id ([int]$windowProcessId) -ErrorAction Stop
        }
        catch {
            return $true
        }

        $executablePath = ''
        try {
            $executablePath = [string]$proc.MainModule.FileName
        }
        catch {
            try {
                $executablePath = [string]$proc.Path
            }
            catch {
                $executablePath = ''
            }
        }

        $commandLine = ''
        if ($commandLineCache.ContainsKey([int]$proc.Id)) {
            $commandLine = [string]$commandLineCache[[int]$proc.Id]
        }
        else {
            try {
                $cim = Get-CimInstance -ClassName Win32_Process -Filter "ProcessId = $($proc.Id)" -ErrorAction Stop
                $commandLine = [string]$cim.CommandLine
            }
            catch {
                $commandLine = ''
            }
            $commandLineCache[[int]$proc.Id] = $commandLine
        }

        $launchArguments = ''
        if (-not [string]::IsNullOrWhiteSpace($commandLine)) {
            $launchInfo = Split-CommandLineExecutableAndArguments -CommandLine $commandLine
            if ([string]::IsNullOrWhiteSpace($executablePath) -and -not [string]::IsNullOrWhiteSpace([string]$launchInfo.ExecutablePath)) {
                $executablePath = [string]$launchInfo.ExecutablePath
            }
            $launchArguments = [string]$launchInfo.Arguments
        }

        if ($ExcludedProcesses -contains $proc.ProcessName) { return $true }

        $rect = New-Object MultiMonitorWinApi+RECT
        if (-not [MultiMonitorWinApi]::GetWindowRect($hWnd, [ref]$rect)) { return $true }

        $width = $rect.Right - $rect.Left
        $height = $rect.Bottom - $rect.Top
        if ($width -lt 120 -or $height -lt 80) { return $true }

        $item = [pscustomobject]@{
            Handle = $hWnd.ToInt64()
            ProcessName = $proc.ProcessName
            ProcessId = $proc.Id
            ExecutablePath = $executablePath
            LaunchArguments = $launchArguments
            Title = $title
            Left = $rect.Left
            Top = $rect.Top
            Width = $width
            Height = $height
        }
        $list.Add($item)
        return $true
    }

    [void][MultiMonitorWinApi]::EnumWindows($callback, [IntPtr]::Zero)
    return $list
}

function Get-Profile {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$Name
    )

    return $Config.Profiles | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
}

function Capture-Layout {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$ProfileName
    )

    $profile = Get-Profile -Config $Config -Name $ProfileName
    if (-not $profile) { throw "Profil '$ProfileName' nicht gefunden." }

    $windows = Get-VisibleTopLevelWindows -ExcludedProcesses $Config.Settings.ExcludedProcesses |
        Sort-Object ProcessName,Title |
        Select-Object -First ([int]$Config.Settings.MaxWindowsPerCapture)

    $layout = @()
    foreach ($w in $windows) {
        $screenForWindow = Resolve-ScreenByRect -Left ([int]$w.Left) -Top ([int]$w.Top) -Width ([int]$w.Width) -Height ([int]$w.Height)
        $relativeRect = $null
        if ($screenForWindow) {
            $relativeRect = Get-RelativeRectOnScreen -ScreenItem $screenForWindow -Left ([int]$w.Left) -Top ([int]$w.Top) -Width ([int]$w.Width) -Height ([int]$w.Height)
        }

        $layout += [ordered]@{
            ProcessName = $w.ProcessName
            ExecutablePath = [string](Get-SafePropertyValue -Object $w -Name 'ExecutablePath' -DefaultValue '')
            LaunchArguments = [string](Get-SafePropertyValue -Object $w -Name 'LaunchArguments' -DefaultValue '')
            Title = $w.Title
            Left = $w.Left
            Top = $w.Top
            Width = $w.Width
            Height = $w.Height
            AssignedMonitor = if ($screenForWindow) { [string]$screenForWindow.DeviceName } else { '' }
            ZonePreset = 'Custom'
            RelativeLeft = if ($relativeRect) { [double]$relativeRect.RelativeLeft } else { -1.0 }
            RelativeTop = if ($relativeRect) { [double]$relativeRect.RelativeTop } else { -1.0 }
            RelativeWidth = if ($relativeRect) { [double]$relativeRect.RelativeWidth } else { -1.0 }
            RelativeHeight = if ($relativeRect) { [double]$relativeRect.RelativeHeight } else { -1.0 }
        }
    }

    $profile.WindowLayouts = $layout
    return $layout.Count
}

function Close-ProfileWindows {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$ProfileName
    )

    $profile = Get-Profile -Config $Config -Name $ProfileName
    if (-not $profile -or -not $profile.WindowLayouts -or $profile.WindowLayouts.Count -eq 0) {
        return
    }

    $current = Get-VisibleTopLevelWindows -ExcludedProcesses $Config.Settings.ExcludedProcesses
    $closed = 0
    $WM_CLOSE = 0x0010

    foreach ($rule in $profile.WindowLayouts) {
        $ruleExecutablePath = [string](Get-SafePropertyValue -Object $rule -Name 'ExecutablePath' -DefaultValue '')
        $ruleTitle = [string](Get-SafePropertyValue -Object $rule -Name 'Title' -DefaultValue '')

        if ([string]::IsNullOrWhiteSpace($ruleExecutablePath)) {
            continue
        }

        $matchingWindows = @($current | Where-Object {
            ([string]::Equals([string](Get-SafePropertyValue -Object $_ -Name 'ExecutablePath' -DefaultValue ''), $ruleExecutablePath, [System.StringComparison]::OrdinalIgnoreCase)) -and
            ([string]::Equals([string](Get-SafePropertyValue -Object $_ -Name 'Title' -DefaultValue ''), $ruleTitle, [System.StringComparison]::OrdinalIgnoreCase))
        })

        foreach ($w in $matchingWindows) {
            try {
                $hWnd = [IntPtr]$w.Handle
                $result = [MultiMonitorWinApi]::PostMessage($hWnd, $WM_CLOSE, [IntPtr]::Zero, [IntPtr]::Zero)
                if ($result) {
                    $closed++
                    Write-DebugLog -Config $Config -Scope 'Close' -Message "Fenster geschlossen: $($w.Title) (PID: $($w.ProcessId))"
                }
            }
            catch {
                Write-DebugLog -Config $Config -Scope 'Close' -Message "Fehler beim Schließen von $($w.Title): $_"
            }
        }
    }

    if ($closed -gt 0) {
        Start-Sleep -Milliseconds 200
        Write-Log "Profil-Fenster geschlossen: $closed Fenster."
    }
}

function Restore-Layout {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$ProfileName,
        [switch]$Detailed
    )

    $profile = Get-Profile -Config $Config -Name $ProfileName
    if (-not $profile) { throw "Profil '$ProfileName' nicht gefunden." }

    if (-not $profile.WindowLayouts -or $profile.WindowLayouts.Count -eq 0) {
        if ($Detailed) {
            return [pscustomobject]@{ Moved = 0; Missing = 0; Invalid = 0; Launched = 0 }
        }
        return 0
    }

    $current = Get-VisibleTopLevelWindows -ExcludedProcesses $Config.Settings.ExcludedProcesses
    $moved = 0
    $missing = 0
    $invalid = 0
    $launched = 0
    $usedHandles = New-Object 'System.Collections.Generic.HashSet[long]'
    $runId = ([Guid]::NewGuid().ToString('N')).Substring(0, 8)

    Write-DebugLog -Config $Config -Scope 'Restore' -Message ("Run={0} Profile='{1}' Rules={2}" -f $runId, $ProfileName, @($profile.WindowLayouts).Count)

    foreach ($rule in $profile.WindowLayouts) {
        $ruleProcessName = [string](Get-SafePropertyValue -Object $rule -Name 'ProcessName' -DefaultValue '')
        $ruleExecutablePath = [string](Get-SafePropertyValue -Object $rule -Name 'ExecutablePath' -DefaultValue '')
        $ruleLaunchArguments = [string](Get-SafePropertyValue -Object $rule -Name 'LaunchArguments' -DefaultValue '')
        $ruleTitle = [string](Get-SafePropertyValue -Object $rule -Name 'Title' -DefaultValue '')
        $ruleLeft = [int](Get-SafePropertyValue -Object $rule -Name 'Left' -DefaultValue 0)
        $ruleTop = [int](Get-SafePropertyValue -Object $rule -Name 'Top' -DefaultValue 0)
        $ruleWidth = [int](Get-SafePropertyValue -Object $rule -Name 'Width' -DefaultValue 800)
        $ruleHeight = [int](Get-SafePropertyValue -Object $rule -Name 'Height' -DefaultValue 600)
        $ruleAssignedMonitor = [string](Get-SafePropertyValue -Object $rule -Name 'AssignedMonitor' -DefaultValue '')
        $ruleZonePreset = [string](Get-SafePropertyValue -Object $rule -Name 'ZonePreset' -DefaultValue 'Custom')
        $ruleRelativeLeft = [double](Get-SafePropertyValue -Object $rule -Name 'RelativeLeft' -DefaultValue -1.0)
        $ruleRelativeTop = [double](Get-SafePropertyValue -Object $rule -Name 'RelativeTop' -DefaultValue -1.0)
        $ruleRelativeWidth = [double](Get-SafePropertyValue -Object $rule -Name 'RelativeWidth' -DefaultValue -1.0)
        $ruleRelativeHeight = [double](Get-SafePropertyValue -Object $rule -Name 'RelativeHeight' -DefaultValue -1.0)

        $zoneRect = $null
        $zoneScreen = $null
        if (
            -not [string]::IsNullOrWhiteSpace($ruleAssignedMonitor) -and
            -not [string]::IsNullOrWhiteSpace($ruleZonePreset) -and
            $ruleZonePreset -ne 'Custom'
        ) {
            $zoneScreen = Get-Screens | Where-Object { $_.DeviceName -eq $ruleAssignedMonitor } | Select-Object -First 1
            if ($zoneScreen) {
                $zoneRect = Get-ZoneRect -ScreenBounds $zoneScreen -Zone $ruleZonePreset
            }
        }

        $targetLeft = $ruleLeft
        $targetTop = $ruleTop
        $targetWidth = $ruleWidth
        $targetHeight = $ruleHeight
        if ($zoneRect) {
            $targetLeft = [int]$zoneRect.Left
            $targetTop = [int]$zoneRect.Top
            $targetWidth = [int]$zoneRect.Width
            $targetHeight = [int]$zoneRect.Height
        }
        elseif (
            -not [string]::IsNullOrWhiteSpace($ruleAssignedMonitor) -and
            $ruleRelativeLeft -ge 0 -and
            $ruleRelativeTop -ge 0 -and
            $ruleRelativeWidth -gt 0 -and
            $ruleRelativeHeight -gt 0
        ) {
            $relativeScreen = if ($zoneScreen) { $zoneScreen } else { Get-Screens | Where-Object { $_.DeviceName -eq $ruleAssignedMonitor } | Select-Object -First 1 }
            if ($relativeScreen) {
                $relativeRect = Get-AbsoluteRectFromRelative -ScreenItem $relativeScreen -RelativeLeft $ruleRelativeLeft -RelativeTop $ruleRelativeTop -RelativeWidth $ruleRelativeWidth -RelativeHeight $ruleRelativeHeight
                if ($relativeRect) {
                    $targetLeft = [int]$relativeRect.Left
                    $targetTop = [int]$relativeRect.Top
                    $targetWidth = [int]$relativeRect.Width
                    $targetHeight = [int]$relativeRect.Height
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace($ruleProcessName) -and [string]::IsNullOrWhiteSpace($ruleTitle)) {
            $invalid++
            Write-DebugLog -Config $Config -Scope 'Restore' -Message ("Run={0} Rule invalid (empty process/title)." -f $runId)
            continue
        }

        $target = $current |
            Where-Object {
                $_.ProcessName -eq $ruleProcessName -and
                $_.Title -eq $ruleTitle -and
                -not $usedHandles.Contains([int64]$_.Handle)
            } |
            Select-Object -First 1

        if (-not $target) {
            $target = $current |
                Where-Object {
                    $_.ProcessName -eq $ruleProcessName -and
                    -not $usedHandles.Contains([int64]$_.Handle)
                } |
                Select-Object -First 1
        }

        if (-not $target -and [bool](Get-SafePropertyValue -Object $Config.Settings -Name 'AutoLaunchMissingWindows' -DefaultValue $true)) {
            if (-not [string]::IsNullOrWhiteSpace($ruleExecutablePath) -and (Test-Path -LiteralPath $ruleExecutablePath)) {
                try {
                    $workingDir = Split-Path -Path $ruleExecutablePath -Parent
                    $launchArgumentsToUse = $ruleLaunchArguments
                    if (
                        ($ruleProcessName -ieq 'Code') -or
                        ([System.IO.Path]::GetFileNameWithoutExtension($ruleExecutablePath) -ieq 'Code')
                    ) {
                        if ([string]::IsNullOrWhiteSpace($launchArgumentsToUse)) {
                            $workspacePath = Resolve-VSCodeWorkspacePathFromTitle -Title $ruleTitle
                            if (-not [string]::IsNullOrWhiteSpace($workspacePath)) {
                                $launchArgumentsToUse = "--new-window `"$workspacePath`""
                            }
                        }
                        elseif ($launchArgumentsToUse -notmatch '(?i)--new-window|--reuse-window') {
                            $launchArgumentsToUse = "--new-window $launchArgumentsToUse"
                        }
                    }

                    $startParams = @{ FilePath = $ruleExecutablePath }
                    if (-not [string]::IsNullOrWhiteSpace($workingDir) -and (Test-Path -LiteralPath $workingDir)) {
                        $startParams.WorkingDirectory = $workingDir
                    }
                    if (-not [string]::IsNullOrWhiteSpace($launchArgumentsToUse)) {
                        $startParams.ArgumentList = $launchArgumentsToUse
                    }
                    Start-Process @startParams | Out-Null
                    $launched++

                    $launchDelay = [int](Get-SafePropertyValue -Object $Config.Settings -Name 'LaunchDelayMs' -DefaultValue 1800)
                    if ($launchDelay -gt 0) {
                        Start-Sleep -Milliseconds $launchDelay
                    }

                    $current = Get-VisibleTopLevelWindows -ExcludedProcesses $Config.Settings.ExcludedProcesses
                    $target = $current |
                        Where-Object {
                            $_.ProcessName -eq $ruleProcessName -and
                            $_.Title -eq $ruleTitle -and
                            -not $usedHandles.Contains([int64]$_.Handle)
                        } |
                        Select-Object -First 1

                    if (-not $target) {
                        $target = $current |
                            Where-Object {
                                $_.ProcessName -eq $ruleProcessName -and
                                -not $usedHandles.Contains([int64]$_.Handle)
                            } |
                            Select-Object -First 1
                    }
                }
                catch {
                }
            }
        }

        if (-not $target) {
            $missing++
            Write-DebugLog -Config $Config -Scope 'Restore' -Message ("Run={0} Missing window Process='{1}' Title='{2}' AssignedMonitor='{3}' Zone='{4}' Stored=({5},{6},{7},{8})" -f $runId, $ruleProcessName, $ruleTitle, $ruleAssignedMonitor, $ruleZonePreset, $ruleLeft, $ruleTop, $ruleWidth, $ruleHeight)
            continue
        }

        [void]$usedHandles.Add([int64]$target.Handle)

        $beforeRect = Get-WindowRectSnapshot -Handle ([int64]$target.Handle)
        $zoneRectText = if ($zoneRect) {
            "({0},{1},{2},{3})" -f [int]$zoneRect.Left, [int]$zoneRect.Top, [int]$zoneRect.Width, [int]$zoneRect.Height
        }
        else {
            'n/a'
        }
        $beforeText = if ($beforeRect) {
            "({0},{1},{2},{3}) dpi={4} screen={5}" -f [int]$beforeRect.Left, [int]$beforeRect.Top, [int]$beforeRect.Width, [int]$beforeRect.Height, [int]$beforeRect.Dpi, [string]$beforeRect.ScreenDeviceName
        }
        else {
            'n/a'
        }

        Write-DebugLog -Config $Config -Scope 'Restore' -Message (
            "Run={0} Move try Handle={1} Process='{2}' Title='{3}' AssignedMonitor='{4}' Zone='{5}' Stored=({6},{7},{8},{9}) Apply=({10},{11},{12},{13}) Relative=({14},{15},{16},{17}) ZoneRect={18} Before={19}" -f
            $runId, [int64]$target.Handle, $ruleProcessName, $ruleTitle, $ruleAssignedMonitor, $ruleZonePreset, $ruleLeft, $ruleTop, $ruleWidth, $ruleHeight, $targetLeft, $targetTop, $targetWidth, $targetHeight, $ruleRelativeLeft, $ruleRelativeTop, $ruleRelativeWidth, $ruleRelativeHeight, $zoneRectText, $beforeText
        )

        $ok = [MultiMonitorWinApi]::MoveWindow([IntPtr]::new([int64]$target.Handle), $targetLeft, $targetTop, $targetWidth, $targetHeight, $true)
        $lastError = [Runtime.InteropServices.Marshal]::GetLastWin32Error()
        $afterRect = Get-WindowRectSnapshot -Handle ([int64]$target.Handle)
        $afterText = if ($afterRect) {
            "({0},{1},{2},{3}) dpi={4} screen={5}" -f [int]$afterRect.Left, [int]$afterRect.Top, [int]$afterRect.Width, [int]$afterRect.Height, [int]$afterRect.Dpi, [string]$afterRect.ScreenDeviceName
        }
        else {
            'n/a'
        }

        Write-DebugLog -Config $Config -Scope 'Restore' -Message (
            "Run={0} Move result Handle={1} Ok={2} LastError={3} After={4}" -f
            $runId, [int64]$target.Handle, [bool]$ok, [int]$lastError, $afterText
        )

        if ($ok) { $moved++ }
    }

    Write-DebugLog -Config $Config -Scope 'Restore' -Message ("Run={0} Done Profile='{1}' Moved={2} Missing={3} Invalid={4} Launched={5}" -f $runId, $ProfileName, $moved, $missing, $invalid, $launched)

    if ($Detailed) {
        return [pscustomobject]@{
            Moved = $moved
            Missing = $missing
            Invalid = $invalid
            Launched = $launched
        }
    }

    return $moved
}

function Update-ProfileExecutablePathsFromCurrentWindows {
    param([Parameter(Mandatory)]$Config)

    $current = Get-VisibleTopLevelWindows -ExcludedProcesses $Config.Settings.ExcludedProcesses
    $updated = 0

    foreach ($profile in $Config.Profiles) {
        foreach ($rule in @($profile.WindowLayouts)) {
            $ruleExecutablePath = [string](Get-SafePropertyValue -Object $rule -Name 'ExecutablePath' -DefaultValue '')
            $ruleLaunchArguments = [string](Get-SafePropertyValue -Object $rule -Name 'LaunchArguments' -DefaultValue '')
            if (
                -not [string]::IsNullOrWhiteSpace($ruleExecutablePath) -and
                -not [string]::IsNullOrWhiteSpace($ruleLaunchArguments)
            ) {
                continue
            }

            $ruleProcessName = [string](Get-SafePropertyValue -Object $rule -Name 'ProcessName' -DefaultValue '')
            $ruleTitle = [string](Get-SafePropertyValue -Object $rule -Name 'Title' -DefaultValue '')
            if ([string]::IsNullOrWhiteSpace($ruleProcessName)) {
                continue
            }

            $match = $current |
                Where-Object {
                    $_.ProcessName -eq $ruleProcessName -and
                    $_.Title -eq $ruleTitle -and
                    -not [string]::IsNullOrWhiteSpace([string]$_.ExecutablePath)
                } |
                Select-Object -First 1

            if (-not $match) {
                $match = $current |
                    Where-Object {
                        $_.ProcessName -eq $ruleProcessName -and
                        -not [string]::IsNullOrWhiteSpace([string]$_.ExecutablePath)
                    } |
                    Select-Object -First 1
            }

            if ($match) {
                if ([string]::IsNullOrWhiteSpace($ruleExecutablePath)) {
                    $rule.ExecutablePath = [string](Get-SafePropertyValue -Object $match -Name 'ExecutablePath' -DefaultValue '')
                    $updated++
                }
                if ([string]::IsNullOrWhiteSpace($ruleLaunchArguments)) {
                    $rule.LaunchArguments = [string](Get-SafePropertyValue -Object $match -Name 'LaunchArguments' -DefaultValue '')
                    if (-not [string]::IsNullOrWhiteSpace([string]$rule.LaunchArguments)) {
                        $updated++
                    }
                }
            }

            if ([string]::IsNullOrWhiteSpace([string](Get-SafePropertyValue -Object $rule -Name 'LaunchArguments' -DefaultValue '')) -and ($ruleProcessName -ieq 'Code')) {
                $workspacePath = Resolve-VSCodeWorkspacePathFromTitle -Title $ruleTitle
                if (-not [string]::IsNullOrWhiteSpace($workspacePath)) {
                    $rule.LaunchArguments = "--new-window `"$workspacePath`""
                    $updated++
                }
            }
        }
    }

    return $updated
}

function Rescue-WindowsToPrimary {
    $screens = [System.Windows.Forms.Screen]::AllScreens
    $primary = $screens | Where-Object { $_.Primary } | Select-Object -First 1
    if (-not $primary) { return 0 }

    $left = ($screens | ForEach-Object { $_.Bounds.Left } | Measure-Object -Minimum).Minimum
    $top = ($screens | ForEach-Object { $_.Bounds.Top } | Measure-Object -Minimum).Minimum
    $right = ($screens | ForEach-Object { $_.Bounds.Right } | Measure-Object -Maximum).Maximum
    $bottom = ($screens | ForEach-Object { $_.Bounds.Bottom } | Measure-Object -Maximum).Maximum

    $windows = Get-VisibleTopLevelWindows -ExcludedProcesses @('dwm','explorer','ShellExperienceHost','StartMenuExperienceHost','ApplicationFrameHost')
    $moved = 0

    foreach ($w in $windows) {
        $wRight = $w.Left + $w.Width
        $wBottom = $w.Top + $w.Height

        $visibleX = ($wRight -gt $left) -and ($w.Left -lt $right)
        $visibleY = ($wBottom -gt $top) -and ($w.Top -lt $bottom)

        if ($visibleX -and $visibleY) { continue }

        $newX = [Math]::Max($primary.WorkingArea.Left + 20, $primary.WorkingArea.Left)
        $newY = [Math]::Max($primary.WorkingArea.Top + 20, $primary.WorkingArea.Top)

        $maxW = [Math]::Max(400, $primary.WorkingArea.Width - 40)
        $maxH = [Math]::Max(300, $primary.WorkingArea.Height - 60)

        $newW = [Math]::Min($w.Width, $maxW)
        $newH = [Math]::Min($w.Height, $maxH)

        if ([MultiMonitorWinApi]::MoveWindow([IntPtr]::new([int64]$w.Handle), $newX, $newY, $newW, $newH, $true)) {
            $moved++
        }
    }

    return $moved
}

function Set-DisplayMode {
    param(
        [ValidateSet('internal','extend')]
        [string]$Mode
    )

    $switch = Join-Path $env:WINDIR 'System32\DisplaySwitch.exe'
    if (-not (Test-Path -LiteralPath $switch)) {
        throw 'DisplaySwitch.exe nicht gefunden.'
    }

    $arg = if ($Mode -eq 'internal') { '/internal' } else { '/extend' }
    Start-Process -FilePath $switch -ArgumentList $arg -WindowStyle Hidden -Wait
}

if ($script:IsDotSourced) {
    Write-Host "Das Skript wurde per Dot-Sourcing geladen. Die GUI wird in diesem Modus nicht automatisch gestartet."
    Write-Host "Bitte starte es direkt mit: powershell -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    return
}

Add-Type -AssemblyName Microsoft.VisualBasic

function Get-SafePropertyValue {
    param(
        [Parameter(Mandatory)]$Object,
        [Parameter(Mandatory)][string]$Name,
        $DefaultValue = $null
    )

    if ($null -eq $Object) { return $DefaultValue }

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) {
            return $Object[$Name]
        }
        return $DefaultValue
    }

    $prop = $Object.PSObject.Properties[$Name]
    if ($null -ne $prop) {
        return $prop.Value
    }

    return $DefaultValue
}

function Ensure-ConfigSchema {
    param([Parameter(Mandatory)]$Config)

    if (-not $Config.Settings) {
        $Config | Add-Member -NotePropertyName Settings -NotePropertyValue ([ordered]@{}) -Force
    }

    if (-not $Config.Profiles) {
        $Config | Add-Member -NotePropertyName Profiles -NotePropertyValue @() -Force
    }

    if ($Config.Profiles.Count -eq 0) {
        $Config.Profiles = @(
            [ordered]@{ Name='Alltag'; DisplayMode='internal'; WindowLayouts=@() },
            [ordered]@{ Name='StreamingGaming'; DisplayMode='extend'; WindowLayouts=@() }
        )
    }

    if (-not $Config.PSObject.Properties.Name.Contains('ActiveProfile') -or [string]::IsNullOrWhiteSpace([string]$Config.ActiveProfile)) {
        $Config | Add-Member -NotePropertyName ActiveProfile -NotePropertyValue ($Config.Profiles[0].Name) -Force
    }

    foreach ($p in $Config.Profiles) {
        if (-not $p.WindowLayouts) { $p.WindowLayouts = @() }
        if (-not $p.DisplayMode) { $p.DisplayMode = 'extend' }

        $normalizedLayouts = @()
        foreach ($layout in $p.WindowLayouts) {
            $normalizedLayouts += [pscustomobject]@{
                ProcessName = [string](Get-SafePropertyValue -Object $layout -Name 'ProcessName' -DefaultValue '')
                ExecutablePath = [string](Get-SafePropertyValue -Object $layout -Name 'ExecutablePath' -DefaultValue '')
                LaunchArguments = [string](Get-SafePropertyValue -Object $layout -Name 'LaunchArguments' -DefaultValue '')
                Title = [string](Get-SafePropertyValue -Object $layout -Name 'Title' -DefaultValue '')
                Left = [int](Get-SafePropertyValue -Object $layout -Name 'Left' -DefaultValue 0)
                Top = [int](Get-SafePropertyValue -Object $layout -Name 'Top' -DefaultValue 0)
                Width = [int](Get-SafePropertyValue -Object $layout -Name 'Width' -DefaultValue 800)
                Height = [int](Get-SafePropertyValue -Object $layout -Name 'Height' -DefaultValue 600)
                AssignedMonitor = [string](Get-SafePropertyValue -Object $layout -Name 'AssignedMonitor' -DefaultValue '')
                ZonePreset = [string](Get-SafePropertyValue -Object $layout -Name 'ZonePreset' -DefaultValue 'Custom')
                RelativeLeft = [double](Get-SafePropertyValue -Object $layout -Name 'RelativeLeft' -DefaultValue -1.0)
                RelativeTop = [double](Get-SafePropertyValue -Object $layout -Name 'RelativeTop' -DefaultValue -1.0)
                RelativeWidth = [double](Get-SafePropertyValue -Object $layout -Name 'RelativeWidth' -DefaultValue -1.0)
                RelativeHeight = [double](Get-SafePropertyValue -Object $layout -Name 'RelativeHeight' -DefaultValue -1.0)
            }
        }
        $p.WindowLayouts = $normalizedLayouts
    }

    if (-not $Config.Settings.PSObject.Properties.Name.Contains('RestoreAfterSwitch')) {
        $Config.Settings | Add-Member -NotePropertyName RestoreAfterSwitch -NotePropertyValue $true -Force
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('SwitchDelayMs')) {
        $Config.Settings | Add-Member -NotePropertyName SwitchDelayMs -NotePropertyValue 2500 -Force
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('AutoLaunchMissingWindows')) {
        $Config.Settings | Add-Member -NotePropertyName AutoLaunchMissingWindows -NotePropertyValue $true -Force
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('LaunchDelayMs')) {
        $Config.Settings | Add-Member -NotePropertyName LaunchDelayMs -NotePropertyValue 1800 -Force
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('UiLanguage')) {
        $Config.Settings | Add-Member -NotePropertyName UiLanguage -NotePropertyValue 'de' -Force
    }
    elseif (@('de','en') -notcontains [string]$Config.Settings.UiLanguage) {
        $Config.Settings.UiLanguage = 'de'
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('DebugLoggingEnabled')) {
        $Config.Settings | Add-Member -NotePropertyName DebugLoggingEnabled -NotePropertyValue $false -Force
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('DebugLogPath')) {
        $Config.Settings | Add-Member -NotePropertyName DebugLogPath -NotePropertyValue "$PSScriptRoot\monitor-debug.log" -Force
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('MaxWindowsPerCapture')) {
        $Config.Settings | Add-Member -NotePropertyName MaxWindowsPerCapture -NotePropertyValue 80 -Force
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('ExcludedProcesses')) {
        $Config.Settings | Add-Member -NotePropertyName ExcludedProcesses -NotePropertyValue @('dwm','explorer','ShellExperienceHost','StartMenuExperienceHost','ApplicationFrameHost','SearchHost','TextInputHost','LockApp') -Force
    }
    if (-not $Config.Settings.PSObject.Properties.Name.Contains('RunAtStartup')) {
        $Config.Settings | Add-Member -NotePropertyName RunAtStartup -NotePropertyValue $false -Force
    }

    return $Config
}

function Get-ActiveProfileName {
    param([Parameter(Mandatory)]$Config)
    if ([string]::IsNullOrWhiteSpace([string]$Config.ActiveProfile)) {
        return $Config.Profiles[0].Name
    }
    return [string]$Config.ActiveProfile
}

function Get-ZoneRect {
    param(
        [Parameter(Mandatory)]$ScreenBounds,
        [Parameter(Mandatory)][string]$Zone
    )

    if ($ScreenBounds.PSObject.Properties.Name.Contains('WorkingAreaX')) {
        $x = [int]$ScreenBounds.WorkingAreaX
        $y = [int]$ScreenBounds.WorkingAreaY
        $w = [int]$ScreenBounds.WorkingAreaWidth
        $h = [int]$ScreenBounds.WorkingAreaHeight
    }
    else {
        $x = [int]$ScreenBounds.X
        $y = [int]$ScreenBounds.Y
        $w = [int]$ScreenBounds.Width
        $h = [int]$ScreenBounds.Height
    }
    $halfW = [int]([Math]::Floor($w / 2))
    $halfH = [int]([Math]::Floor($h / 2))

    switch ($Zone) {
        'LeftHalf'     { return @{ Left=$x; Top=$y; Width=$halfW; Height=$h } }
        'RightHalf'    { return @{ Left=($x + $halfW); Top=$y; Width=($w - $halfW); Height=$h } }
        'TopHalf'      { return @{ Left=$x; Top=$y; Width=$w; Height=$halfH } }
        'BottomHalf'   { return @{ Left=$x; Top=($y + $halfH); Width=$w; Height=($h - $halfH) } }
        'TopLeft'      { return @{ Left=$x; Top=$y; Width=$halfW; Height=$halfH } }
        'TopRight'     { return @{ Left=($x + $halfW); Top=$y; Width=($w - $halfW); Height=$halfH } }
        'BottomLeft'   { return @{ Left=$x; Top=($y + $halfH); Width=$halfW; Height=($h - $halfH) } }
        'BottomRight'  { return @{ Left=($x + $halfW); Top=($y + $halfH); Width=($w - $halfW); Height=($h - $halfH) } }
        default        { return @{ Left=$x; Top=$y; Width=$w; Height=$h } }
    }
}

function Get-CanvasZoneAtPosition {
    param([double]$CenterX, [double]$CenterY)
    
    $targetMap = $null
    foreach ($mapItem in $script:canvasScreenMap.Values) {
        if ($CenterX -ge $mapItem.CanvasX -and $CenterX -le ($mapItem.CanvasX + $mapItem.CanvasW) -and
            $CenterY -ge $mapItem.CanvasY -and $CenterY -le ($mapItem.CanvasY + $mapItem.CanvasH)) {
            $targetMap = $mapItem
            break
        }
    }
    
    if (-not $targetMap) { return $null }
    
    $zone = Get-ZonePresetFromCanvasPosition -MonitorMap $targetMap -CenterX $CenterX -CenterY $CenterY
    return [pscustomobject]@{
        Monitor = $targetMap.Screen
        Zone = $zone
        MonitorMap = $targetMap
    }
}

function Render-ZonePreview {
    if (-not $controls.ContainsKey('CanvasZonePreview') -or $null -eq $controls.CanvasZonePreview) { return }

    $canvas = $controls.CanvasZonePreview
    $canvas.Children.Clear()

    $pad = 8
    $canvasW = if ($canvas.ActualWidth -gt 50) { [double]$canvas.ActualWidth } else { [double]$canvas.Width }
    $canvasH = if ($canvas.ActualHeight -gt 50) { [double]$canvas.ActualHeight } else { [double]$canvas.Height }
    if ($canvasW -le 0 -or $canvasH -le 0) { return }

    $innerX = $pad
    $innerY = $pad
    $innerW = [Math]::Max(220, $canvasW - (2 * $pad))
    $innerH = [Math]::Max(120, $canvasH - (2 * $pad))
    $gap = 6
    $cellW = [Math]::Floor(($innerW - (2 * $gap)) / 3)
    $cellH = [Math]::Floor(($innerH - (2 * $gap)) / 3)

    $zones = @(
        @{ Name='TopLeft'; X=$innerX; Y=$innerY; W=$cellW; H=$cellH },
        @{ Name='TopHalf'; X=($innerX + $cellW + $gap); Y=$innerY; W=$cellW; H=$cellH },
        @{ Name='TopRight'; X=($innerX + (2 * ($cellW + $gap))); Y=$innerY; W=$cellW; H=$cellH },
        @{ Name='LeftHalf'; X=$innerX; Y=($innerY + $cellH + $gap); W=$cellW; H=$cellH },
        @{ Name='Fullscreen'; X=($innerX + $cellW + $gap); Y=($innerY + $cellH + $gap); W=$cellW; H=$cellH },
        @{ Name='RightHalf'; X=($innerX + (2 * ($cellW + $gap))); Y=($innerY + $cellH + $gap); W=$cellW; H=$cellH },
        @{ Name='BottomLeft'; X=$innerX; Y=($innerY + (2 * ($cellH + $gap))); W=$cellW; H=$cellH },
        @{ Name='BottomHalf'; X=($innerX + $cellW + $gap); Y=($innerY + (2 * ($cellH + $gap))); W=$cellW; H=$cellH },
        @{ Name='BottomRight'; X=($innerX + (2 * ($cellW + $gap))); Y=($innerY + (2 * ($cellH + $gap))); W=$cellW; H=$cellH }
    )

    foreach ($zone in $zones) {
        $btn = New-Object System.Windows.Controls.Button
        $btn.Content = [string]$zone.Name
        $btn.Tag = [string]$zone.Name
        $btn.FontSize = 11
        $btn.Padding = New-Object System.Windows.Thickness(2)
        $btn.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(180,14,165,233))
        $btn.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(56,189,248))
        $btn.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255,255,255))
        $btn.Opacity = if ($zone.Name -eq 'Fullscreen') { 0.95 } else { 0.88 }

        $btn.Width = [double]$zone.W
        $btn.Height = [double]$zone.H

        $null = $btn.Add_Click({
            param($sender, $e)
            $zoneName = [string]$sender.Tag
            if (-not [string]::IsNullOrWhiteSpace($zoneName)) {
                Apply-ZoneToPendingDrop -ZoneName $zoneName
            }
        })

        [System.Windows.Controls.Canvas]::SetLeft($btn, [double]$zone.X)
        [System.Windows.Controls.Canvas]::SetTop($btn, [double]$zone.Y)
        $canvas.Children.Add($btn) | Out-Null
    }
}

function Set-WindowRuleForProfile {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$ProfileName,
        [Parameter(Mandatory)]$WindowItem,
        [Parameter(Mandatory)]$ScreenItem,
        [Parameter(Mandatory)][string]$ZonePreset
    )

    $profile = Get-Profile -Config $Config -Name $ProfileName
    if (-not $profile) { throw "Profil '$ProfileName' nicht gefunden." }

    $targetRect = Get-ZoneRect -ScreenBounds $ScreenItem -Zone $ZonePreset
    $relativeRect = Get-RelativeRectOnScreen -ScreenItem $ScreenItem -Left ([int]$targetRect.Left) -Top ([int]$targetRect.Top) -Width ([int]$targetRect.Width) -Height ([int]$targetRect.Height)

    $existing = $profile.WindowLayouts |
        Where-Object {
            $_.ProcessName -eq $WindowItem.ProcessName -and
            $_.Title -eq $WindowItem.Title
        } |
        Select-Object -First 1

    if ($existing) {
        $existing.ExecutablePath = [string](Get-SafePropertyValue -Object $WindowItem -Name 'ExecutablePath' -DefaultValue '')
        $existing.LaunchArguments = [string](Get-SafePropertyValue -Object $WindowItem -Name 'LaunchArguments' -DefaultValue '')
        $existing.Left = [int]$targetRect.Left
        $existing.Top = [int]$targetRect.Top
        $existing.Width = [int]$targetRect.Width
        $existing.Height = [int]$targetRect.Height
        $existing.AssignedMonitor = [string]$ScreenItem.DeviceName
        $existing.ZonePreset = [string]$ZonePreset
        if ($relativeRect) {
            $existing.RelativeLeft = [double]$relativeRect.RelativeLeft
            $existing.RelativeTop = [double]$relativeRect.RelativeTop
            $existing.RelativeWidth = [double]$relativeRect.RelativeWidth
            $existing.RelativeHeight = [double]$relativeRect.RelativeHeight
        }
        else {
            $existing.RelativeLeft = -1.0
            $existing.RelativeTop = -1.0
            $existing.RelativeWidth = -1.0
            $existing.RelativeHeight = -1.0
        }
    }
    else {
        $profile.WindowLayouts += [ordered]@{
            ProcessName = [string]$WindowItem.ProcessName
            ExecutablePath = [string](Get-SafePropertyValue -Object $WindowItem -Name 'ExecutablePath' -DefaultValue '')
            LaunchArguments = [string](Get-SafePropertyValue -Object $WindowItem -Name 'LaunchArguments' -DefaultValue '')
            Title = [string]$WindowItem.Title
            Left = [int]$targetRect.Left
            Top = [int]$targetRect.Top
            Width = [int]$targetRect.Width
            Height = [int]$targetRect.Height
            AssignedMonitor = [string]$ScreenItem.DeviceName
            ZonePreset = [string]$ZonePreset
            RelativeLeft = if ($relativeRect) { [double]$relativeRect.RelativeLeft } else { -1.0 }
            RelativeTop = if ($relativeRect) { [double]$relativeRect.RelativeTop } else { -1.0 }
            RelativeWidth = if ($relativeRect) { [double]$relativeRect.RelativeWidth } else { -1.0 }
            RelativeHeight = if ($relativeRect) { [double]$relativeRect.RelativeHeight } else { -1.0 }
        }
    }
}

function Get-ZonePresetFromCanvasPosition {
    param(
        [Parameter(Mandatory)]$MonitorMap,
        [Parameter(Mandatory)][double]$CenterX,
        [Parameter(Mandatory)][double]$CenterY
    )

    if ($MonitorMap.CanvasW -le 0 -or $MonitorMap.CanvasH -le 0) {
        return 'Fullscreen'
    }

    $relX = ($CenterX - $MonitorMap.CanvasX) / $MonitorMap.CanvasW
    $relY = ($CenterY - $MonitorMap.CanvasY) / $MonitorMap.CanvasH

    if ([Math]::Abs($relX - 0.5) -le 0.12 -and [Math]::Abs($relY - 0.5) -le 0.12) {
        return 'Fullscreen'
    }

    if ([Math]::Abs($relX - 0.5) -le 0.14) {
        if ($relY -lt 0.5) { return 'TopHalf' }
        return 'BottomHalf'
    }

    if ($relY -ge 0.33 -and $relY -le 0.67) {
        if ($relX -lt 0.5) { return 'LeftHalf' }
        return 'RightHalf'
    }

    if ($relY -lt 0.33) {
        if ($relX -lt 0.5) { return 'TopLeft' }
        return 'TopRight'
    }

    if ($relX -lt 0.5) { return 'BottomLeft' }
    return 'BottomRight'
}

function Update-WindowRulePlacement {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$ProfileName,
        [Parameter(Mandatory)][string]$ProcessName,
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$AssignedMonitor,
        [Parameter(Mandatory)][string]$ZonePreset
    )

    $profile = Get-Profile -Config $Config -Name $ProfileName
    if (-not $profile) { return $false }

    $rule = $profile.WindowLayouts |
        Where-Object {
            $_.ProcessName -eq $ProcessName -and
            $_.Title -eq $Title
        } |
        Select-Object -First 1
    if (-not $rule) { return $false }

    $screenItem = Get-Screens | Where-Object { $_.DeviceName -eq $AssignedMonitor } | Select-Object -First 1
    if (-not $screenItem) { return $false }

    $targetRect = Get-ZoneRect -ScreenBounds $screenItem -Zone $ZonePreset
    $relativeRect = Get-RelativeRectOnScreen -ScreenItem $screenItem -Left ([int]$targetRect.Left) -Top ([int]$targetRect.Top) -Width ([int]$targetRect.Width) -Height ([int]$targetRect.Height)

    $rule.Left = [int]$targetRect.Left
    $rule.Top = [int]$targetRect.Top
    $rule.Width = [int]$targetRect.Width
    $rule.Height = [int]$targetRect.Height
    $rule.AssignedMonitor = [string]$AssignedMonitor
    $rule.ZonePreset = [string]$ZonePreset
    if ($relativeRect) {
        $rule.RelativeLeft = [double]$relativeRect.RelativeLeft
        $rule.RelativeTop = [double]$relativeRect.RelativeTop
        $rule.RelativeWidth = [double]$relativeRect.RelativeWidth
        $rule.RelativeHeight = [double]$relativeRect.RelativeHeight
    }
    else {
        $rule.RelativeLeft = -1.0
        $rule.RelativeTop = -1.0
        $rule.RelativeWidth = -1.0
        $rule.RelativeHeight = -1.0
    }
    return $true
}

function Remove-WindowRuleFromProfile {
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$ProfileName,
        [Parameter(Mandatory)][string]$ProcessName,
        [Parameter(Mandatory)][string]$Title,
        [string]$AssignedMonitor,
        [string]$ZonePreset
    )

    $profile = Get-Profile -Config $Config -Name $ProfileName
    if (-not $profile) { return $false }

    $before = @($profile.WindowLayouts).Count
    $profile.WindowLayouts = @(
        $profile.WindowLayouts | Where-Object {
            $ruleProcessName = [string](Get-SafePropertyValue -Object $_ -Name 'ProcessName' -DefaultValue '')
            $ruleTitle = [string](Get-SafePropertyValue -Object $_ -Name 'Title' -DefaultValue '')
            $ruleAssignedMonitor = [string](Get-SafePropertyValue -Object $_ -Name 'AssignedMonitor' -DefaultValue '')
            $ruleZonePreset = [string](Get-SafePropertyValue -Object $_ -Name 'ZonePreset' -DefaultValue '')

            if ($ruleProcessName -ne $ProcessName -or $ruleTitle -ne $Title) {
                return $true
            }

            if (-not [string]::IsNullOrWhiteSpace($AssignedMonitor) -and $ruleAssignedMonitor -ne $AssignedMonitor) {
                return $true
            }

            if (-not [string]::IsNullOrWhiteSpace($ZonePreset) -and $ruleZonePreset -ne $ZonePreset) {
                return $true
            }

            return $false
        }
    )

    return (@($profile.WindowLayouts).Count -lt $before)
}

[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Multi-Monitor Profile Tool" Height="760" Width="1240"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize" MinHeight="700" MinWidth="950">
    <Window.Resources>
        <SolidColorBrush x:Key="BgMain" Color="#0B1220"/>
        <SolidColorBrush x:Key="PanelBg" Color="#111827"/>
        <SolidColorBrush x:Key="CardBg" Color="#0F172A"/>
        <SolidColorBrush x:Key="Stroke" Color="#334155"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#E2E8F0"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#94A3B8"/>
        <SolidColorBrush x:Key="Accent" Color="#0EA5E9"/>
        <SolidColorBrush x:Key="AccentHover" Color="#0284C7"/>

        <Style TargetType="Window">
            <Setter Property="Background" Value="{StaticResource BgMain}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>

        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
        </Style>

        <Style TargetType="GroupBox">
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Stroke}"/>
            <Setter Property="Background" Value="{StaticResource CardBg}"/>
            <Setter Property="Padding" Value="6"/>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#0B1324"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Stroke}"/>
            <Setter Property="Padding" Value="6,4"/>
        </Style>

        <Style TargetType="ListBox">
            <Setter Property="Background" Value="#0B1324"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Stroke}"/>
        </Style>

        <Style TargetType="ListView">
            <Setter Property="Background" Value="#0B1324"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Stroke}"/>
        </Style>

        <Style TargetType="Button">
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Background" Value="{StaticResource Accent}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Accent}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" Margin="2"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource AccentHover}"/>
                                <Setter TargetName="Bd" Property="BorderBrush" Value="{StaticResource AccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="Bd" Property="Opacity" Value="0.5"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="12">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="160"/>
            <ColumnDefinition Width="10"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="{StaticResource PanelBg}" BorderBrush="{StaticResource Stroke}" BorderThickness="1" CornerRadius="12" Padding="10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <TextBlock Grid.Row="0" Text="Multi-Monitor Tool" FontSize="16" FontWeight="SemiBold" Margin="4,2,4,14" TextWrapping="Wrap"/>

                <Button x:Name="BtnNavDashboard" Grid.Row="1" Margin="4" Content="Dashboard" HorizontalContentAlignment="Left" FontSize="12"/>
                <Button x:Name="BtnNavEditor" Grid.Row="2" Margin="4" Content="Layout-Editor" HorizontalContentAlignment="Left" FontSize="12"/>
                <Button x:Name="BtnNavProfiles" Grid.Row="3" Margin="4" Content="Profile" HorizontalContentAlignment="Left" FontSize="12"/>
                <Button x:Name="BtnNavSettings" Grid.Row="4" Margin="4,4,4,0" VerticalAlignment="Top" Content="Einstellungen" HorizontalContentAlignment="Left" FontSize="12"/>
                <Button x:Name="BtnSaveSettings" Grid.Row="5" Margin="4,12,4,0" Background="#16A34A" BorderBrush="#16A34A" Content="Speichern" HorizontalContentAlignment="Left" FontSize="12"/>
                <Button x:Name="BtnHeaderReload" Grid.Row="6" Margin="4" Background="#475569" BorderBrush="#475569" Content="GUI neu starten" HorizontalContentAlignment="Left" FontSize="12"/>
                <Button x:Name="BtnNavHelp" Grid.Row="8" Margin="4" Background="#475569" BorderBrush="#475569" Content="Hilfe / Tutorial" HorizontalContentAlignment="Left" FontSize="12"/>

                <Border Grid.Row="9" Margin="4" Padding="10" CornerRadius="8" BorderBrush="#1E3A5F" BorderThickness="1" Background="#0B2744">
                    <StackPanel>
                        <TextBlock Text="Aktives Profil" Foreground="#BFDBFE" FontSize="12"/>
                        <TextBlock x:Name="TxtActiveProfileChip" FontSize="15" FontWeight="Bold"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <Grid Grid.Column="2">
            <Grid x:Name="ViewDashboard" Visibility="Visible">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" CornerRadius="12" Padding="14" Margin="0,0,0,10" BorderBrush="#2C3F5F" BorderThickness="1">
                    <Border.Background>
                        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                            <GradientStop Color="#0F172A" Offset="0"/>
                            <GradientStop Color="#0B2744" Offset="1"/>
                        </LinearGradientBrush>
                    </Border.Background>
                    <Grid>
                        <TextBlock Text="Dashboard" FontSize="22" FontWeight="SemiBold" VerticalAlignment="Center"/>
                    </Grid>
                </Border>

                <UniformGrid Grid.Row="1" Columns="3" Margin="0,0,0,10">
                    <Border Margin="0,0,8,0" Padding="10" Background="#0F172A" BorderBrush="{StaticResource Stroke}" BorderThickness="1" CornerRadius="10">
                        <StackPanel>
                            <TextBlock Text="Aktives Profil" Foreground="{StaticResource TextSecondary}"/>
                            <TextBlock x:Name="TxtCardProfile" FontSize="19" FontWeight="SemiBold"/>
                        </StackPanel>
                    </Border>
                    <Border Margin="0,0,8,0" Padding="10" Background="#0F172A" BorderBrush="{StaticResource Stroke}" BorderThickness="1" CornerRadius="10">
                        <StackPanel>
                            <TextBlock Text="Monitore aktiv" Foreground="{StaticResource TextSecondary}"/>
                            <TextBlock x:Name="TxtCardScreens" FontSize="19" FontWeight="SemiBold"/>
                        </StackPanel>
                    </Border>
                    <Border Padding="10" Background="#0F172A" BorderBrush="{StaticResource Stroke}" BorderThickness="1" CornerRadius="10">
                        <StackPanel>
                            <TextBlock Text="Letzter Scan" Foreground="{StaticResource TextSecondary}"/>
                            <TextBlock x:Name="TxtCardLastScan" FontSize="19" FontWeight="SemiBold"/>
                        </StackPanel>
                    </Border>
                </UniformGrid>

                <Grid Grid.Row="2">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="3*"/>
                        <ColumnDefinition Width="2*"/>
                    </Grid.ColumnDefinitions>

                    <GroupBox Header="Aktive Monitore" Grid.Column="0" Margin="0,0,10,0">
                        <ListView x:Name="LvScreens" Margin="8" Foreground="#E2E8F0" Background="#0B1324" BorderBrush="#1E3A5F">
                            <ListView.ItemContainerStyle>
                                <Style TargetType="ListViewItem">
                                    <Setter Property="Foreground" Value="#E2E8F0"/>
                                    <Setter Property="Background" Value="Transparent"/>
                                    <Setter Property="Padding" Value="4,3"/>
                                    <Style.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter Property="Background" Value="#1E293B"/>
                                            <Setter Property="Foreground" Value="#F8FAFC"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter Property="Background" Value="#38BDF8"/>
                                            <Setter Property="Foreground" Value="#001B2E"/>
                                            <Setter Property="FontWeight" Value="SemiBold"/>
                                        </Trigger>
                                        <MultiTrigger>
                                            <MultiTrigger.Conditions>
                                                <Condition Property="IsSelected" Value="True"/>
                                                <Condition Property="Selector.IsSelectionActive" Value="False"/>
                                            </MultiTrigger.Conditions>
                                            <Setter Property="Background" Value="#7DD3FC"/>
                                            <Setter Property="Foreground" Value="#001B2E"/>
                                        </MultiTrigger>
                                    </Style.Triggers>
                                </Style>
                            </ListView.ItemContainerStyle>
                            <ListView.View>
                                <GridView>
                                    <GridViewColumn Header="Monitor" Width="250" DisplayMemberBinding="{Binding DisplayLabel}"/>
                                    <GridViewColumn Header="Typ" Width="110" DisplayMemberBinding="{Binding RoleName}"/>
                                    <GridViewColumn Header="Primaer" Width="70" DisplayMemberBinding="{Binding Primary}"/>
                                    <GridViewColumn Header="Bounds" Width="190" DisplayMemberBinding="{Binding BoundsText}"/>
                                    <GridViewColumn Header="WorkArea" Width="190" DisplayMemberBinding="{Binding WorkText}"/>
                                </GridView>
                            </ListView.View>
                        </ListView>
                    </GroupBox>

                    <GroupBox Header="Schnellaktionen" Grid.Column="1">
                        <StackPanel Margin="8">
                            <Button x:Name="BtnRefresh" Margin="0,0,0,8" Content="Monitor-Scan"/>
                            <Button x:Name="BtnSwitchAlltag" Margin="0,0,0,8" Content="Modus Alltag"/>
                            <Button x:Name="BtnSwitchStreaming" Margin="0,0,0,8" Content="Modus Streaming/Gaming"/>
                            <Button x:Name="BtnCaptureActive" Margin="0,0,0,8" Content="Layout aktives Profil speichern"/>
                            <Button x:Name="BtnApplyActive" Margin="0,0,0,8" Content="Layout aktives Profil anwenden"/>
                            <Button x:Name="BtnRescue" Margin="0,0,0,8" Background="#16A34A" BorderBrush="#16A34A" Content="Fenster retten"/>
                            <Button x:Name="BtnExit" Margin="0,12,0,0" Background="#7F1D1D" BorderBrush="#7F1D1D" Content="Beenden"/>
                        </StackPanel>
                    </GroupBox>
                </Grid>

                <GroupBox Grid.Row="3" Header="Live-Log" Margin="0,10,0,0">
                    <TextBox x:Name="TxtLog" Margin="6" Background="#020617" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Height="130"/>
                </GroupBox>
            </Grid>

            <Grid x:Name="ViewEditor" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <GroupBox Header="Layout-Editor" Grid.Row="0" Margin="0,0,0,10">
                    <Grid Margin="8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="250"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock VerticalAlignment="Center" Text="Profil:" Margin="0,0,8,0"/>
                            <ComboBox x:Name="CmbEditorProfile" Width="240"/>
                            <Button x:Name="BtnEditorRefreshRules" Margin="8,0,0,0" Content="Regeln anzeigen"/>
                            <Button x:Name="BtnEditorSaveProfile" Margin="8,0,0,0" Background="#16A34A" BorderBrush="#16A34A" Content="Profil speichern"/>
                        </StackPanel>
                        <TextBlock Grid.Column="1" HorizontalAlignment="Right" VerticalAlignment="Center" Foreground="{StaticResource TextSecondary}" Text="Apps links ziehen und auf Monitor droppen"/>
                    </Grid>
                </GroupBox>

                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="185"/>
                        <ColumnDefinition Width="10"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <GroupBox Grid.Column="0" Header="Geoeffnete Programme">
                        <Grid Margin="8">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Button x:Name="BtnRefreshWindows" Grid.Row="0" Margin="0,0,0,8" Content="Programme aktualisieren"/>
                            <ListBox x:Name="LbOpenWindows" Grid.Row="1">
                                <ListBox.ItemContainerStyle>
                                    <Style TargetType="ListBoxItem">
                                        <Setter Property="Padding" Value="4,3"/>
                                        <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                                    </Style>
                                </ListBox.ItemContainerStyle>
                                <ListBox.ItemTemplate>
                                    <DataTemplate>
                                        <TextBlock Text="{Binding ShortText}" TextTrimming="CharacterEllipsis" FontSize="12">
                                            <TextBlock.ToolTip>
                                                <ToolTip>
                                                    <StackPanel Margin="4">
                                                        <TextBlock FontWeight="SemiBold" Text="{Binding ProcName}" Foreground="#0EA5E9"/>
                                                        <TextBlock Text="{Binding Title}" TextWrapping="Wrap" MaxWidth="380" Margin="0,2,0,4"/>
                                                        <TextBlock Text="{Binding TooltipInfo}" Foreground="#94A3B8" FontSize="11"/>
                                                    </StackPanel>
                                                </ToolTip>
                                            </TextBlock.ToolTip>
                                        </TextBlock>
                                    </DataTemplate>
                                </ListBox.ItemTemplate>
                            </ListBox>
                        </Grid>
                    </GroupBox>

                    <GroupBox Grid.Column="2" Header="Dynamische Monitorflaeche" MinHeight="380">
                        <Grid Margin="8">
                            <Canvas x:Name="CanvasLayout" Background="#0B1324" MinHeight="320"/>
                            <Border x:Name="ZonePicker" Visibility="Collapsed" Background="#111827" BorderBrush="#0EA5E9" BorderThickness="1" CornerRadius="8" Padding="8" HorizontalAlignment="Right" VerticalAlignment="Top" Margin="10">
                                <StackPanel>
                                    <TextBlock x:Name="TxtZoneTarget" Margin="0,0,0,6" FontWeight="SemiBold"/>
                                    <Border BorderBrush="#334155" BorderThickness="1" CornerRadius="8" Padding="6" Margin="0,0,0,6" Background="#0B1324">
                                        <Canvas x:Name="CanvasZonePreview" Width="380" Height="210"/>
                                    </Border>
                                    <Button x:Name="BtnZoneCancel" Background="#475569" BorderBrush="#475569" Content="Abbrechen"/>
                                </StackPanel>
                            </Border>
                        </Grid>
                    </GroupBox>
                </Grid>

                <GroupBox x:Name="GrpEditorLog" Grid.Row="2" Header="Editor-Log" Margin="0,10,0,0">
                    <TextBox x:Name="TxtEditorLog" Margin="6" Background="#020617" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Height="110"/>
                </GroupBox>
            </Grid>

            <Grid x:Name="ViewProfiles" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <GroupBox Header="Profilverwaltung" Grid.Row="0" Margin="0,0,0,10">
                    <StackPanel Orientation="Horizontal" Margin="8">
                        <Button x:Name="BtnAddProfile" Margin="0,0,8,0" Content="Neues Profil"/>
                        <Button x:Name="BtnDeleteProfile" Margin="0,0,8,0" Background="#7F1D1D" BorderBrush="#7F1D1D" Content="Profil loeschen"/>
                        <Button x:Name="BtnSetActiveProfile" Margin="0,0,8,0" Content="Als aktiv setzen"/>
                        <Button x:Name="BtnCaptureProfile" Margin="0,0,8,0" Content="Layout speichern"/>
                        <Button x:Name="BtnApplyProfile" Content="Layout anwenden"/>
                    </StackPanel>
                </GroupBox>

                <ListView x:Name="LvProfiles" Grid.Row="1" Foreground="#E2E8F0" Background="#0B1324" BorderBrush="#1E3A5F">
                    <ListView.ItemContainerStyle>
                        <Style TargetType="ListViewItem">
                            <Setter Property="Foreground" Value="#E2E8F0"/>
                            <Setter Property="Background" Value="Transparent"/>
                            <Setter Property="Padding" Value="4,3"/>
                            <Style.Triggers>
                                <Trigger Property="IsMouseOver" Value="True">
                                    <Setter Property="Background" Value="#1E293B"/>
                                    <Setter Property="Foreground" Value="#F8FAFC"/>
                                </Trigger>
                                <Trigger Property="IsSelected" Value="True">
                                    <Setter Property="Background" Value="#38BDF8"/>
                                    <Setter Property="Foreground" Value="#001B2E"/>
                                    <Setter Property="FontWeight" Value="SemiBold"/>
                                </Trigger>
                                <MultiTrigger>
                                    <MultiTrigger.Conditions>
                                        <Condition Property="IsSelected" Value="True"/>
                                        <Condition Property="Selector.IsSelectionActive" Value="False"/>
                                    </MultiTrigger.Conditions>
                                    <Setter Property="Background" Value="#7DD3FC"/>
                                    <Setter Property="Foreground" Value="#001B2E"/>
                                </MultiTrigger>
                            </Style.Triggers>
                        </Style>
                    </ListView.ItemContainerStyle>
                    <ListView.View>
                        <GridView>
                            <GridViewColumn Header="Profil" Width="220" DisplayMemberBinding="{Binding Name}"/>
                            <GridViewColumn Header="DisplayMode" Width="140" DisplayMemberBinding="{Binding DisplayMode}"/>
                            <GridViewColumn Header="Regeln" Width="90" DisplayMemberBinding="{Binding RuleCount}"/>
                            <GridViewColumn Header="Aktiv" Width="80" DisplayMemberBinding="{Binding IsActive}"/>
                        </GridView>
                    </ListView.View>
                </ListView>
            </Grid>

            <Grid x:Name="ViewSettings" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <GroupBox Header="Einstellungen" Grid.Row="0" Margin="0,0,0,10">
                    <Grid Margin="8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="2*"/>
                            <ColumnDefinition Width="2*"/>
                        </Grid.ColumnDefinitions>

                        <StackPanel Grid.Column="0" Margin="0,0,10,0" TextElement.Foreground="#E2E8F0">
                            <TextBlock x:Name="TxtConfigPath" Foreground="#E2E8F0" FontSize="13" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,10"/>
                            <CheckBox x:Name="ChkAutoRestore" Foreground="#E2E8F0" Content="Layout nach Moduswechsel automatisch wiederherstellen" Margin="0,0,0,8"/>
                            <CheckBox x:Name="ChkAutoLaunchMissing" Foreground="#E2E8F0" Content="Fehlende Programme beim Profil anwenden automatisch starten" Margin="0,0,0,8"/>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock x:Name="TxtDelayLabel" VerticalAlignment="Center" Margin="0,0,8,0" Text="Verzoegerung (ms):"/>
                                <TextBox x:Name="TxtDelay" Width="120"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                                <TextBlock x:Name="TxtLaunchDelayLabel" VerticalAlignment="Center" Margin="0,0,8,0" Text="Startwartezeit (ms):"/>
                                <TextBox x:Name="TxtLaunchDelay" Width="120"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                                <TextBlock x:Name="TxtLanguageLabel" VerticalAlignment="Center" Margin="0,0,8,0" Text="Sprache:"/>
                                <ComboBox x:Name="CmbLanguage" Width="120"/>
                            </StackPanel>
                            <CheckBox x:Name="ChkRunAtStartup" Foreground="#E2E8F0" Content="Mit Windows starten" Margin="0,8,0,0"/>
                        </StackPanel>

                        <GroupBox x:Name="GrpExcludedProcesses" Grid.Column="1" Header="Excluded Processes">
                            <Grid Margin="8">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                </Grid.RowDefinitions>
                                <TextBlock x:Name="TxtExcludedHint" Grid.Row="0" Foreground="#94A3B8" TextWrapping="Wrap" Margin="0,0,0,8" Text="Systemprozesse hier eintragen, die nie erfasst oder verschoben werden sollen."/>
                                <ListBox x:Name="LbExcluded" Grid.Row="1"/>
                                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,0">
                                    <TextBox x:Name="TxtExcludedNew" Width="170"/>
                                    <Button x:Name="BtnExcludedAdd" Margin="8,0,0,0" Content="Hinzufuegen"/>
                                    <Button x:Name="BtnExcludedRemove" Margin="8,0,0,0" Background="#7F1D1D" BorderBrush="#7F1D1D" Content="Entfernen"/>
                                </StackPanel>
                            </Grid>
                        </GroupBox>
                    </Grid>
                </GroupBox>
            </Grid>

            <Grid x:Name="ViewHelp" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <Border Grid.Row="0" CornerRadius="12" Padding="14" Margin="0,0,0,10" BorderBrush="#2C3F5F" BorderThickness="1" Background="#0F172A">
                    <StackPanel>
                        <TextBlock Text="Hilfe-Tutorial" FontSize="22" FontWeight="SemiBold"/>
                        <TextBlock Foreground="{StaticResource TextSecondary}" Text="Kurzanleitung fuer Einsteiger: in 5 Schritten zum funktionierenden Profil."/>
                    </StackPanel>
                </Border>

                <Grid Grid.Row="1">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="2*"/>
                        <ColumnDefinition Width="2*"/>
                    </Grid.ColumnDefinitions>

                    <GroupBox Header="Schritt-fuer-Schritt" Grid.Column="0" Margin="0,0,10,0">
                        <StackPanel Margin="10">
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="1. Gehe auf Dashboard und druecke Monitor-Scan, damit alle Bildschirme erkannt werden."/>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="2. Gehe auf Profile und setze dein gewuenschtes Arbeitsprofil als aktiv."/>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="3. Gehe auf Layout-Editor und aktualisiere Geoeffnete Programme."/>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="4. Ziehe ein Programm links auf den passenden Monitor rechts und waehle eine Zone (z. B. LeftHalf oder Fullscreen)."/>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="5. Speichere und teste im Dashboard mit Layout aktives Profil anwenden."/>
                        </StackPanel>
                    </GroupBox>

                    <GroupBox Header="Wichtige Hinweise" Grid.Column="1">
                        <StackPanel Margin="10">
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="Wenn Fenster verschwinden: nutze Fenster retten im Dashboard."/>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="Beamer nur bei Bedarf: nutze Modus Alltag fuer den normalen Betrieb."/>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="Wenn ein Programm nicht erscheint: oeffne es zuerst sichtbar und aktualisiere dann die Programmliste."/>
                            <TextBlock TextWrapping="Wrap" Margin="0,0,0,8" Text="Profil-Idee: Alltag fuer Arbeit, Streaming/Gaming fuer Beamer-Setups."/>
                        </StackPanel>
                    </GroupBox>
                </Grid>

                <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,10,0,0">
                    <Button x:Name="BtnHelpToDashboard" Margin="4" Content="Zum Dashboard"/>
                    <Button x:Name="BtnHelpToEditor" Margin="4" Content="Zum Layout-Editor"/>
                </StackPanel>
            </Grid>
        </Grid>
    </Grid>
</Window>
"@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$controls = @{}
@(
    'BtnNavDashboard','BtnNavEditor','BtnNavProfiles','BtnNavSettings','BtnNavHelp','TxtActiveProfileChip',
    'ViewDashboard','ViewEditor','ViewProfiles','ViewSettings','ViewHelp',
    'TxtCardProfile','TxtCardScreens','TxtCardLastScan','LvScreens',
    'BtnHeaderReload',
    'BtnRefresh','BtnSwitchAlltag','BtnSwitchStreaming','BtnCaptureActive','BtnApplyActive','BtnRescue','BtnExit',
    'TxtLog','GrpEditorLog','TxtEditorLog',
    'CmbEditorProfile','BtnEditorRefreshRules','BtnEditorSaveProfile','BtnRefreshWindows','LbOpenWindows','CanvasLayout','ZonePicker','TxtZoneTarget','CanvasZonePreview',
    'BtnZoneCancel',
    'BtnAddProfile','BtnDeleteProfile','BtnSetActiveProfile','BtnCaptureProfile','BtnApplyProfile','LvProfiles',
    'TxtConfigPath','ChkAutoRestore','ChkAutoLaunchMissing','TxtDelayLabel','TxtDelay','TxtLaunchDelayLabel','TxtLaunchDelay','TxtLanguageLabel','CmbLanguage','BtnSaveSettings','ChkRunAtStartup','GrpExcludedProcesses','TxtExcludedHint','LbExcluded','TxtExcludedNew','BtnExcludedAdd','BtnExcludedRemove'
    ,'BtnHelpToDashboard','BtnHelpToEditor'
) | ForEach-Object {
    $controls[$_] = $window.FindName($_)
}

$script:config = Ensure-ConfigSchema -Config (Load-Config -Path $ConfigPath)
$pathMigrations = Update-ProfileExecutablePathsFromCurrentWindows -Config $script:config
Save-Config -Config $script:config -Path $ConfigPath

$script:lastScan = 'Noch kein Scan'
$script:openWindows = @()
$script:openWindowsByKey = @{}
$script:canvasScreenMap = @{}
$script:pendingDrop = $null
$script:dragChip = $null
$script:dragChipStartPoint = $null
$script:dragChipStartLeft = 0
$script:dragChipStartTop = 0

function Write-Log {
    param([string]$Message)
    $stamp = Get-Date -Format 'HH:mm:ss'
    $controls.TxtLog.AppendText("[$stamp] $Message`r`n")
    $controls.TxtLog.ScrollToEnd()
}

function Write-EditorLog {
    param([string]$Message)
    if (-not $controls.ContainsKey('TxtEditorLog') -or $null -eq $controls.TxtEditorLog) { return }
    $stamp = Get-Date -Format 'HH:mm:ss'
    $controls.TxtEditorLog.AppendText("[$stamp] $Message`r`n")
    $controls.TxtEditorLog.ScrollToEnd()
}

$script:i18n = @{
    de = @{
        window_title = 'Multi-Monitor Profile Tool'
        settings_path_prefix = 'Konfiguration:'
        nav_dashboard = 'Dashboard'
        nav_editor = 'Layout-Editor'
        nav_profiles = 'Profile'
        nav_settings = 'Einstellungen'
        nav_help = 'Hilfe / Tutorial'
        btn_header_reload = 'GUI neu starten'
        btn_refresh = 'Monitor-Scan'
        btn_switch_alltag = 'Modus Alltag'
        btn_switch_streaming = 'Modus Streaming/Gaming'
        btn_capture_active = 'Layout aktives Profil speichern'
        btn_apply_active = 'Layout aktives Profil anwenden'
        btn_rescue = 'Fenster retten'
        btn_exit = 'Beenden'
        btn_refresh_windows = 'Programme aktualisieren'
        btn_zone_cancel = 'Abbrechen'
        btn_editor_refresh = 'Regeln anzeigen'
        btn_editor_save = 'Profil speichern'
        btn_add_profile = 'Neues Profil'
        btn_delete_profile = 'Profil loeschen'
        btn_set_active_profile = 'Als aktiv setzen'
        btn_capture_profile = 'Layout speichern'
        btn_apply_profile = 'Layout anwenden'
        chk_auto_restore = 'Layout nach Moduswechsel automatisch wiederherstellen'
        chk_auto_launch = 'Fehlende Programme beim Profil anwenden automatisch starten'
        lbl_delay = 'Verzoegerung (ms):'
        lbl_launch_delay = 'Startwartezeit (ms):'
        lbl_language = 'Sprache:'
        btn_save_settings = 'Speichern'
        grp_excluded_processes = 'Ausgeschlossene Prozesse'
        txt_excluded_hint = 'Systemprozesse hier eintragen, die nie erfasst oder verschoben werden sollen. Normale Apps wie Code oder Browser nicht hinzufuegen.'
        btn_excluded_add = 'Hinzufuegen'
        btn_excluded_remove = 'Entfernen'
        chk_run_at_startup = 'Mit Windows starten'
        btn_help_to_dashboard = 'Zum Dashboard'
        btn_help_to_editor = 'Zum Layout-Editor'
        grp_editor_log = 'Editor-Log'
        editor_log_started = 'Editor-Log bereit. Aktionen im Layout-Editor werden hier angezeigt.'
        log_tool_started = 'Tool gestartet. Nutze den Layout-Editor fuer Drag-and-Drop-Zonen je Profil.'
    }
    en = @{
        window_title = 'Multi-Monitor Profile Tool'
        settings_path_prefix = 'Configuration:'
        nav_dashboard = 'Dashboard'
        nav_editor = 'Layout Editor'
        nav_profiles = 'Profiles'
        nav_settings = 'Settings'
        nav_help = 'Help / Tutorial'
        btn_header_reload = 'Restart GUI'
        btn_refresh = 'Scan monitors'
        btn_switch_alltag = 'Work mode'
        btn_switch_streaming = 'Streaming/Gaming mode'
        btn_capture_active = 'Save active profile layout'
        btn_apply_active = 'Apply active profile layout'
        btn_rescue = 'Rescue windows'
        btn_exit = 'Exit'
        btn_refresh_windows = 'Refresh programs'
        btn_zone_cancel = 'Cancel'
        btn_editor_refresh = 'Show rules'
        btn_editor_save = 'Save profile'
        btn_add_profile = 'New profile'
        btn_delete_profile = 'Delete profile'
        btn_set_active_profile = 'Set active'
        btn_capture_profile = 'Save layout'
        btn_apply_profile = 'Apply layout'
        chk_auto_restore = 'Automatically restore layout after mode switch'
        chk_auto_launch = 'Automatically launch missing programs when applying profile'
        lbl_delay = 'Delay (ms):'
        lbl_launch_delay = 'Launch wait (ms):'
        lbl_language = 'Language:'
        btn_save_settings = 'Save'
        grp_excluded_processes = 'Excluded Processes'
        txt_excluded_hint = 'Add system processes that should never be captured or moved. Do not add normal apps like Code or browsers.'
        btn_excluded_add = 'Add'
        btn_excluded_remove = 'Remove'
        chk_run_at_startup = 'Run at Windows startup'
        btn_help_to_dashboard = 'Go to dashboard'
        btn_help_to_editor = 'Go to layout editor'
        grp_editor_log = 'Editor Log'
        editor_log_started = 'Editor log ready. Layout editor actions are shown here.'
        log_tool_started = 'Tool started. Use the layout editor for drag-and-drop zones per profile.'
    }
}

function Get-UiLanguage {
    $value = [string](Get-SafePropertyValue -Object $config.Settings -Name 'UiLanguage' -DefaultValue 'de')
    if (@('de','en') -contains $value) {
        return $value
    }
    return 'de'
}

function T {
    param([Parameter(Mandatory)][string]$Key)

    $lang = Get-UiLanguage
    if ($script:i18n.ContainsKey($lang) -and $script:i18n[$lang].ContainsKey($Key)) {
        return [string]$script:i18n[$lang][$Key]
    }
    if ($script:i18n['de'].ContainsKey($Key)) {
        return [string]$script:i18n['de'][$Key]
    }
    return $Key
}

function Apply-UiLanguage {
    $window.Title = T 'window_title'

    $controls.BtnNavDashboard.Content = T 'nav_dashboard'
    $controls.BtnNavEditor.Content = T 'nav_editor'
    $controls.BtnNavProfiles.Content = T 'nav_profiles'
    $controls.BtnNavSettings.Content = T 'nav_settings'
    $controls.BtnNavHelp.Content = T 'nav_help'
    $controls.BtnHeaderReload.Content = T 'btn_header_reload'
    $controls.BtnRefresh.Content = T 'btn_refresh'
    $controls.BtnSwitchAlltag.Content = T 'btn_switch_alltag'
    $controls.BtnSwitchStreaming.Content = T 'btn_switch_streaming'
    $controls.BtnCaptureActive.Content = T 'btn_capture_active'
    $controls.BtnApplyActive.Content = T 'btn_apply_active'
    $controls.BtnRescue.Content = T 'btn_rescue'
    $controls.BtnExit.Content = T 'btn_exit'
    $controls.BtnRefreshWindows.Content = T 'btn_refresh_windows'
    $controls.BtnZoneCancel.Content = T 'btn_zone_cancel'
    $controls.BtnEditorRefreshRules.Content = T 'btn_editor_refresh'
    $controls.BtnEditorSaveProfile.Content = T 'btn_editor_save'
    $controls.BtnAddProfile.Content = T 'btn_add_profile'
    $controls.BtnDeleteProfile.Content = T 'btn_delete_profile'
    $controls.BtnSetActiveProfile.Content = T 'btn_set_active_profile'
    $controls.BtnCaptureProfile.Content = T 'btn_capture_profile'
    $controls.BtnApplyProfile.Content = T 'btn_apply_profile'
    $controls.ChkAutoRestore.Content = T 'chk_auto_restore'
    $controls.ChkAutoLaunchMissing.Content = T 'chk_auto_launch'
    $controls.TxtDelayLabel.Text = T 'lbl_delay'
    $controls.TxtLaunchDelayLabel.Text = T 'lbl_launch_delay'
    $controls.TxtLanguageLabel.Text = T 'lbl_language'
    $controls.BtnSaveSettings.Content = T 'btn_save_settings'
    $controls.GrpExcludedProcesses.Header = T 'grp_excluded_processes'
    $controls.TxtExcludedHint.Text = T 'txt_excluded_hint'
    $controls.BtnExcludedAdd.Content = T 'btn_excluded_add'
    $controls.BtnExcludedRemove.Content = T 'btn_excluded_remove'
    $controls.ChkRunAtStartup.Content = T 'chk_run_at_startup'

    if ($controls.ContainsKey('BtnHelpToDashboard') -and $null -ne $controls['BtnHelpToDashboard']) {
        $controls['BtnHelpToDashboard'].Content = T 'btn_help_to_dashboard'
    }
    if ($controls.ContainsKey('BtnHelpToEditor') -and $null -ne $controls['BtnHelpToEditor']) {
        $controls['BtnHelpToEditor'].Content = T 'btn_help_to_editor'
    }

    if ($controls.ContainsKey('GrpEditorLog') -and $null -ne $controls['GrpEditorLog']) {
        $controls.GrpEditorLog.Header = T 'grp_editor_log'
    }

    $controls.TxtConfigPath.Text = "$(T 'settings_path_prefix')`n$ConfigPath"
}

function Show-View {
    param([ValidateSet('Dashboard','Editor','Profiles','Settings','Help')][string]$Name)

    $controls.ViewDashboard.Visibility = 'Collapsed'
    $controls.ViewEditor.Visibility = 'Collapsed'
    $controls.ViewProfiles.Visibility = 'Collapsed'
    $controls.ViewSettings.Visibility = 'Collapsed'
    $controls.ViewHelp.Visibility = 'Collapsed'

    switch ($Name) {
        'Dashboard' { $controls.ViewDashboard.Visibility = 'Visible' }
        'Editor'    { $controls.ViewEditor.Visibility = 'Visible' }
        'Profiles'  { $controls.ViewProfiles.Visibility = 'Visible' }
        'Settings'  { $controls.ViewSettings.Visibility = 'Visible' }
        'Help'      { $controls.ViewHelp.Visibility = 'Visible' }
    }
}

function Update-ScreenList {
    $data = foreach ($s in Get-Screens) {
        [pscustomobject]@{
            DisplayLabel = $s.DisplayLabel
            RoleName = $s.RoleName
            Primary = if ($s.Primary) { 'Ja' } else { 'Nein' }
            BoundsText = "X=$($s.BoundsX) Y=$($s.BoundsY) W=$($s.BoundsWidth) H=$($s.BoundsHeight)"
            WorkText = "X=$($s.WorkingAreaX) Y=$($s.WorkingAreaY) W=$($s.WorkingAreaWidth) H=$($s.WorkingAreaHeight)"
        }
    }
    $controls.LvScreens.ItemsSource = $data
    $script:lastScan = Get-Date -Format 'dd.MM.yyyy HH:mm:ss'
}

function Update-ProfileSelectors {
    $names = @($config.Profiles | ForEach-Object { $_.Name })
    $controls.CmbEditorProfile.ItemsSource = $names

    $active = Get-ActiveProfileName -Config $config
    if ($names -contains $active) {
        $controls.CmbEditorProfile.SelectedItem = $active
    }
    elseif ($names.Count -gt 0) {
        $controls.CmbEditorProfile.SelectedIndex = 0
    }
}

function Update-ProfileList {
    $active = Get-ActiveProfileName -Config $config
    $items = foreach ($p in $config.Profiles) {
        [pscustomobject]@{
            Name = $p.Name
            DisplayMode = $p.DisplayMode
            RuleCount = @($p.WindowLayouts).Count
            IsActive = if ($p.Name -eq $active) { 'Ja' } else { 'Nein' }
        }
    }
    $controls.LvProfiles.ItemsSource = $items
}

function Refresh-DashboardCards {
    $active = Get-ActiveProfileName -Config $config
    $controls.TxtActiveProfileChip.Text = $active
    $controls.TxtCardProfile.Text = $active
    $controls.TxtCardScreens.Text = [string]([System.Windows.Forms.Screen]::AllScreens.Count)
    $controls.TxtCardLastScan.Text = $script:lastScan
}

function Update-SettingsView {
    $controls.ChkAutoRestore.IsChecked = [bool]$config.Settings.RestoreAfterSwitch
    $controls.ChkAutoLaunchMissing.IsChecked = [bool](Get-SafePropertyValue -Object $config.Settings -Name 'AutoLaunchMissingWindows' -DefaultValue $true)
    $controls.TxtDelay.Text = [string]$config.Settings.SwitchDelayMs
    $controls.TxtLaunchDelay.Text = [string](Get-SafePropertyValue -Object $config.Settings -Name 'LaunchDelayMs' -DefaultValue 1800)
    $controls.CmbLanguage.ItemsSource = @(
        [pscustomobject]@{ Name='Deutsch'; Value='de' },
        [pscustomobject]@{ Name='English'; Value='en' }
    )
    $controls.CmbLanguage.DisplayMemberPath = 'Name'
    $controls.CmbLanguage.SelectedValuePath = 'Value'
    $controls.CmbLanguage.SelectedValue = [string](Get-SafePropertyValue -Object $config.Settings -Name 'UiLanguage' -DefaultValue 'de')
    $controls.LbExcluded.ItemsSource = @($config.Settings.ExcludedProcesses)
    # Startup-Checkbox: aktuellen Registry-Zustand lesen (launcher.bat)
    $regKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
    $regName = 'MultiMonitorProfileTool'
    $regEntry = Get-ItemProperty -Path $regKey -Name $regName -ErrorAction SilentlyContinue
    $controls.ChkRunAtStartup.IsChecked = ($null -ne $regEntry)
    Apply-UiLanguage
}

function Reload-GuiState {
    try {
        Write-Log 'GUI-Reload gestartet...'
        $script:monitorMetadataCache = $null
        $script:config = Ensure-ConfigSchema -Config (Load-Config -Path $ConfigPath)
        Save-Config -Config $script:config -Path $ConfigPath

        Update-ScreenList
        Update-ProfileSelectors
        Update-ProfileList
        Update-SettingsView
        Refresh-DashboardCards
        Refresh-OpenWindows
        Render-LayoutCanvas

        Write-Log "GUI neu geladen: Konfiguration, Monitore, Profile und Layout aktualisiert (Monitore: $([System.Windows.Forms.Screen]::AllScreens.Count))."
        return $script:config
    }
    catch {
        Write-Log "GUI-Reload fehlgeschlagen: $($_.Exception.Message)"
        return $null
    }
}

function Restart-ToolProcess {
    if ([string]::IsNullOrWhiteSpace($PSCommandPath) -or -not (Test-Path -LiteralPath $PSCommandPath)) {
        throw 'Skriptpfad fuer Neustart nicht verfuegbar.'
    }

    $argsList = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-STA',
        '-File', "`"$PSCommandPath`"",
        '-ConfigPath', "`"$ConfigPath`""
    )

    Start-Process -FilePath 'powershell.exe' -ArgumentList $argsList | Out-Null
    Write-Log 'GUI wird neu gestartet, um Code-Aenderungen zu uebernehmen...'
    $window.Close()
}

function Refresh-OpenWindows {
    $script:openWindows = Get-VisibleTopLevelWindows -ExcludedProcesses $config.Settings.ExcludedProcesses |
        Sort-Object ProcessName,Title

    $script:openWindowsByKey = @{}
    $list = foreach ($w in $script:openWindows) {
        $key = "$($w.ProcessName)|$($w.Title)|$($w.Handle)"
        $script:openWindowsByKey[$key] = $w
        [pscustomobject]@{
            Key         = $key
            DisplayText = "[$($w.ProcessName)] $($w.Title)"
            ShortText   = "[$($w.ProcessName)]"
            ProcName    = $w.ProcessName
            Title       = $w.Title
            ExecutablePath = [string](Get-SafePropertyValue -Object $w -Name 'ExecutablePath' -DefaultValue '')
            LaunchArguments = [string](Get-SafePropertyValue -Object $w -Name 'LaunchArguments' -DefaultValue '')
            TooltipInfo = "PID: $($w.ProcessId)   |   Pos: $($w.Left), $($w.Top)   |   Größe: $($w.Width) × $($w.Height)"
        }
    }
    $controls.LbOpenWindows.ItemsSource = $list
}

function Resolve-SelectedEditorProfile {
    $name = [string]$controls.CmbEditorProfile.SelectedItem
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = Get-ActiveProfileName -Config $config
    }
    return $name
}

function Get-CanvasZoneRect {
    param(
        [Parameter(Mandatory)]$MonitorMap,
        [Parameter(Mandatory)][string]$Zone,
        [int]$StackIndex = 0
    )

    $headerOffset = 34
    $padding = 10

    $innerX = $MonitorMap.CanvasX + $padding
    $innerY = $MonitorMap.CanvasY + $headerOffset
    $innerW = [Math]::Max(80, $MonitorMap.CanvasW - (2 * $padding))
    $innerH = [Math]::Max(60, $MonitorMap.CanvasH - $headerOffset - $padding)

    $halfW = [int]([Math]::Floor($innerW / 2))
    $halfH = [int]([Math]::Floor($innerH / 2))

    $rect = switch ($Zone) {
        'TopLeft'     { @{ X=$innerX; Y=$innerY; W=$halfW; H=$halfH } }
        'TopRight'    { @{ X=($innerX + $halfW); Y=$innerY; W=($innerW - $halfW); H=$halfH } }
        'BottomLeft'  { @{ X=$innerX; Y=($innerY + $halfH); W=$halfW; H=($innerH - $halfH) } }
        'BottomRight' { @{ X=($innerX + $halfW); Y=($innerY + $halfH); W=($innerW - $halfW); H=($innerH - $halfH) } }
        'LeftHalf'    { @{ X=$innerX; Y=$innerY; W=$halfW; H=$innerH } }
        'RightHalf'   { @{ X=($innerX + $halfW); Y=$innerY; W=($innerW - $halfW); H=$innerH } }
        'TopHalf'     { @{ X=$innerX; Y=$innerY; W=$innerW; H=$halfH } }
        'BottomHalf'  { @{ X=$innerX; Y=($innerY + $halfH); W=$innerW; H=($innerH - $halfH) } }
        default       { @{ X=$innerX; Y=$innerY; W=$innerW; H=$innerH } }
    }

    $inset = [Math]::Min([Math]::Max($StackIndex, 0) * 6, 24)
    $finalW = [Math]::Max(52, $rect.W - (2 * $inset))
    $finalH = [Math]::Max(36, $rect.H - (2 * $inset))

    return [pscustomobject]@{
        X = [int]($rect.X + $inset)
        Y = [int]($rect.Y + $inset)
        W = [int]$finalW
        H = [int]$finalH
    }
}

function Render-LayoutCanvas {
    $controls.CanvasLayout.Children.Clear()
    $script:canvasScreenMap = @{}

    $screens = @(Get-Screens)
    if ($screens.Count -eq 0) { return }

    $minX = ($screens | Measure-Object -Property BoundsX -Minimum).Minimum
    $minY = ($screens | Measure-Object -Property BoundsY -Minimum).Minimum
    $maxRight = ($screens | ForEach-Object { $_.BoundsX + $_.BoundsWidth } | Measure-Object -Maximum).Maximum
    $maxBottom = ($screens | ForEach-Object { $_.BoundsY + $_.BoundsHeight } | Measure-Object -Maximum).Maximum
    $virtualW = [Math]::Max(1, $maxRight - $minX)
    $virtualH = [Math]::Max(1, $maxBottom - $minY)

    # Nutze die tatsaechlich verfuegbare Flaeche, sonst werden Monitore abgeschnitten,
    # wenn das Fenster kleiner als die alten Mindestwerte ist.
    $canvasW = [int]$controls.CanvasLayout.ActualWidth
    if ($canvasW -le 0) { $canvasW = 760 }
    $canvasH = [int]$controls.CanvasLayout.ActualHeight
    if ($canvasH -le 0) { $canvasH = 470 }

    $scale = [Math]::Min(($canvasW - 20) / $virtualW, ($canvasH - 20) / $virtualH)
    $offsetX = 10
    $offsetY = 10

    foreach ($screen in $screens) {
        $x = [int](($screen.BoundsX - $minX) * $scale + $offsetX)
        $y = [int](($screen.BoundsY - $minY) * $scale + $offsetY)
        $w = [int]([Math]::Max(120, $screen.BoundsWidth * $scale))
        $h = [int]([Math]::Max(80, $screen.BoundsHeight * $scale))

        $border = New-Object System.Windows.Controls.Border
        $border.Width = $w
        $border.Height = $h
        $border.CornerRadius = New-Object System.Windows.CornerRadius(8)
        $border.BorderThickness = New-Object System.Windows.Thickness(2)
        $border.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(14,165,233))
        $border.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(50,30,64,175))
        $border.Tag = $screen.DeviceName
        $border.AllowDrop = $true

        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = "$($screen.DisplayLabel) - $($screen.BoundsWidth)x$($screen.BoundsHeight)"
        $label.Margin = New-Object System.Windows.Thickness(6)
        $label.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(226,232,240))
        $border.Child = $label

        $null = $border.Add_DragOver({ param($sender, $e) $e.Effects = 'Copy'; $e.Handled = $true })
        $null = $border.Add_Drop({
            param($sender, $e)
            try {
                $key = [string]$e.Data.GetData([System.String])
                if ([string]::IsNullOrWhiteSpace($key)) { return }
                if (-not $script:openWindowsByKey.ContainsKey($key)) { return }

                $win = $script:openWindowsByKey[$key]
                $device = [string]$sender.Tag
                $screenItem = (Get-Screens | Where-Object { $_.DeviceName -eq $device } | Select-Object -First 1)
                if (-not $screenItem) { return }

                $script:pendingDrop = [pscustomobject]@{
                    WindowItem = $win
                    ScreenItem = $screenItem
                }

                $targetName = if ([string]::IsNullOrWhiteSpace([string]$screenItem.DisplayLabel)) { $device } else { [string]$screenItem.DisplayLabel }
                $controls.TxtZoneTarget.Text = "Ziel: [$($win.ProcessName)] $($win.Title) auf $targetName"
                Render-ZonePreview
                $controls.ZonePicker.Visibility = 'Visible'
            }
            catch {
                Write-Log "Drop fehlgeschlagen: $($_.Exception.Message)"
            }
        })

        [System.Windows.Controls.Canvas]::SetLeft($border, $x)
        [System.Windows.Controls.Canvas]::SetTop($border, $y)
        $controls.CanvasLayout.Children.Add($border) | Out-Null

        $script:canvasScreenMap[$screen.DeviceName] = [pscustomobject]@{
            CanvasX = $x
            CanvasY = $y
            CanvasW = $w
            CanvasH = $h
            Screen = $screen
        }
    }

    $profileName = Resolve-SelectedEditorProfile
    $profile = Get-Profile -Config $config -Name $profileName
    if (-not $profile) { return }

    $zoneStacks = @{}

    foreach ($rule in $profile.WindowLayouts) {
        $ruleAssignedMonitor = [string](Get-SafePropertyValue -Object $rule -Name 'AssignedMonitor' -DefaultValue '')
        $ruleZonePreset = [string](Get-SafePropertyValue -Object $rule -Name 'ZonePreset' -DefaultValue 'Custom')

        if ([string]::IsNullOrWhiteSpace($ruleAssignedMonitor)) { continue }
        if (-not $script:canvasScreenMap.ContainsKey($ruleAssignedMonitor)) { continue }

        $map = $script:canvasScreenMap[$ruleAssignedMonitor]
        $zone = if ([string]::IsNullOrWhiteSpace($ruleZonePreset)) { 'Custom' } else { $ruleZonePreset }
        $zoneKey = "{0}|{1}" -f $ruleAssignedMonitor, $zone
        if (-not $zoneStacks.ContainsKey($zoneKey)) {
            $zoneStacks[$zoneKey] = 0
        }
        $zoneIndex = [int]$zoneStacks[$zoneKey]
        $zoneStacks[$zoneKey] = $zoneIndex + 1

        $zoneRect = Get-CanvasZoneRect -MonitorMap $map -Zone $zone -StackIndex $zoneIndex

        $zoneBlock = New-Object System.Windows.Controls.Border
        $zoneBlock.CornerRadius = New-Object System.Windows.CornerRadius(10)
        $zoneBlock.BorderThickness = New-Object System.Windows.Thickness(2)
        $zoneBlock.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(56,189,248))
        $zoneBlock.Background = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(105,2,132,199))
        $zoneBlock.Padding = New-Object System.Windows.Thickness(6,4,6,4)
        $zoneBlock.Cursor = [System.Windows.Input.Cursors]::SizeAll
        $zoneBlock.Width = $zoneRect.W
        $zoneBlock.Height = $zoneRect.H
        $zoneBlock.Tag = [pscustomobject]@{
            ProcessName = [string](Get-SafePropertyValue -Object $rule -Name 'ProcessName' -DefaultValue '')
            Title = [string](Get-SafePropertyValue -Object $rule -Name 'Title' -DefaultValue '')
            AssignedMonitor = $ruleAssignedMonitor
            ZonePreset = [string]$zone
        }

        $zonePanel = New-Object System.Windows.Controls.DockPanel

        $zoneLabel = New-Object System.Windows.Controls.TextBlock
        $zoneLabel.Text = $zone
        $zoneLabel.FontSize = 10
        $zoneLabel.FontWeight = [System.Windows.FontWeights]::SemiBold
        $zoneLabel.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(220,186,230,253))
        $zoneLabel.HorizontalAlignment = 'Right'
        [System.Windows.Controls.DockPanel]::SetDock($zoneLabel, [System.Windows.Controls.Dock]::Top)

        $appLabel = New-Object System.Windows.Controls.TextBlock
        $appLabel.Text = "$($rule.ProcessName)"
        $appLabel.FontSize = 13
        $appLabel.FontWeight = [System.Windows.FontWeights]::Bold
        $appLabel.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(255,255,255))
        $appLabel.TextTrimming = 'CharacterEllipsis'

        $titleLabel = New-Object System.Windows.Controls.TextBlock
        $titleLabel.Text = [string]$rule.Title
        $titleLabel.FontSize = 10
        $titleLabel.Foreground = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromArgb(220,226,232,240))
        $titleLabel.TextTrimming = 'CharacterEllipsis'

        $stack = New-Object System.Windows.Controls.StackPanel
        $stack.Orientation = 'Vertical'
        $stack.Children.Add($zoneLabel) | Out-Null
        $stack.Children.Add($appLabel) | Out-Null
        $stack.Children.Add($titleLabel) | Out-Null
        $zonePanel.Children.Add($stack) | Out-Null
        $zoneBlock.Child = $zonePanel

        [System.Windows.Controls.Canvas]::SetLeft($zoneBlock, $zoneRect.X)
        [System.Windows.Controls.Canvas]::SetTop($zoneBlock, $zoneRect.Y)

        $null = $zoneBlock.Add_MouseLeftButtonDown({
            param($sender, $e)
            $script:dragChip = $sender
            $script:dragChipStartPoint = $e.GetPosition($controls.CanvasLayout)
            $script:dragChipStartLeft = [double][System.Windows.Controls.Canvas]::GetLeft($sender)
            $script:dragChipStartTop = [double][System.Windows.Controls.Canvas]::GetTop($sender)
            $sender.Opacity = 0.85
            $sender.CaptureMouse() | Out-Null
            $e.Handled = $true
        })

        $null = $zoneBlock.Add_MouseMove({
            param($sender, $e)
            if ($null -eq $script:dragChip -or $script:dragChip -ne $sender) { return }
            if (-not $sender.IsMouseCaptured) { return }

            $pos = $e.GetPosition($controls.CanvasLayout)
            $dx = $pos.X - $script:dragChipStartPoint.X
            $dy = $pos.Y - $script:dragChipStartPoint.Y

            [System.Windows.Controls.Canvas]::SetLeft($sender, $script:dragChipStartLeft + $dx)
            [System.Windows.Controls.Canvas]::SetTop($sender, $script:dragChipStartTop + $dy)
            
            # Live-Preview: Berechne Zielzone während des Drag-Vorgangs
            $width = if ($sender.ActualWidth -gt 0) { [double]$sender.ActualWidth } else { 170.0 }
            $height = if ($sender.ActualHeight -gt 0) { [double]$sender.ActualHeight } else { 90.0 }
            $centerX = ($script:dragChipStartLeft + $dx) + ($width / 2)
            $centerY = ($script:dragChipStartTop + $dy) + ($height / 2)
            
            $zoneInfo = Get-CanvasZoneAtPosition -CenterX $centerX -CenterY $centerY
            
            # Highlight alle Monitor-Borders, dann nur die Ziel-Monitor heller machen
            foreach ($border in $controls.CanvasLayout.Children | Where-Object { $_ -is [System.Windows.Controls.Border] -and $_.Tag -is [string] }) {
                $border.BorderThickness = New-Object System.Windows.Thickness(2)
                $border.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(14,165,233))
                $border.Opacity = 0.7
            }
            
            if ($zoneInfo) {
                $targetDeviceName = $zoneInfo.Monitor.DeviceName
                $targetMonitorBorder = $controls.CanvasLayout.Children | Where-Object { $_.Tag -eq $targetDeviceName } | Select-Object -First 1
                if ($targetMonitorBorder) {
                    $targetMonitorBorder.BorderThickness = New-Object System.Windows.Thickness(3)
                    $targetMonitorBorder.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(34,197,94))
                    $targetMonitorBorder.Opacity = 1.0
                }
                
                # Zeige Zielinfo an
                $displayLabel = if ([string]::IsNullOrWhiteSpace($zoneInfo.Monitor.DisplayLabel)) { $targetDeviceName } else { $zoneInfo.Monitor.DisplayLabel }
                if ($controls.ContainsKey('TxtZoneTarget')) {
                    $controls.TxtZoneTarget.Text = "Zielzone: $($zoneInfo.Zone) auf $displayLabel"
                }
            } else {
                # Keine gültige Zone unter Maus
                if ($controls.ContainsKey('TxtZoneTarget')) {
                    $controls.TxtZoneTarget.Text = "Keine gültige Monitor-Zone unter Maus"
                }
            }
            
            $e.Handled = $true
        })

        $null = $zoneBlock.Add_MouseLeftButtonUp({
            param($sender, $e)
            if ($null -eq $script:dragChip -or $script:dragChip -ne $sender) { return }

            if ($sender.IsMouseCaptured) {
                $sender.ReleaseMouseCapture()
            }
            $sender.Opacity = 1

            $left = [double][System.Windows.Controls.Canvas]::GetLeft($sender)
            $top = [double][System.Windows.Controls.Canvas]::GetTop($sender)
            $width = if ($sender.ActualWidth -gt 0) { [double]$sender.ActualWidth } else { 170.0 }
            $height = if ($sender.ActualHeight -gt 0) { [double]$sender.ActualHeight } else { 90.0 }
            $centerX = $left + ($width / 2)
            $centerY = $top + ($height / 2)

            # Nutze die gleiche Logik wie in der Live-Preview
            $zoneInfo = Get-CanvasZoneAtPosition -CenterX $centerX -CenterY $centerY

            if ($zoneInfo) {
                $meta = $sender.Tag
                $profileName = Resolve-SelectedEditorProfile
                $updated = Update-WindowRulePlacement -Config $config -ProfileName $profileName -ProcessName ([string]$meta.ProcessName) -Title ([string]$meta.Title) -AssignedMonitor ([string]$zoneInfo.Monitor.DeviceName) -ZonePreset ([string]$zoneInfo.Zone)
                if ($updated) {
                    Save-Config -Config $config -Path $ConfigPath
                    Write-Log "Zuordnung verschoben: [$($meta.ProcessName)] -> $($zoneInfo.Zone) auf $($zoneInfo.Monitor.DisplayLabel)"
                    Write-EditorLog "Zuordnung verschoben: [$($meta.ProcessName)] -> $($zoneInfo.Zone) auf $($zoneInfo.Monitor.DisplayLabel)"
                }
            }
            
            # Stelle Monitor-Borders zurück
            foreach ($border in $controls.CanvasLayout.Children | Where-Object { $_ -is [System.Windows.Controls.Border] -and $_.Tag -is [string] }) {
                $border.BorderThickness = New-Object System.Windows.Thickness(2)
                $border.BorderBrush = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.Color]::FromRgb(14,165,233))
                $border.Opacity = 1.0
            }

            $script:dragChip = $null
            Render-LayoutCanvas
            $e.Handled = $true
        })

        $null = $zoneBlock.Add_MouseRightButtonUp({
            param($sender, $e)
            $meta = $sender.Tag
            if ($null -eq $meta) { return }

            $profileName = Resolve-SelectedEditorProfile
            $removed = Remove-WindowRuleFromProfile -Config $config -ProfileName $profileName -ProcessName ([string]$meta.ProcessName) -Title ([string]$meta.Title) -AssignedMonitor ([string]$meta.AssignedMonitor) -ZonePreset ([string]$meta.ZonePreset)
            if ($removed) {
                Save-Config -Config $config -Path $ConfigPath
                Write-Log "Zuordnung geloescht: [$($meta.ProcessName)] $($meta.Title)"
                Write-EditorLog "Zuordnung geloescht: [$($meta.ProcessName)] $($meta.Title)"
                Render-LayoutCanvas
            }
            $e.Handled = $true
        })

        $controls.CanvasLayout.Children.Add($zoneBlock) | Out-Null
    }
}

function Apply-ZoneToPendingDrop {
    param([Parameter(Mandatory)][string]$ZoneName)

    if (-not $script:pendingDrop) { return }
    $profileName = Resolve-SelectedEditorProfile
    Set-WindowRuleForProfile -Config $config -ProfileName $profileName -WindowItem $script:pendingDrop.WindowItem -ScreenItem $script:pendingDrop.ScreenItem -ZonePreset $ZoneName
    Save-Config -Config $config -Path $ConfigPath
    Update-ProfileList
    Refresh-DashboardCards
    Render-LayoutCanvas

    $targetName = if ([string]::IsNullOrWhiteSpace([string]$script:pendingDrop.ScreenItem.DisplayLabel)) { [string]$script:pendingDrop.ScreenItem.DeviceName } else { [string]$script:pendingDrop.ScreenItem.DisplayLabel }
    Write-Log "Regel gesetzt: [$($script:pendingDrop.WindowItem.ProcessName)] auf $targetName als $ZoneName im Profil $profileName."
    Write-EditorLog "Regel gesetzt: [$($script:pendingDrop.WindowItem.ProcessName)] auf $targetName als $ZoneName im Profil $profileName."
    $script:pendingDrop = $null
    if ($controls.ContainsKey('CanvasZonePreview') -and $null -ne $controls.CanvasZonePreview) {
        $controls.CanvasZonePreview.Children.Clear()
    }
    $controls.ZonePicker.Visibility = 'Collapsed'
}

function Switch-ToProfileDisplayMode {
    param([Parameter(Mandatory)][string]$ProfileName)
    $profile = Get-Profile -Config $config -Name $ProfileName
    if (-not $profile) { throw "Profil '$ProfileName' nicht gefunden." }
    Close-ProfileWindows -Config $config -ProfileName $ProfileName
    Set-DisplayMode -Mode $profile.DisplayMode
    Start-Sleep -Milliseconds ([int]$config.Settings.SwitchDelayMs)
    if ([bool]$config.Settings.RestoreAfterSwitch) {
        $result = Restore-Layout -Config $config -ProfileName $ProfileName -Detailed
        Write-Log "Layout fuer $ProfileName nach Moduswechsel: verschoben=$($result.Moved), gestartet=$($result.Launched), nicht gefunden=$($result.Missing), ungueltige Regeln=$($result.Invalid)."
    }
}

$controls.BtnNavDashboard.Add_Click({ Show-View -Name 'Dashboard' })
$controls.BtnNavEditor.Add_Click({ Show-View -Name 'Editor'; Render-LayoutCanvas })
$controls.BtnNavProfiles.Add_Click({ Show-View -Name 'Profiles'; Update-ProfileList })
$controls.BtnNavSettings.Add_Click({ Show-View -Name 'Settings'; Update-SettingsView })
$controls.BtnNavHelp.Add_Click({ Show-View -Name 'Help' })

$controls.BtnRefresh.Add_Click({
    try {
        Update-ScreenList
        Refresh-DashboardCards
        Render-LayoutCanvas
        Write-Log 'Monitor-Scan abgeschlossen.'
    }
    catch {
        Write-Log "Fehler bei Monitor-Scan: $($_.Exception.Message)"
    }
})

$controls.BtnHeaderReload.Add_Click({
    try {
        Restart-ToolProcess
    }
    catch {
        Write-Log "GUI-Neustart fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnSwitchAlltag.Add_Click({
    try {
        $profile = Get-Profile -Config $config -Name 'Alltag'
        if (-not $profile) { throw 'Profil Alltag fehlt.' }
        $config.ActiveProfile = 'Alltag'
        Save-Config -Config $config -Path $ConfigPath
        Switch-ToProfileDisplayMode -ProfileName 'Alltag'
        Update-ScreenList
        Refresh-DashboardCards
        Update-ProfileList
    }
    catch {
        Write-Log "Moduswechsel Alltag fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnSwitchStreaming.Add_Click({
    try {
        $profile = Get-Profile -Config $config -Name 'StreamingGaming'
        if (-not $profile) { throw 'Profil StreamingGaming fehlt.' }
        $config.ActiveProfile = 'StreamingGaming'
        Save-Config -Config $config -Path $ConfigPath
        Switch-ToProfileDisplayMode -ProfileName 'StreamingGaming'
        Update-ScreenList
        Refresh-DashboardCards
        Update-ProfileList
    }
    catch {
        Write-Log "Moduswechsel Streaming/Gaming fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnCaptureActive.Add_Click({
    try {
        $profileName = Get-ActiveProfileName -Config $config
        $count = Capture-Layout -Config $config -ProfileName $profileName
        Save-Config -Config $config -Path $ConfigPath
        Update-ProfileList
        Render-LayoutCanvas
        Write-Log "Layout fuer $profileName gespeichert ($count Fenster)."
    }
    catch {
        Write-Log "Layout speichern fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnApplyActive.Add_Click({
        Close-ProfileWindows -Config $config -ProfileName $profileName
    try {
        $profileName = Get-ActiveProfileName -Config $config
        $result = Restore-Layout -Config $config -ProfileName $profileName -Detailed
        Write-Log "Layout fuer $profileName angewendet: verschoben=$($result.Moved), gestartet=$($result.Launched), nicht gefunden=$($result.Missing), ungueltige Regeln=$($result.Invalid)."
    }
    catch {
        Write-Log "Layout anwenden fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnRescue.Add_Click({
    try {
        $moved = Rescue-WindowsToPrimary
        Write-Log "Fenster-Rettung abgeschlossen ($moved Fenster verschoben)."
    }
    catch {
        Write-Log "Fenster-Rettung fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnRefreshWindows.Add_Click({
    try {
        Refresh-OpenWindows
        Write-Log "Programm-Liste aktualisiert ($($script:openWindows.Count) Fenster)."
        Write-EditorLog "Programm-Liste aktualisiert ($($script:openWindows.Count) Fenster)."
    }
    catch {
        Write-Log "Programm-Scan fehlgeschlagen: $($_.Exception.Message)"
        Write-EditorLog "Programm-Scan fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.LbOpenWindows.Add_PreviewMouseMove({
    param($sender, $e)
    if ($e.LeftButton -ne [System.Windows.Input.MouseButtonState]::Pressed) { return }
    $item = $sender.SelectedItem
    if (-not $item) { return }
    [System.Windows.DragDrop]::DoDragDrop($sender, [string]$item.Key, [System.Windows.DragDropEffects]::Copy) | Out-Null
})

$controls.CmbEditorProfile.Add_SelectionChanged({
    try {
        $profileName = Resolve-SelectedEditorProfile
        Render-LayoutCanvas

        $profile = Get-Profile -Config $config -Name $profileName
        $count = if ($profile) { @($profile.WindowLayouts).Count } else { 0 }
        Write-Log "Layout-Editor aktualisiert fuer Profil $profileName (Regeln: $count)."
        Write-EditorLog "Layout-Editor aktualisiert fuer Profil $profileName (Regeln: $count)."
    }
    catch {
        Write-Log "Profilwechsel im Layout-Editor fehlgeschlagen: $($_.Exception.Message)"
        Write-EditorLog "Profilwechsel im Layout-Editor fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnEditorRefreshRules.Add_Click({
    try {
        $profileName = Resolve-SelectedEditorProfile
        Render-LayoutCanvas

        $profile = Get-Profile -Config $config -Name $profileName
        $count = if ($profile) { @($profile.WindowLayouts).Count } else { 0 }
        if ($count -eq 0) {
            Write-Log "Keine Regeln im Profil $profileName. Ziehe Apps links auf einen Monitor und waehle eine Zone."
            Write-EditorLog "Keine Regeln im Profil $profileName. Ziehe Apps links auf einen Monitor und waehle eine Zone."
        }
        else {
            Write-Log "Regeln angezeigt fuer Profil $profileName (Anzahl: $count)."
            Write-EditorLog "Regeln angezeigt fuer Profil $profileName (Anzahl: $count)."
        }
    }
    catch {
        Write-Log "Regeln anzeigen fehlgeschlagen: $($_.Exception.Message)"
        Write-EditorLog "Regeln anzeigen fehlgeschlagen: $($_.Exception.Message)"
    }
})
$controls.BtnEditorSaveProfile.Add_Click({
    try {
        $profileName = Resolve-SelectedEditorProfile
        $profile = Get-Profile -Config $config -Name $profileName
        if (-not $profile) { throw "Profil '$profileName' nicht gefunden." }

        Save-Config -Config $config -Path $ConfigPath
        Update-ProfileList
        Refresh-DashboardCards
        Write-Log "Profil gespeichert: $profileName (Regeln: $(@($profile.WindowLayouts).Count))."
        Write-EditorLog "Profil gespeichert: $profileName (Regeln: $(@($profile.WindowLayouts).Count))."
    }
    catch {
        Write-Log "Profil speichern fehlgeschlagen: $($_.Exception.Message)"
        Write-EditorLog "Profil speichern fehlgeschlagen: $($_.Exception.Message)"
    }
})
$controls.CanvasLayout.Add_SizeChanged({ Render-LayoutCanvas })
$controls.CanvasZonePreview.Add_SizeChanged({ if ($controls.ZonePicker.Visibility -eq 'Visible') { Render-ZonePreview } })

$controls.BtnZoneCancel.Add_Click({
    $script:pendingDrop = $null
    if ($controls.ContainsKey('CanvasZonePreview') -and $null -ne $controls.CanvasZonePreview) {
        $controls.CanvasZonePreview.Children.Clear()
    }
    $controls.ZonePicker.Visibility = 'Collapsed'
})

$controls.BtnAddProfile.Add_Click({
    try {
        $name = [Microsoft.VisualBasic.Interaction]::InputBox('Profilname eingeben:', 'Neues Profil', 'NeuesProfil')
        if ([string]::IsNullOrWhiteSpace($name)) { return }
        $name = $name.Trim()

        if ($config.Profiles | Where-Object { $_.Name -eq $name }) {
            throw 'Profilname existiert bereits.'
        }

        $mode = 'extend'
        if ($name -match 'alltag|work|office') { $mode = 'internal' }

        $config.Profiles += [ordered]@{
            Name = $name
            DisplayMode = $mode
            WindowLayouts = @()
        }
        Save-Config -Config $config -Path $ConfigPath
        Update-ProfileSelectors
        Update-ProfileList
        Write-Log "Profil erstellt: $name"
    }
    catch {
        Write-Log "Profil erstellen fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnDeleteProfile.Add_Click({
    try {
        $selected = $controls.LvProfiles.SelectedItem
        if (-not $selected) { throw 'Bitte ein Profil auswaehlen.' }
        $name = [string]$selected.Name

        if ($config.Profiles.Count -le 1) {
            throw 'Mindestens ein Profil muss bestehen bleiben.'
        }

        $config.Profiles = @($config.Profiles | Where-Object { $_.Name -ne $name })
        if ($config.ActiveProfile -eq $name) {
            $config.ActiveProfile = $config.Profiles[0].Name
        }

        Save-Config -Config $config -Path $ConfigPath
        Update-ProfileSelectors
        Update-ProfileList
        Refresh-DashboardCards
        Render-LayoutCanvas
        Write-Log "Profil geloescht: $name"
    }
    catch {
        Write-Log "Profil loeschen fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnSetActiveProfile.Add_Click({
    try {
        $selected = $controls.LvProfiles.SelectedItem
        if (-not $selected) { throw 'Bitte ein Profil auswaehlen.' }
        $config.ActiveProfile = [string]$selected.Name
        Save-Config -Config $config -Path $ConfigPath
        Update-ProfileSelectors
        Update-ProfileList
        Refresh-DashboardCards
        Write-Log "Aktives Profil gesetzt: $($config.ActiveProfile)"
    }
    catch {
        Write-Log "Aktives Profil setzen fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnCaptureProfile.Add_Click({
    try {
        $selected = $controls.LvProfiles.SelectedItem
        $name = if ($selected) { [string]$selected.Name } else { Get-ActiveProfileName -Config $config }
        if ([string]::IsNullOrWhiteSpace($name)) { throw 'Kein aktives Profil verfuegbar.' }
        $count = Capture-Layout -Config $config -ProfileName $name
        Save-Config -Config $config -Path $ConfigPath
        Update-ProfileList
        Render-LayoutCanvas
        Write-Log "Layout fuer $name gespeichert ($count Fenster)."
    }
    catch {
        Write-Log "Layout speichern fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnApplyProfile.Add_Click({
        Close-ProfileWindows -Config $config -ProfileName $name
    try {
        $selected = $controls.LvProfiles.SelectedItem
        $name = if ($selected) { [string]$selected.Name } else { Get-ActiveProfileName -Config $config }
        if ([string]::IsNullOrWhiteSpace($name)) { throw 'Kein aktives Profil verfuegbar.' }
        $result = Restore-Layout -Config $config -ProfileName $name -Detailed
        Write-Log "Layout fuer $name angewendet: verschoben=$($result.Moved), gestartet=$($result.Launched), nicht gefunden=$($result.Missing), ungueltige Regeln=$($result.Invalid)."
    }
    catch {
        Write-Log "Layout anwenden fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnSaveSettings.Add_Click({
    try {
        $delay = 0
        $launchDelay = 0
        if (-not [int]::TryParse($controls.TxtDelay.Text, [ref]$delay)) {
            throw 'Verzoegerung muss eine Zahl sein.'
        }
        if ($delay -lt 500 -or $delay -gt 20000) {
            throw 'Verzoegerung muss zwischen 500 und 20000 ms liegen.'
        }
        if (-not [int]::TryParse($controls.TxtLaunchDelay.Text, [ref]$launchDelay)) {
            throw 'Startwartezeit muss eine Zahl sein.'
        }
        if ($launchDelay -lt 0 -or $launchDelay -gt 20000) {
            throw 'Startwartezeit muss zwischen 0 und 20000 ms liegen.'
        }

        $config.Settings.RestoreAfterSwitch = [bool]$controls.ChkAutoRestore.IsChecked
        $config.Settings.AutoLaunchMissingWindows = [bool]$controls.ChkAutoLaunchMissing.IsChecked
        $config.Settings.SwitchDelayMs = $delay
        $config.Settings.LaunchDelayMs = $launchDelay
        $selectedLanguage = [string]$controls.CmbLanguage.SelectedValue
        if (@('de','en') -notcontains $selectedLanguage) {
            $selectedLanguage = 'de'
        }
        $config.Settings.UiLanguage = $selectedLanguage
        # Startup-Eintrag in Registry setzen oder entfernen (via launcher.bat)
        $regKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
        $regName = 'MultiMonitorProfileTool'
        if ([bool]$controls.ChkRunAtStartup.IsChecked) {
            $launcherBat = Join-Path -Path $PSScriptRoot -ChildPath 'launcher.bat'
            if (Test-Path -LiteralPath $launcherBat) {
                Set-ItemProperty -Path $regKey -Name $regName -Value "`"$launcherBat`"" -Type String
                $config.Settings.RunAtStartup = $true
                Write-Log 'Autostart aktiviert.'
            } else {
                throw "launcher.bat nicht gefunden: $launcherBat"
            }
        } else {
            Remove-ItemProperty -Path $regKey -Name $regName -ErrorAction SilentlyContinue
            $config.Settings.RunAtStartup = $false
            Write-Log 'Autostart deaktiviert.'
        }
        Save-Config -Config $config -Path $ConfigPath
        Update-SettingsView
        Write-Log 'Einstellungen gespeichert.'
    }
    catch {
        Write-Log "Settings speichern fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnExcludedAdd.Add_Click({
    try {
        $newProc = [string]$controls.TxtExcludedNew.Text
        if ([string]::IsNullOrWhiteSpace($newProc)) { return }
        $newProc = $newProc.Trim()
        if ($config.Settings.ExcludedProcesses -contains $newProc) { return }
        $config.Settings.ExcludedProcesses = @($config.Settings.ExcludedProcesses + $newProc)
        Save-Config -Config $config -Path $ConfigPath
        Update-SettingsView
        $controls.TxtExcludedNew.Text = ''
        Write-Log "Excluded process hinzugefuegt: $newProc"
    }
    catch {
        Write-Log "Hinzufuegen fehlgeschlagen: $($_.Exception.Message)"
    }
})

$controls.BtnExcludedRemove.Add_Click({
    try {
        $selected = [string]$controls.LbExcluded.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selected)) { return }
        $config.Settings.ExcludedProcesses = @($config.Settings.ExcludedProcesses | Where-Object { $_ -ne $selected })
        Save-Config -Config $config -Path $ConfigPath
        Update-SettingsView
        Write-Log "Excluded process entfernt: $selected"
    }
    catch {
        Write-Log "Entfernen fehlgeschlagen: $($_.Exception.Message)"
    }
})

if ($controls.ContainsKey('BtnHelpToDashboard') -and $null -ne $controls['BtnHelpToDashboard']) {
    $controls['BtnHelpToDashboard'].Add_Click({ Show-View -Name 'Dashboard' })
}
if ($controls.ContainsKey('BtnHelpToEditor') -and $null -ne $controls['BtnHelpToEditor']) {
    $controls['BtnHelpToEditor'].Add_Click({ Show-View -Name 'Editor'; Render-LayoutCanvas })
}

$controls.BtnExit.Add_Click({ $window.Close() })

Update-ScreenList
Update-ProfileSelectors
Update-ProfileList
Update-SettingsView
Refresh-DashboardCards
Refresh-OpenWindows
Render-LayoutCanvas
Show-View -Name 'Dashboard'
Write-Log (T 'log_tool_started')
Write-EditorLog (T 'editor_log_started')

$window.Add_ContentRendered({
    try {
        # Erzwingt einen sichtbaren Start auf dem Hauptmonitor und holt das Fenster in den Vordergrund.
        $primary = [System.Windows.Forms.Screen]::PrimaryScreen
        if ($null -ne $primary) {
            $work = $primary.WorkingArea
            $targetWidth = [Math]::Min([double]$window.Width, [double]($work.Width - 40))
            $targetHeight = [Math]::Min([double]$window.Height, [double]($work.Height - 40))

            $window.WindowStartupLocation = [System.Windows.WindowStartupLocation]::Manual
            $window.Left = $work.Left + [Math]::Max(0, [Math]::Floor(($work.Width - $targetWidth) / 2))
            $window.Top = $work.Top + [Math]::Max(0, [Math]::Floor(($work.Height - $targetHeight) / 2))
            $window.Width = $targetWidth
            $window.Height = $targetHeight
        }

        $window.WindowState = [System.Windows.WindowState]::Normal
        $window.ShowInTaskbar = $true
        $window.Topmost = $true
        $window.Activate() | Out-Null
        $window.Focus() | Out-Null
        $window.Topmost = $false
    }
    catch {
        Write-Log "Hinweis: Sichtbarkeits-Start konnte nicht vollständig angewendet werden: $($_.Exception.Message)"
    }
})

$null = $window.ShowDialog()
