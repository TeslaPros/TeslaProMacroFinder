param(
    [switch]$Deep
)

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "TeslaPro Macro Finder Ultra V2"

$Script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Script:BaseDir) { $Script:BaseDir = (Get-Location).Path }

$Script:ReportTxt  = Join-Path $Script:BaseDir "TeslaPro_Report.txt"
$Script:ReportJson = Join-Path $Script:BaseDir "TeslaPro_Report.json"
$Script:Inventory  = Join-Path $Script:BaseDir "TeslaPro_Inventory.txt"

$Script:Items = New-Object System.Collections.Generic.List[object]
$Script:Analyzed = New-Object System.Collections.Generic.List[object]
$Script:AllProcesses = @()

function Reset-Screen {
    Clear-Host
    try { [Console]::CursorVisible = $false } catch {}
}

function Restore-Screen {
    try { [Console]::CursorVisible = $true } catch {}
}

function Safe-Text {
    param([object]$Value)
    if ($null -eq $Value) { return "-" }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return "-" }
    return $s
}

function Cut-Text {
    param([string]$Text, [int]$Max = 72)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "-" }
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max - 3) + "..."
}

function Write-Color {
    param(
        [string]$Text,
        [ConsoleColor]$Color = [ConsoleColor]::Gray,
        [switch]$NoNewline
    )
    if ($NoNewline) {
        Write-Host $Text -ForegroundColor $Color -NoNewline
    } else {
        Write-Host $Text -ForegroundColor $Color
    }
}

function Write-Bar {
    param(
        [char]$Char = '═',
        [ConsoleColor]$Color = [ConsoleColor]::DarkCyan
    )
    $w = [Math]::Max(80, $Host.UI.RawUI.WindowSize.Width - 2)
    Write-Host ($Char.ToString() * $w) -ForegroundColor $Color
}

function Write-Header {
    param([string]$Title)
    Write-Bar
    Write-Color ("  " + $Title) Cyan
    Write-Bar
}

function Show-Splash {
    Reset-Screen
    $art = @(
        "████████╗███████╗███████╗██╗      █████╗ ██████╗ ██████╗  ██████╗ ",
        "╚══██╔══╝██╔════╝██╔════╝██║     ██╔══██╗██╔══██╗██╔══██╗██╔═══██╗",
        "   ██║   █████╗  ███████╗██║     ███████║██████╔╝██████╔╝██║   ██║",
        "   ██║   ██╔══╝  ╚════██║██║     ██╔══██║██╔═══╝ ██╔══██╗██║   ██║",
        "   ██║   ███████╗███████║███████╗██║  ██║██║     ██║  ██║╚██████╔╝",
        "   ╚═╝   ╚══════╝╚══════╝╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝  ╚═╝ ╚═════╝ ",
        "",
        "               Macro Finder Ultra V2"
    )
    foreach ($line in $art) {
        Write-Color $line Green
        Start-Sleep -Milliseconds 35
    }
    Write-Color ""
    Write-Color "  Advanced Macro / Script / Peripheral Scanner" DarkCyan
    Write-Color "  Deep Scan: $Deep" Gray
    Write-Color ""
    Start-Sleep -Milliseconds 300
}

function Spinner {
    param([string]$Text = "Loading", [int]$Loops = 16)
    $frames = @('|','/','-','\')
    for ($i = 0; $i -lt $Loops; $i++) {
        foreach ($f in $frames) {
            Write-Host ("`r  {0} {1}" -f $Text, $f) -ForegroundColor Yellow -NoNewline
            Start-Sleep -Milliseconds 55
        }
    }
    Write-Host ("`r  {0} done.     " -f $Text) -ForegroundColor Green
}

function Progress-Step {
    param([int]$Percent, [string]$Text)
    $width = 32
    $filled = [Math]::Floor(($Percent / 100) * $width)
    $bar = ('█' * $filled).PadRight($width, '░')
    Write-Host ("`r  [{0}] {1,3}%  {2}" -f $bar, $Percent, $Text) -NoNewline -ForegroundColor Green
}

function Pause-Key {
    Write-Color ""
    Write-Color "  Druk op een toets om terug te gaan..." DarkGray
    [void][System.Console]::ReadKey($true)
}

function Add-FoundItem {
    param(
        [string]$Category,
        [string]$Vendor,
        [string]$Name,
        [string]$Path,
        [string]$Source,
        [string]$Reason
    )
    $Script:Items.Add([PSCustomObject]@{
        Category = Safe-Text $Category
        Vendor   = Safe-Text $Vendor
        Name     = Safe-Text $Name
        Path     = Safe-Text $Path
        Source   = Safe-Text $Source
        Reason   = Safe-Text $Reason
    }) | Out-Null
}

function Get-Vendor {
    param([string]$Blob)
    $b = (Safe-Text $Blob).ToLowerInvariant()
    switch -Regex ($b) {
        'logitech|lghub|logi'        { return 'Logitech' }
        'razer|synapse'              { return 'Razer' }
        'autohotkey|\.ahk'           { return 'AutoHotkey' }
        'steelseries| gg|gg '        { return 'SteelSeries' }
        'corsair|icue'               { return 'Corsair' }
        'bloody|a4tech|x7'           { return 'Bloody/A4Tech' }
        'x-mouse|xmouse'             { return 'X-Mouse' }
        'lua|\.lua'                  { return 'Lua' }
        default                      { return 'Unknown' }
    }
}

function Get-KeywordHits {
    param([string]$Path)

    $keywords = @(
        'macro','macros','autohotkey','ahk','lua','rapidfire','autoclick','auto click',
        'toggle','minecraft','ghub','lghub','synapse','razer','logitech','hotkey',
        'sendinput','click','script','recoil','jitter','bind','autostrafe','autojump',
        'bhop','trigger','spamclick','doubleclick','keyspam','loop','repeat','turbo'
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "-" }
        $item = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($item.PSIsContainer) { return "-" }
        if ($item.Length -gt 4MB) { return "-" }

        $ext = $item.Extension.ToLowerInvariant()
        if ($ext -notin '.ahk','.lua','.txt','.cfg','.conf','.ini','.json','.xml','.log','.bat','.cmd','.ps1') {
            return "-"
        }

        $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $hits = foreach ($k in $keywords) {
            if ($content -match [regex]::Escape($k)) { $k }
        }
        $uniq = $hits | Sort-Object -Unique
        if (-not $uniq) { return "-" }
        return ($uniq -join ', ')
    }
    catch {
        return "-"
    }
}

function Get-RiskScore {
    param(
        [string]$Name,
        [string]$Path,
        [string]$Keywords,
        [string]$Vendor,
        [string]$Category
    )

    $blob = (($Name + ' ' + $Path + ' ' + $Keywords + ' ' + $Vendor + ' ' + $Category).ToLowerInvariant())
    $score = 0

    foreach ($k in @('macro','autoclick','rapidfire','toggle','recoil','sendinput','trigger','spamclick','turbo','bhop','autostrafe','autojump','jitter')) {
        if ($blob -match [regex]::Escape($k)) { $score += 11 }
    }

    foreach ($k in @('autohotkey','ahk','lua','.ahk','.lua')) {
        if ($blob -match [regex]::Escape($k)) { $score += 18 }
    }

    foreach ($k in @('logitech','razer','synapse','ghub','steelseries','corsair','bloody','x-mouse')) {
        if ($blob -match [regex]::Escape($k)) { $score += 6 }
    }

    if ($Category -match 'Process') { $score += 8 }
    if ($Category -match 'Script')  { $score += 12 }
    if ($Category -match 'DeepScanHit') { $score += 10 }

    if ($score -gt 100) { $score = 100 }
    return $score
}

function Get-RiskLabel {
    param([int]$Score)
    if ($Score -ge 70) { return 'HIGH' }
    if ($Score -ge 35) { return 'MEDIUM' }
    return 'LOW'
}

function Get-RiskColor {
    param([int]$Score)
    if ($Score -ge 70) { return [ConsoleColor]::Red }
    if ($Score -ge 35) { return [ConsoleColor]::Yellow }
    return [ConsoleColor]::Green
}

function Collect-Processes {
    $terms = @('lghub','logitech','razer','synapse','autohotkey','steelseries','corsair','icue','bloody','xmouse','lua')
    $Script:AllProcesses = Get-CimInstance Win32_Process

    foreach ($p in $Script:AllProcesses) {
        $name = Safe-Text $p.Name
        $exe  = Safe-Text $p.ExecutablePath
        $cmd  = Safe-Text $p.CommandLine
        $blob = ($name + ' ' + $exe + ' ' + $cmd).ToLowerInvariant()

        if ($terms | Where-Object { $blob.Contains($_) }) {
            Add-FoundItem -Category "Process" -Vendor (Get-Vendor $blob) -Name $name -Path $exe -Source "Running Processes" -Reason "Matched process name or command line"
        }
    }
}

function Collect-RegistryApps {
    $keys = @(
        'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )

    foreach ($key in $keys) {
        Get-ItemProperty $key | ForEach-Object {
            $dn = Safe-Text $_.DisplayName
            $il = Safe-Text $_.InstallLocation
            if ($dn -match 'Logitech|G HUB|Razer|Synapse|AutoHotkey|SteelSeries|GG|Corsair|iCUE|Bloody|X-Mouse|Lua|Macro') {
                Add-FoundItem -Category "InstalledApp" -Vendor (Get-Vendor "$dn $il") -Name $dn -Path $il -Source $key -Reason "Matched uninstall registry entry"
            }
        }
    }
}

function Collect-ProgramFolders {
    $roots = @($env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432) |
        Where-Object { $_ -and (Test-Path $_) } |
        Select-Object -Unique

    foreach ($root in $roots) {
        Get-ChildItem -LiteralPath $root -Directory | ForEach-Object {
            if ($_.Name -match 'Logitech|Razer|AutoHotkey|SteelSeries|Corsair|Bloody|X-Mouse|Lua') {
                Add-FoundItem -Category "ProgramFolder" -Vendor (Get-Vendor $_.FullName) -Name $_.Name -Path $_.FullName -Source $root -Reason "Matched program folder name"
            }
        }
    }
}

function Collect-Shortcuts {
    $roots = @(
        "$env:PUBLIC\Desktop",
        "$env:USERPROFILE\Desktop",
        "$env:APPDATA\Microsoft\Windows\Start Menu",
        "$env:ProgramData\Microsoft\Windows\Start Menu",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    ) | Where-Object { $_ -and (Test-Path $_) }

    foreach ($root in $roots) {
        Get-ChildItem -LiteralPath $root -Recurse -File -Include *.lnk,*.url | ForEach-Object {
            if ($_.Name -match 'Logitech|Razer|AutoHotkey|SteelSeries|Corsair|Bloody|X-Mouse|macro|minecraft') {
                Add-FoundItem -Category "Shortcut" -Vendor (Get-Vendor $_.FullName) -Name $_.Name -Path $_.FullName -Source $root -Reason "Matched shortcut name"
            }
        }
    }
}

function Collect-TargetFiles {
    $roots = @(
        "$env:LOCALAPPDATA\LGHUB",
        "$env:APPDATA\LGHUB",
        "$env:PROGRAMDATA\LGHUB",
        "$env:LOCALAPPDATA\Logitech",
        "$env:PROGRAMDATA\Logishrd",

        "$env:LOCALAPPDATA\Razer",
        "$env:APPDATA\Razer",
        "$env:PROGRAMDATA\Razer",

        "$env:LOCALAPPDATA\SteelSeries",
        "$env:APPDATA\SteelSeries",
        "$env:PROGRAMDATA\SteelSeries",

        "$env:LOCALAPPDATA\Corsair",
        "$env:APPDATA\Corsair",
        "$env:PROGRAMDATA\Corsair",

        "$env:USERPROFILE\Desktop",
        "$env:USERPROFILE\Documents",
        "$env:USERPROFILE\Downloads",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    $patterns = @('*.ahk','*.lua','*.bat','*.cmd','*.ps1','*.json','*.ini','*.cfg','*.conf','*.xml','*.txt','*.log','*.db')

    foreach ($root in $roots) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $root -Recurse -File -Filter $pat | ForEach-Object {
                $vendor = Get-Vendor $_.FullName
                $category = if ($_.Extension -match '\.(ahk|lua|bat|cmd|ps1)$') { "Script" } else { "PossibleConfig" }
                $reason = if ($_.Name -match 'macro|minecraft|autohotkey|ahk|lua|rapidfire|toggle|recoil|click|logitech|razer|synapse|ghub|steelseries|corsair|bloody|xmouse') {
                    "Matched suspicious file name"
                } else {
                    "Matched file pattern $pat"
                }

                Add-FoundItem -Category $category -Vendor $vendor -Name $_.Name -Path $_.FullName -Source $root -Reason $reason
            }
        }
    }
}

function Collect-DeepFiles {
    if (-not $Deep) { return }

    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -and (Test-Path $_.Root) }
    $patterns = @('*.ahk','*.lua','*macro*.txt','*macro*.ini','*macro*.cfg','*autoclick*.txt','*rapidfire*.txt','*minecraft*.txt')

    foreach ($drive in $drives) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $drive.Root -Recurse -File -Filter $pat | ForEach-Object {
                Add-FoundItem -Category "DeepScanHit" -Vendor (Get-Vendor $_.FullName) -Name $_.Name -Path $_.FullName -Source $drive.Root -Reason "Deep scan pattern match"
            }
        }
    }
}

function Analyze-Item {
    param([object]$Item)

    $exists = 'NO'
    $created = '-'
    $modified = '-'
    $accessed = '-'
    $size = '-'
    $keywords = '-'
    $running = 'NO'

    if ($Item.Category -eq 'Process') { $running = 'YES' }

    if ($Item.Path -ne '-' -and (Test-Path -LiteralPath $Item.Path)) {
        $exists = 'YES'
        try {
            $it = Get-Item -LiteralPath $Item.Path -ErrorAction Stop
            $created = $it.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
            $modified = $it.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            $accessed = $it.LastAccessTime.ToString('yyyy-MM-dd HH:mm:ss')
            if (-not $it.PSIsContainer) { $size = [string]$it.Length }
            $keywords = Get-KeywordHits -Path $Item.Path
        } catch {}
    }

    $score = Get-RiskScore -Name $Item.Name -Path $Item.Path -Keywords $keywords -Vendor $Item.Vendor -Category $Item.Category
    $label = Get-RiskLabel -Score $score

    [PSCustomObject]@{
        Category  = $Item.Category
        Vendor    = $Item.Vendor
        Name      = $Item.Name
        Path      = $Item.Path
        Source    = $Item.Source
        Reason    = $Item.Reason
        Exists    = $exists
        Created   = $created
        Modified  = $modified
        Accessed  = $accessed
        Size      = $size
        Keywords  = $keywords
        Running   = $running
        RiskScore = $score
        RiskLabel = $label
    }
}

function Export-Reports {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("TeslaPro Macro Finder Ultra V2 Report")
    $lines.Add(("Generated: " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')))
    $lines.Add("")

    $lines.Add("SYSTEM SUMMARY")
    $lines.Add("----------------------------------------------------------------")
    $lines.Add(("Found items        : " + $Script:Analyzed.Count))
    $lines.Add(("Existing items     : " + (($Script:Analyzed | Where-Object Exists -eq 'YES').Count)))
    $lines.Add(("Missing items      : " + (($Script:Analyzed | Where-Object Exists -eq 'NO').Count)))
    $lines.Add(("Running processes  : " + (($Script:Analyzed | Where-Object Running -eq 'YES').Count)))
    $lines.Add(("High risk          : " + (($Script:Analyzed | Where-Object RiskLabel -eq 'HIGH').Count)))
    $lines.Add(("Medium risk        : " + (($Script:Analyzed | Where-Object RiskLabel -eq 'MEDIUM').Count)))
    $lines.Add(("Low risk           : " + (($Script:Analyzed | Where-Object RiskLabel -eq 'LOW').Count)))
    $lines.Add("")

    $lines.Add("DETAILED RESULTS")
    $lines.Add("----------------------------------------------------------------")
    foreach ($x in $Script:Analyzed | Sort-Object RiskScore -Descending, Vendor, Name) {
        $lines.Add(("[{0}] {1}" -f $x.Vendor, $x.Name))
        $lines.Add(("  Type      : " + $x.Category))
        $lines.Add(("  Risk      : " + $x.RiskLabel + " (" + $x.RiskScore + ")"))
        $lines.Add(("  Running   : " + $x.Running))
        $lines.Add(("  Exists    : " + $x.Exists))
        $lines.Add(("  Created   : " + $x.Created))
        $lines.Add(("  Accessed  : " + $x.Accessed))
        $lines.Add(("  Modified  : " + $x.Modified))
        $lines.Add(("  Size      : " + $x.Size))
        $lines.Add(("  Keywords  : " + $x.Keywords))
        $lines.Add(("  Path      : " + $x.Path))
        $lines.Add(("  Source    : " + $x.Source))
        $lines.Add(("  Reason    : " + $x.Reason))
        $lines.Add("----------------------------------------------------------------")
    }

    $lines | Set-Content -LiteralPath $Script:ReportTxt -Encoding UTF8

    $Script:Analyzed |
        Sort-Object RiskScore -Descending, Vendor, Name |
        ConvertTo-Json -Depth 5 |
        Set-Content -LiteralPath $Script:ReportJson -Encoding UTF8

    $Script:Analyzed | ForEach-Object {
        "$($_.Category)|$($_.Vendor)|$($_.Name)|$($_.Path)|$($_.Exists)|$($_.Created)|$($_.Modified)|$($_.Accessed)|$($_.Size)|$($_.Keywords)|$($_.RiskLabel)|$($_.RiskScore)|$($_.Reason)"
    } | Set-Content -LiteralPath $Script:Inventory -Encoding UTF8
}

function Draw-Tabs {
    param([string]$Active)

    $tabs = @(
        @{ K = "1"; T = "Summary" },
        @{ K = "2"; T = "Apps" },
        @{ K = "3"; T = "Scripts" },
        @{ K = "4"; T = "High-Risk" },
        @{ K = "5"; T = "All" },
        @{ K = "6"; T = "Export" },
        @{ K = "Q"; T = "Quit" }
    )

    foreach ($tab in $tabs) {
        $label = " [$($tab.K)] $($tab.T) "
        if ($tab.T -eq $Active) {
            Write-Color $label Black -NoNewline
            $Host.UI.RawUI.BackgroundColor = "DarkCyan"
            Clear-Host
        }
    }
}

function Show-TopTabs {
    param([string]$Active)

    $tabs = @(
        @{ K = "1"; T = "Summary" },
        @{ K = "2"; T = "Apps" },
        @{ K = "3"; T = "Scripts" },
        @{ K = "4"; T = "High-Risk" },
        @{ K = "5"; T = "All" },
        @{ K = "6"; T = "Export" },
        @{ K = "Q"; T = "Quit" }
    )

    foreach ($tab in $tabs) {
        $text = " [$($tab.K)] $($tab.T) "
        if ($tab.T -eq $Active) {
            Write-Host $text -ForegroundColor Black -BackgroundColor DarkCyan -NoNewline
        } else {
            Write-Host $text -ForegroundColor Gray -BackgroundColor Black -NoNewline
        }
    }
    Write-Host ""
    Write-Bar '─' DarkGray
}

function Show-Summary {
    Reset-Screen
    Write-Header "TeslaPro Macro Finder Ultra V2"
    Show-TopTabs "Summary"

    $all = $Script:Analyzed
    $total = $all.Count
    $exists = ($all | Where-Object Exists -eq 'YES').Count
    $missing = ($all | Where-Object Exists -eq 'NO').Count
    $running = ($all | Where-Object Running -eq 'YES').Count
    $scripts = ($all | Where-Object { $_.Category -match 'Script|PossibleConfig|DeepScanHit' }).Count
    $apps = ($all | Where-Object { $_.Category -match 'InstalledApp|ProgramFolder|Process|Shortcut' }).Count
    $high = ($all | Where-Object RiskLabel -eq 'HIGH').Count
    $med = ($all | Where-Object RiskLabel -eq 'MEDIUM').Count
    $low = ($all | Where-Object RiskLabel -eq 'LOW').Count

    Write-Color ""
    Write-Color "  Scan overzicht" Cyan
    Write-Bar '─' DarkGray
    Write-Color ("  Total items      : {0}" -f $total) White
    Write-Color ("  Existing         : {0}" -f $exists) Green
    Write-Color ("  Missing/Deleted  : {0}" -f $missing) Yellow
    Write-Color ("  Running          : {0}" -f $running) Cyan
    Write-Color ("  Apps/Tools       : {0}" -f $apps) Magenta
    Write-Color ("  Scripts/Configs  : {0}" -f $scripts) DarkCyan
    Write-Color ""
    Write-Color ("  LOW risk         : {0}" -f $low) Green
    Write-Color ("  MEDIUM risk      : {0}" -f $med) Yellow
    Write-Color ("  HIGH risk        : {0}" -f $high) Red
    Write-Color ""

    Write-Color "  Top vendors" Cyan
    Write-Bar '─' DarkGray
    $vendors = $all | Group-Object Vendor | Sort-Object Count -Descending
    if (-not $vendors) {
        Write-Color "  Geen items gevonden." DarkGray
    } else {
        foreach ($v in $vendors | Select-Object -First 12) {
            Write-Color ("  {0,-18} {1,4}" -f $v.Name, $v.Count) Gray
        }
    }

    Write-Color ""
    Write-Color "  Gebruik de tabs bovenaan om te navigeren." DarkGray
}

function Show-CardPage {
    param(
        [string]$TabName,
        [string]$Title,
        [object[]]$Rows
    )

    Reset-Screen
    Write-Header "TeslaPro Macro Finder Ultra V2"
    Show-TopTabs $TabName
    Write-Color ("  " + $Title) Cyan
    Write-Bar '─' DarkGray
    Write-Color ""

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Color "  Geen resultaten." DarkGray
        Pause-Key
        return
    }

    $index = 0
    foreach ($x in $Rows) {
        $index++
        $riskColor = Get-RiskColor $x.RiskScore

        Write-Color ("  ┌─ Item #{0}" -f $index) DarkGray
        Write-Color ("  │ Name     : {0}" -f (Cut-Text $x.Name 68)) White
        Write-Color ("  │ Vendor   : {0}" -f $x.Vendor) Gray
        Write-Host ("  │ Risk     : {0} ({1})" -f $x.RiskLabel, $x.RiskScore) -ForegroundColor $riskColor
        Write-Color ("  │ Type     : {0}" -f $x.Category) Gray
        Write-Color ("  │ Running  : {0}" -f $x.Running) Gray
        Write-Color ("  │ Exists   : {0}" -f $x.Exists) Gray
        Write-Color ("  │ Created  : {0}" -f $x.Created) Gray
        Write-Color ("  │ Accessed : {0}" -f $x.Accessed) Gray
        Write-Color ("  │ Modified : {0}" -f $x.Modified) Gray
        Write-Color ("  │ Size     : {0}" -f $x.Size) Gray
        Write-Color ("  │ Keywords : {0}" -f (Cut-Text $x.Keywords 68)) Gray
        Write-Color ("  │ Path     : {0}" -f (Cut-Text $x.Path 68)) DarkGray
        Write-Color ("  │ Reason   : {0}" -f (Cut-Text $x.Reason 68)) DarkGray
        Write-Color "  └──────────────────────────────────────────────────────────────" DarkGray
        Write-Color ""
    }

    Pause-Key
}

function Show-ExportInfo {
    Reset-Screen
    Write-Header "TeslaPro Macro Finder Ultra V2"
    Show-TopTabs "Export"
    Write-Color ""
    Write-Color "  Export locaties" Cyan
    Write-Bar '─' DarkGray
    Write-Color ("  TXT  : " + $Script:ReportTxt) Green
    Write-Color ("  JSON : " + $Script:ReportJson) Green
    Write-Color ("  INV  : " + $Script:Inventory) Green
    Write-Color ""
    Write-Color "  Opmerking" Yellow
    Write-Bar '─' DarkGray
    Write-Color "  • Last Accessed kan onbetrouwbaar zijn op Windows." Gray
    Write-Color "  • Last Modified is meestal de meest bruikbare tijd." Gray
    Write-Color "  • Sommige apps gebruiken binary/db opslag." Gray
    Pause-Key
}

function Run-Scan {
    Reset-Screen
    Write-Header "TeslaPro Scan Engine V2"
    Write-Color ""
    Spinner "TeslaPro scan modules initialiseren" 10

    $steps = @(
        @{ P = 10; T = "Running processes analyseren"; A = { Collect-Processes } },
        @{ P = 24; T = "Registry apps controleren"; A = { Collect-RegistryApps } },
        @{ P = 40; T = "Program folders scannen"; A = { Collect-ProgramFolders } },
        @{ P = 56; T = "Shortcuts en startup zoeken"; A = { Collect-Shortcuts } },
        @{ P = 76; T = "Scripts en configs inspecteren"; A = { Collect-TargetFiles } },
        @{ P = 88; T = "Deep scan uitvoeren"; A = { Collect-DeepFiles } },
        @{ P = 96; T = "Resultaten analyseren"; A = {
            $unique = $Script:Items | Sort-Object Category, Vendor, Name, Path -Unique
            foreach ($u in $unique) {
                $Script:Analyzed.Add((Analyze-Item -Item $u)) | Out-Null
            }
        }},
        @{ P = 100; T = "Rapporten exporteren"; A = { Export-Reports } }
    )

    foreach ($step in $steps) {
        Progress-Step -Percent $step.P -Text $step.T
        & $step.A
        Start-Sleep -Milliseconds 120
    }

    Write-Host ""
    Write-Color ""
    Write-Color "  Scan voltooid." Green
    Start-Sleep -Milliseconds 400
}

function Main-Menu {
    while ($true) {
        Show-Summary
        $key = [System.Console]::ReadKey($true)

        switch ($key.Key) {
            'D1' { Show-Summary; Pause-Key }
            'NumPad1' { Show-Summary; Pause-Key }

            'D2' {
                $rows = $Script:Analyzed |
                    Where-Object { $_.Category -match 'InstalledApp|ProgramFolder|Process|Shortcut' } |
                    Sort-Object RiskScore -Descending, Vendor, Name
                Show-CardPage -TabName "Apps" -Title "Apps / Processen / Shortcuts" -Rows $rows
            }
            'NumPad2' {
                $rows = $Script:Analyzed |
                    Where-Object { $_.Category -match 'InstalledApp|ProgramFolder|Process|Shortcut' } |
                    Sort-Object RiskScore -Descending, Vendor, Name
                Show-CardPage -TabName "Apps" -Title "Apps / Processen / Shortcuts" -Rows $rows
            }

            'D3' {
                $rows = $Script:Analyzed |
                    Where-Object { $_.Category -match 'Script|PossibleConfig|DeepScanHit' } |
                    Sort-Object RiskScore -Descending, Vendor, Name
                Show-CardPage -TabName "Scripts" -Title "Scripts / Configs / Deep Scan Hits" -Rows $rows
            }
            'NumPad3' {
                $rows = $Script:Analyzed |
                    Where-Object { $_.Category -match 'Script|PossibleConfig|DeepScanHit' } |
                    Sort-Object RiskScore -Descending, Vendor, Name
                Show-CardPage -TabName "Scripts" -Title "Scripts / Configs / Deep Scan Hits" -Rows $rows
            }

            'D4' {
                $rows = $Script:Analyzed |
                    Where-Object { $_.RiskLabel -eq 'HIGH' } |
                    Sort-Object RiskScore -Descending, Vendor, Name
                Show-CardPage -TabName "High-Risk" -Title "High Risk Items" -Rows $rows
            }
            'NumPad4' {
                $rows = $Script:Analyzed |
                    Where-Object { $_.RiskLabel -eq 'HIGH' } |
                    Sort-Object RiskScore -Descending, Vendor, Name
                Show-CardPage -TabName "High-Risk" -Title "High Risk Items" -Rows $rows
            }

            'D5' {
                $rows = $Script:Analyzed | Sort-Object RiskScore -Descending, Vendor, Name
                Show-CardPage -TabName "All" -Title "Alle Resultaten" -Rows $rows
            }
            'NumPad5' {
                $rows = $Script:Analyzed | Sort-Object RiskScore -Descending, Vendor, Name
                Show-CardPage -TabName "All" -Title "Alle Resultaten" -Rows $rows
            }

            'D6' { Show-ExportInfo }
            'NumPad6' { Show-ExportInfo }

            'Q' { break }
            'Escape' { break }
        }
    }
}

try {
    Show-Splash
    Run-Scan
    Main-Menu
}
finally {
    Restore-Screen
    Reset-Screen
    Write-Header "TeslaPro Macro Finder Ultra V2"
    Write-Color ""
    Write-Color "  Scan afgerond." Green
    Write-Color ("  TXT  : " + $Script:ReportTxt) Gray
    Write-Color ("  JSON : " + $Script:ReportJson) Gray
    Write-Color ("  INV  : " + $Script:Inventory) Gray
    Write-Color ""
}