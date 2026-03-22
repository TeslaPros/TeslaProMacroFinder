param([switch]$Deep)

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "TeslaPro Macro Finder Ultra V3"

# ==============================
# Paths
# ==============================
$Script:BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($Script:BaseDir)) {
    $Script:BaseDir = (Get-Location).Path
}

$Script:ReportTxt  = Join-Path $Script:BaseDir "TeslaPro_Report.txt"
$Script:ReportJson = Join-Path $Script:BaseDir "TeslaPro_Report.json"
$Script:Inventory  = Join-Path $Script:BaseDir "TeslaPro_Inventory.txt"

$Script:Items     = New-Object System.Collections.Generic.List[object]
$Script:Analyzed  = New-Object System.Collections.Generic.List[object]

# ==============================
# UI
# ==============================
function Clear-UI {
    Clear-Host
}

function Write-Bar {
    param(
        [string]$Char = "=",
        [ConsoleColor]$Color = [ConsoleColor]::DarkCyan
    )
    $w = [Math]::Max(70, $Host.UI.RawUI.WindowSize.Width - 1)
    Write-Host ($Char * $w) -ForegroundColor $Color
}

function Write-Title {
    Clear-UI
    Write-Bar
    Write-Host "              TeslaPro Macro Finder Ultra V3" -ForegroundColor Green
    Write-Host "          Advanced Macro / Script / Peripheral Scan" -ForegroundColor Cyan
    Write-Bar
    Write-Host ""
}

function Spinner {
    param(
        [string]$Text = "Loading",
        [int]$Loops = 18
    )
    $frames = @('|','/','-','\')
    for ($i = 0; $i -lt $Loops; $i++) {
        foreach ($f in $frames) {
            Write-Host "`r$Text $f" -NoNewline -ForegroundColor Yellow
            Start-Sleep -Milliseconds 55
        }
    }
    Write-Host "`r$Text done.   " -ForegroundColor Green
}

function Show-Progress {
    param(
        [int]$Percent,
        [string]$Text
    )
    $width = 34
    $filled = [Math]::Floor(($Percent / 100) * $width)
    $bar = ('#' * $filled).PadRight($width, '.')
    Write-Host "`r[$bar] $Percent%  $Text" -NoNewline -ForegroundColor Green
}

function Pause-Key {
    Write-Host ""
    Write-Host "Druk op een toets om verder te gaan..." -ForegroundColor DarkGray
    [void][System.Console]::ReadKey($true)
}

function Show-Tabs {
    param([string]$Active)

    $tabs = @(
        @{K='1';T='Summary'},
        @{K='2';T='Apps'},
        @{K='3';T='Scripts'},
        @{K='4';T='High-Risk'},
        @{K='5';T='All'},
        @{K='6';T='Export'},
        @{K='Q';T='Quit'}
    )

    foreach ($tab in $tabs) {
        $label = " [$($tab.K)] $($tab.T) "
        if ($tab.T -eq $Active) {
            Write-Host $label -NoNewline -ForegroundColor Black -BackgroundColor DarkCyan
        } else {
            Write-Host $label -NoNewline -ForegroundColor Gray -BackgroundColor Black
        }
    }
    Write-Host ""
    Write-Bar "-" DarkGray
}

# ==============================
# Helpers
# ==============================
function Safe-Text {
    param([object]$Value)
    if ($null -eq $Value) { return "-" }
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return "-" }
    return $s
}

function Cut-Text {
    param(
        [string]$Text,
        [int]$Max = 76
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return "-" }
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max - 3) + "..."
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
        'logitech|lghub|logi'       { return 'Logitech' }
        'razer|synapse'             { return 'Razer' }
        'autohotkey|\.ahk'          { return 'AutoHotkey' }
        'corsair|icue'              { return 'Corsair' }
        'steelseries|gg'            { return 'SteelSeries' }
        'bloody|a4tech|x7'          { return 'Bloody/A4Tech' }
        'x-mouse|xmouse'            { return 'X-Mouse' }
        'lua|\.lua'                 { return 'Lua' }
        default                     { return 'Unknown' }
    }
}

function Get-KeywordHits {
    param([string]$Path)

    $keywords = @(
        'macro','macros','autohotkey','ahk','lua','rapidfire','autoclick','auto click',
        'toggle','minecraft','ghub','lghub','synapse','razer','logitech','hotkey',
        'sendinput','click','script','recoil','jitter','bind','loop','repeat','turbo',
        'spamclick','trigger','bhop','autostrafe','autojump'
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

    $blob = (($Name + " " + $Path + " " + $Keywords + " " + $Vendor + " " + $Category).ToLowerInvariant())
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
    if ($Category -match 'Script') { $score += 12 }
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

function Sort-Rows {
    param([object[]]$Rows)
    return @($Rows | Sort-Object @{Expression='RiskScore';Descending=$true}, Vendor, Name)
}

# ==============================
# Collectors
# ==============================
function Collect-Processes {
    $terms = @('lghub','logitech','razer','synapse','autohotkey','steelseries','corsair','icue','bloody','xmouse','lua')
    $procs = Get-CimInstance Win32_Process

    foreach ($p in $procs) {
        $name = Safe-Text $p.Name
        $exe  = Safe-Text $p.ExecutablePath
        $cmd  = Safe-Text $p.CommandLine
        $blob = ($name + ' ' + $exe + ' ' + $cmd).ToLowerInvariant()

        if ($terms | Where-Object { $blob.Contains($_) }) {
            Add-FoundItem -Category 'Process' -Vendor (Get-Vendor $blob) -Name $name -Path $exe -Source 'Running Processes' -Reason 'Matched process name or command line'
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
                Add-FoundItem -Category 'InstalledApp' -Vendor (Get-Vendor "$dn $il") -Name $dn -Path $il -Source $key -Reason 'Matched uninstall registry entry'
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
                Add-FoundItem -Category 'ProgramFolder' -Vendor (Get-Vendor $_.FullName) -Name $_.Name -Path $_.FullName -Source $root -Reason 'Matched program folder name'
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
                Add-FoundItem -Category 'Shortcut' -Vendor (Get-Vendor $_.FullName) -Name $_.Name -Path $_.FullName -Source $root -Reason 'Matched shortcut name'
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
                $category = if ($_.Extension -match '\.(ahk|lua|bat|cmd|ps1)$') { 'Script' } else { 'PossibleConfig' }
                $reason = if ($_.Name -match 'macro|minecraft|autohotkey|ahk|lua|rapidfire|toggle|recoil|click|logitech|razer|synapse|ghub|steelseries|corsair|bloody|xmouse') {
                    'Matched suspicious file name'
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
                Add-FoundItem -Category 'DeepScanHit' -Vendor (Get-Vendor $_.FullName) -Name $_.Name -Path $_.FullName -Source $drive.Root -Reason 'Deep scan pattern match'
            }
        }
    }
}

# ==============================
# Analyze
# ==============================
function Analyze-Item {
    param([object]$Item)

    $exists = 'NO'
    $created = '-'
    $modified = '-'
    $accessed = '-'
    $size = '-'
    $keywords = '-'
    $running = 'NO'

    if ($Item.Category -eq 'Process') {
        $running = 'YES'
    }

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

# ==============================
# Export
# ==============================
function Export-Reports {
    $sorted = Sort-Rows $Script:Analyzed

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("TeslaPro Macro Finder Ultra V3 Report")
    $lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add("")

    $lines.Add("SYSTEM SUMMARY")
    $lines.Add("----------------------------------------------------------------")
    $lines.Add("Found items        : $($sorted.Count)")
    $lines.Add("Existing items     : $(($sorted | Where-Object Exists -eq 'YES').Count)")
    $lines.Add("Missing items      : $(($sorted | Where-Object Exists -eq 'NO').Count)")
    $lines.Add("Running processes  : $(($sorted | Where-Object Running -eq 'YES').Count)")
    $lines.Add("High risk          : $(($sorted | Where-Object RiskLabel -eq 'HIGH').Count)")
    $lines.Add("Medium risk        : $(($sorted | Where-Object RiskLabel -eq 'MEDIUM').Count)")
    $lines.Add("Low risk           : $(($sorted | Where-Object RiskLabel -eq 'LOW').Count)")
    $lines.Add("")

    $lines.Add("DETAILED RESULTS")
    $lines.Add("----------------------------------------------------------------")

    foreach ($x in $sorted) {
        $lines.Add("[$($x.Vendor)] $($x.Name)")
        $lines.Add("  Type      : $($x.Category)")
        $lines.Add("  Risk      : $($x.RiskLabel) ($($x.RiskScore))")
        $lines.Add("  Running   : $($x.Running)")
        $lines.Add("  Exists    : $($x.Exists)")
        $lines.Add("  Created   : $($x.Created)")
        $lines.Add("  Accessed  : $($x.Accessed)")
        $lines.Add("  Modified  : $($x.Modified)")
        $lines.Add("  Size      : $($x.Size)")
        $lines.Add("  Keywords  : $($x.Keywords)")
        $lines.Add("  Path      : $($x.Path)")
        $lines.Add("  Source    : $($x.Source)")
        $lines.Add("  Reason    : $($x.Reason)")
        $lines.Add("----------------------------------------------------------------")
    }

    $lines | Set-Content -LiteralPath $Script:ReportTxt -Encoding UTF8

    $sorted | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Script:ReportJson -Encoding UTF8

    $sorted | ForEach-Object {
        "$($_.Category)|$($_.Vendor)|$($_.Name)|$($_.Path)|$($_.Exists)|$($_.Created)|$($_.Modified)|$($_.Accessed)|$($_.Size)|$($_.Keywords)|$($_.RiskLabel)|$($_.RiskScore)|$($_.Reason)"
    } | Set-Content -LiteralPath $Script:Inventory -Encoding UTF8
}

# ==============================
# Pages
# ==============================
function Show-Summary {
    Write-Title
    Show-Tabs 'Summary'

    $rows = $Script:Analyzed
    $total   = $rows.Count
    $exists  = ($rows | Where-Object Exists -eq 'YES').Count
    $missing = ($rows | Where-Object Exists -eq 'NO').Count
    $running = ($rows | Where-Object Running -eq 'YES').Count
    $apps    = ($rows | Where-Object { $_.Category -match 'InstalledApp|ProgramFolder|Process|Shortcut' }).Count
    $scripts = ($rows | Where-Object { $_.Category -match 'Script|PossibleConfig|DeepScanHit' }).Count
    $high    = ($rows | Where-Object RiskLabel -eq 'HIGH').Count
    $med     = ($rows | Where-Object RiskLabel -eq 'MEDIUM').Count
    $low     = ($rows | Where-Object RiskLabel -eq 'LOW').Count

    Write-Host "Scan overzicht" -ForegroundColor Cyan
    Write-Bar "-" DarkGray
    Write-Host "Total items      : $total" -ForegroundColor White
    Write-Host "Existing         : $exists" -ForegroundColor Green
    Write-Host "Missing/Deleted  : $missing" -ForegroundColor Yellow
    Write-Host "Running          : $running" -ForegroundColor Cyan
    Write-Host "Apps/Tools       : $apps" -ForegroundColor Magenta
    Write-Host "Scripts/Configs  : $scripts" -ForegroundColor DarkCyan
    Write-Host ""
    Write-Host "LOW risk         : $low" -ForegroundColor Green
    Write-Host "MEDIUM risk      : $med" -ForegroundColor Yellow
    Write-Host "HIGH risk        : $high" -ForegroundColor Red
    Write-Host ""
    Write-Host "Top vendors" -ForegroundColor Cyan
    Write-Bar "-" DarkGray

    $vendors = $rows | Group-Object Vendor | Sort-Object Count -Descending
    if (-not $vendors) {
        Write-Host "Geen items gevonden." -ForegroundColor DarkGray
    } else {
        foreach ($v in $vendors | Select-Object -First 12) {
            Write-Host ("{0,-18} {1,4}" -f $v.Name, $v.Count) -ForegroundColor Gray
        }
    }

    Write-Host ""
    Write-Host "Gebruik de tabs bovenaan om te navigeren." -ForegroundColor DarkGray
}

function Show-CardPage {
    param(
        [string]$TabName,
        [string]$TitleText,
        [object[]]$Rows
    )

    Write-Title
    Show-Tabs $TabName
    Write-Host $TitleText -ForegroundColor Cyan
    Write-Bar "-" DarkGray
    Write-Host ""

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Host "Geen resultaten." -ForegroundColor DarkGray
        Pause-Key
        return
    }

    $i = 0
    foreach ($x in $Rows) {
        $i++
        $riskColor = Get-RiskColor $x.RiskScore

        Write-Host "┌─ Item #$i" -ForegroundColor DarkGray
        Write-Host ("│ Name     : {0}" -f (Cut-Text $x.Name 74)) -ForegroundColor White
        Write-Host ("│ Vendor   : {0}" -f $x.Vendor) -ForegroundColor Gray
        Write-Host ("│ Risk     : {0} ({1})" -f $x.RiskLabel, $x.RiskScore) -ForegroundColor $riskColor
        Write-Host ("│ Type     : {0}" -f $x.Category) -ForegroundColor Gray
        Write-Host ("│ Running  : {0}" -f $x.Running) -ForegroundColor Gray
        Write-Host ("│ Exists   : {0}" -f $x.Exists) -ForegroundColor Gray
        Write-Host ("│ Created  : {0}" -f $x.Created) -ForegroundColor Gray
        Write-Host ("│ Accessed : {0}" -f $x.Accessed) -ForegroundColor Gray
        Write-Host ("│ Modified : {0}" -f $x.Modified) -ForegroundColor Gray
        Write-Host ("│ Size     : {0}" -f $x.Size) -ForegroundColor Gray
        Write-Host ("│ Keywords : {0}" -f (Cut-Text $x.Keywords 74)) -ForegroundColor Gray
        Write-Host ("│ Path     : {0}" -f (Cut-Text $x.Path 74)) -ForegroundColor DarkGray
        Write-Host ("│ Reason   : {0}" -f (Cut-Text $x.Reason 74)) -ForegroundColor DarkGray
        Write-Host "└────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host ""
    }

    Pause-Key
}

function Show-ExportInfo {
    Write-Title
    Show-Tabs 'Export'

    Write-Host "Export locaties" -ForegroundColor Cyan
    Write-Bar "-" DarkGray
    Write-Host "TXT  : $Script:ReportTxt" -ForegroundColor Green
    Write-Host "JSON : $Script:ReportJson" -ForegroundColor Green
    Write-Host "INV  : $Script:Inventory" -ForegroundColor Green
    Write-Host ""
    Write-Host "Opmerking" -ForegroundColor Yellow
    Write-Bar "-" DarkGray
    Write-Host "Last Accessed kan onbetrouwbaar zijn op Windows." -ForegroundColor Gray
    Write-Host "Last Modified is meestal de meest bruikbare tijd." -ForegroundColor Gray
    Write-Host "Sommige apps gebruiken binary/db opslag." -ForegroundColor Gray

    Pause-Key
}

# ==============================
# Main scan
# ==============================
function Run-Scan {
    Write-Title
    Write-Host "TeslaPro scan engine start..." -ForegroundColor Cyan
    Write-Host ""
    Spinner "Modules initialiseren" 10

    $steps = @(
        @{P=10; T='Running processes analyseren'; A={ Collect-Processes }},
        @{P=24; T='Registry apps controleren';   A={ Collect-RegistryApps }},
        @{P=40; T='Program folders scannen';     A={ Collect-ProgramFolders }},
        @{P=56; T='Shortcuts en startup zoeken'; A={ Collect-Shortcuts }},
        @{P=76; T='Scripts en configs inspecteren'; A={ Collect-TargetFiles }},
        @{P=88; T='Deep scan uitvoeren';         A={ Collect-DeepFiles }},
        @{P=96; T='Resultaten analyseren';       A={
            $unique = $Script:Items | Sort-Object Category, Vendor, Name, Path -Unique
            foreach ($u in $unique) {
                $Script:Analyzed.Add((Analyze-Item -Item $u)) | Out-Null
            }
        }},
        @{P=100; T='Rapporten exporteren';       A={ Export-Reports }}
    )

    foreach ($step in $steps) {
        Show-Progress -Percent $step.P -Text $step.T
        & $step.A
        Start-Sleep -Milliseconds 120
    }

    Write-Host ""
    Write-Host ""
    Write-Host "Scan voltooid." -ForegroundColor Green
    Start-Sleep -Milliseconds 350
}

function Main-Menu {
    while ($true) {
        Show-Summary
        $key = [System.Console]::ReadKey($true).Key

        switch ($key) {
            'D1' { Show-Summary; Pause-Key }
            'NumPad1' { Show-Summary; Pause-Key }

            'D2' {
                $rows = Sort-Rows ($Script:Analyzed | Where-Object { $_.Category -match 'InstalledApp|ProgramFolder|Process|Shortcut' })
                Show-CardPage -TabName 'Apps' -TitleText 'Apps / Processen / Shortcuts' -Rows $rows
            }
            'NumPad2' {
                $rows = Sort-Rows ($Script:Analyzed | Where-Object { $_.Category -match 'InstalledApp|ProgramFolder|Process|Shortcut' })
                Show-CardPage -TabName 'Apps' -TitleText 'Apps / Processen / Shortcuts' -Rows $rows
            }

            'D3' {
                $rows = Sort-Rows ($Script:Analyzed | Where-Object { $_.Category -match 'Script|PossibleConfig|DeepScanHit' })
                Show-CardPage -TabName 'Scripts' -TitleText 'Scripts / Configs / Deep Scan Hits' -Rows $rows
            }
            'NumPad3' {
                $rows = Sort-Rows ($Script:Analyzed | Where-Object { $_.Category -match 'Script|PossibleConfig|DeepScanHit' })
                Show-CardPage -TabName 'Scripts' -TitleText 'Scripts / Configs / Deep Scan Hits' -Rows $rows
            }

            'D4' {
                $rows = Sort-Rows ($Script:Analyzed | Where-Object { $_.RiskLabel -eq 'HIGH' })
                Show-CardPage -TabName 'High-Risk' -TitleText 'High Risk Items' -Rows $rows
            }
            'NumPad4' {
                $rows = Sort-Rows ($Script:Analyzed | Where-Object { $_.RiskLabel -eq 'HIGH' })
                Show-CardPage -TabName 'High-Risk' -TitleText 'High Risk Items' -Rows $rows
            }

            'D5' {
                $rows = Sort-Rows $Script:Analyzed
                Show-CardPage -TabName 'All' -TitleText 'Alle Resultaten' -Rows $rows
            }
            'NumPad5' {
                $rows = Sort-Rows $Script:Analyzed
                Show-CardPage -TabName 'All' -TitleText 'Alle Resultaten' -Rows $rows
            }

            'D6' { Show-ExportInfo }
            'NumPad6' { Show-ExportInfo }

            'Q' { break }
            'Escape' { break }
        }
    }
}

# ==============================
# Run
# ==============================
try {
    Write-Title
    Run-Scan
    Main-Menu
}
finally {
    Write-Title
    Write-Host "Scan afgerond." -ForegroundColor Green
    Write-Host "TXT  : $Script:ReportTxt" -ForegroundColor Gray
    Write-Host "JSON : $Script:ReportJson" -ForegroundColor Gray
    Write-Host "INV  : $Script:Inventory" -ForegroundColor Gray
    Write-Host ""
}