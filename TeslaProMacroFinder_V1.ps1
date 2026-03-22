param([switch]$Deep)

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "TeslaProMacroFinder V1"

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($BaseDir)) { $BaseDir = (Get-Location).Path }

$ReportTxt  = Join-Path $BaseDir "TeslaProMacroFinder_Report.txt"
$ReportJson = Join-Path $BaseDir "TeslaProMacroFinder_Report.json"

$Found    = New-Object System.Collections.Generic.List[object]
$Analyzed = New-Object System.Collections.Generic.List[object]

function Clear-UI { Clear-Host }

function Line([string]$Char = "═", [ConsoleColor]$Color = [ConsoleColor]::DarkCyan) {
    $w = [Math]::Max(90, $Host.UI.RawUI.WindowSize.Width - 1)
    Write-Host ($Char * $w) -ForegroundColor $Color
}

function Header {
    Clear-UI
    Line
    Write-Host "                      TeslaProMacroFinder V1" -ForegroundColor Green
    Write-Host "               Macro / Mouse / Software Audit Console" -ForegroundColor Cyan
    Line
    Write-Host ""
}

function Tabs([string]$Active) {
    $tabs = @(
        @{K='1';T='Summary'},
        @{K='2';T='Software'},
        @{K='3';T='Configs'},
        @{K='4';T='Recent'},
        @{K='5';T='High-Risk'},
        @{K='6';T='Manual Checks'},
        @{K='7';T='Exports'},
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
    Line "-" DarkGray
}

function Spinner([string]$Text = "Loading", [int]$Loops = 10) {
    $frames = @('|','/','-','\')
    for ($i = 0; $i -lt $Loops; $i++) {
        foreach ($f in $frames) {
            Write-Host "`r$Text $f" -NoNewline -ForegroundColor Yellow
            Start-Sleep -Milliseconds 60
        }
    }
    Write-Host "`r$Text done.   " -ForegroundColor Green
}

function Progress([int]$Percent, [string]$Text) {
    $width = 34
    $filled = [Math]::Floor(($Percent / 100) * $width)
    $bar = ('█' * $filled).PadRight($width, '░')
    Write-Host "`r[$bar] $Percent%  $Text" -NoNewline -ForegroundColor Green
}

function Pause-Key {
    Write-Host ""
    Write-Host "Druk op een toets om terug te gaan..." -ForegroundColor DarkGray
    [void][System.Console]::ReadKey($true)
}

function SafeText([object]$x) {
    if ($null -eq $x) { return "-" }
    $s = [string]$x
    if ([string]::IsNullOrWhiteSpace($s)) { return "-" }
    return $s
}

function CutText([string]$Text, [int]$Max = 88) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return "-" }
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max - 3) + "..."
}

function VendorFrom([string]$Blob) {
    $b = (SafeText $Blob).ToLowerInvariant()
    switch -Regex ($b) {
        'logitech|lghub|logi'        { 'Logitech'; break }
        'razer|synapse'              { 'Razer'; break }
        'steelseries|gg'             { 'SteelSeries'; break }
        'roccat|swarm'               { 'Roccat'; break }
        'corsair|icue|cue'           { 'Corsair'; break }
        'bloody|a4tech|x7'           { 'Bloody/A4Tech'; break }
        'autohotkey|\.ahk'           { 'AutoHotkey'; break }
        'x-mouse|xmouse|xmbc'        { 'X-Mouse'; break }
        'asus|armoury'               { 'ASUS'; break }
        'by-combo|bycombo|glorious'  { 'Glorious/Ajazz-like'; break }
        'redragon|motospeed'         { 'Redragon/MotoSpeed'; break }
        'coolermaster'               { 'Cooler Master'; break }
        'lua|\.lua'                  { 'Lua'; break }
        default                      { 'Unknown' }
    }
}

function Add-Hit(
    [string]$Category,
    [string]$Vendor,
    [string]$Name,
    [string]$Path,
    [string]$Source,
    [string]$Reason
) {
    $Found.Add([PSCustomObject]@{
        Category = SafeText $Category
        Vendor   = SafeText $Vendor
        Name     = SafeText $Name
        Path     = SafeText $Path
        Source   = SafeText $Source
        Reason   = SafeText $Reason
    }) | Out-Null
}

function KeywordHits([string]$Path) {
    $keywords = @(
        'macro','macros','autohotkey','ahk','lua','rapidfire','autoclick','auto click',
        'toggle','minecraft','ghub','lghub','synapse','razer','logitech','hotkey',
        'sendinput','click','script','recoil','jitter','bind','loop','repeat','turbo',
        'spamclick','trigger','bhop','autostrafe','autojump','delay','leftmousebutton','button 4'
    )
    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "-" }
        $it = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($it.PSIsContainer) { return "-" }
        if ($it.Length -gt 6MB) { return "-" }

        $ext = $it.Extension.ToLowerInvariant()
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
    } catch {
        return "-"
    }
}

function RiskScore($Name, $Path, $Keywords, $Vendor, $Category) {
    $blob = (($Name + ' ' + $Path + ' ' + $Keywords + ' ' + $Vendor + ' ' + $Category).ToLowerInvariant())
    $score = 0

    foreach ($k in @('macro','autoclick','rapidfire','recoil','sendinput','trigger','spamclick','turbo','bhop','autostrafe','autojump','jitter','delay','repeat')) {
        if ($blob -match [regex]::Escape($k)) { $score += 10 }
    }
    foreach ($k in @('autohotkey','ahk','lua','.ahk','.lua')) {
        if ($blob -match [regex]::Escape($k)) { $score += 18 }
    }
    foreach ($k in @('logitech','razer','synapse','ghub','steelseries','roccat','swarm','corsair','bloody','x-mouse','xmbc')) {
        if ($blob -match [regex]::Escape($k)) { $score += 6 }
    }
    if ($Category -match 'Process') { $score += 8 }
    if ($Category -match 'Script') { $score += 12 }
    if ($score -gt 100) { $score = 100 }
    return $score
}

function RiskLabel([int]$Score) {
    if ($Score -ge 70) { return 'HIGH' }
    if ($Score -ge 35) { return 'MEDIUM' }
    return 'LOW'
}

function RiskColor([int]$Score) {
    if ($Score -ge 70) { return [ConsoleColor]::Red }
    if ($Score -ge 35) { return [ConsoleColor]::Yellow }
    return [ConsoleColor]::Green
}

function SortRows([object[]]$Rows) {
    @($Rows | Sort-Object @{Expression='RiskScore';Descending=$true}, @{Expression='Modified';Descending=$true}, Vendor, Name)
}

function Collect-Processes {
    $terms = @('lghub','logitech','razer','synapse','steelseries','gg','roccat','swarm','corsair','icue','bloody','xmouse','autohotkey')
    Get-CimInstance Win32_Process | ForEach-Object {
        $name = SafeText $_.Name
        $exe  = SafeText $_.ExecutablePath
        $cmd  = SafeText $_.CommandLine
        $blob = ($name + ' ' + $exe + ' ' + $cmd).ToLowerInvariant()
        if ($terms | Where-Object { $blob.Contains($_) }) {
            Add-Hit 'Process' (VendorFrom $blob) $name $exe 'Running Processes' 'Matched process name or command line'
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
            $dn = SafeText $_.DisplayName
            $il = SafeText $_.InstallLocation
            if ($dn -match 'Logitech|G HUB|Razer|Synapse|SteelSeries|GG|Roccat|Swarm|Corsair|iCUE|Bloody|X-Mouse|AutoHotkey|Macro|Armoury|Cooler Master|Redragon|MotoSpeed') {
                Add-Hit 'InstalledApp' (VendorFrom "$dn $il") $dn $il $key 'Matched uninstall registry entry'
            }
        }
    }
}

function Collect-KnownPaths {
    $known = @(
        "$env:LOCALAPPDATA\Logitech\Logitech Gaming Software",
        "$env:LOCALAPPDATA\LGHUB",
        "$env:LOCALAPPDATA\Razer",
        "$env:PROGRAMDATA\Razer\Synapse3\Accounts",
        "$env:LOCALAPPDATA\Razer\Synapse3\Log",
        "$env:LOCALAPPDATA\steelseries-engine-3-client\Local Storage\leveldb",
        "$env:APPDATA\ROCCAT\SWARM",
        "$env:APPDATA\BY-COMBO2",
        "$env:APPDATA\BYCOMBO-2",
        "$env:APPDATA\Corsair\CUE",
        "$env:PROGRAMFILES(X86)\Bloody7\Bloody7\Data\Mouse\English\ScriptsMacros\GunLib",
        "$env:USERPROFILE\Documents\ASUS\ROG\ROG Armoury\common",
        "$env:LOCALAPPDATA\CoolerMaster",
        "$env:APPDATA\CoolerMaster",
        "$env:PROGRAMDATA\CoolerMaster",
        "$env:USERPROFILE\Documents",
        "$env:USERPROFILE\Desktop",
        "$env:USERPROFILE\Downloads",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    $patterns = @('*.ahk','*.lua','*.json','*.ini','*.cfg','*.conf','*.xml','*.txt','*.log','*.db','*.dat','*.bat','*.cmd','*.ps1')

    foreach ($root in $known) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $root -Recurse -File -Filter $pat | ForEach-Object {
                $vendor = VendorFrom $_.FullName
                $category = if ($_.Extension -match '\.(ahk|lua|bat|cmd|ps1)$') { 'Script' } else { 'Config/File' }
                $reason = if ($_.Name -match 'macro|minecraft|autohotkey|ahk|lua|rapidfire|toggle|recoil|click|ghub|synapse|swarm|cue|icue|macrodb') {
                    'Matched suspicious file name'
                } else {
                    "Matched file pattern $pat"
                }
                Add-Hit $category $vendor $_.Name $_.FullName $root $reason
            }
        }
    }
}

function Collect-Deep {
    if (-not $Deep) { return }
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -and (Test-Path $_.Root) }
    $patterns = @('*.ahk','*.lua','*macro*.txt','*macro*.ini','*macro*.cfg','*rapidfire*.txt','*minecraft*.txt')
    foreach ($d in $drives) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $d.Root -Recurse -File -Filter $pat | ForEach-Object {
                Add-Hit 'DeepScanHit' (VendorFrom $_.FullName) $_.Name $_.FullName $d.Root 'Deep scan pattern match'
            }
        }
    }
}

function Analyze-Item([object]$Item) {
    $exists = 'NO'
    $created = '-'
    $modified = '-'
    $accessed = '-'
    $size = '-'
    $keywords = '-'
    $running = if ($Item.Category -eq 'Process') { 'YES' } else { 'NO' }

    if ($Item.Path -ne '-' -and (Test-Path -LiteralPath $Item.Path)) {
        $exists = 'YES'
        try {
            $it = Get-Item -LiteralPath $Item.Path -ErrorAction Stop
            $created  = $it.CreationTime.ToString('yyyy-MM-dd HH:mm:ss')
            $modified = $it.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
            $accessed = $it.LastAccessTime.ToString('yyyy-MM-dd HH:mm:ss')
            if (-not $it.PSIsContainer) { $size = [string]$it.Length }
            $keywords = KeywordHits $Item.Path
        } catch {}
    }

    $score = RiskScore $Item.Name $Item.Path $keywords $Item.Vendor $Item.Category
    $label = RiskLabel $score

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
    $sorted = SortRows $Analyzed

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("TeslaProMacroFinder V1")
    $lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add("")
    $lines.Add("SUMMARY")
    $lines.Add("----------------------------------------------------------------")
    $lines.Add("Found items       : $($sorted.Count)")
    $lines.Add("Existing items    : $(($sorted | Where-Object Exists -eq 'YES').Count)")
    $lines.Add("Missing items     : $(($sorted | Where-Object Exists -eq 'NO').Count)")
    $lines.Add("Running processes : $(($sorted | Where-Object Running -eq 'YES').Count)")
    $lines.Add("High risk         : $(($sorted | Where-Object RiskLabel -eq 'HIGH').Count)")
    $lines.Add("Medium risk       : $(($sorted | Where-Object RiskLabel -eq 'MEDIUM').Count)")
    $lines.Add("Low risk          : $(($sorted | Where-Object RiskLabel -eq 'LOW').Count)")
    $lines.Add("")
    $lines.Add("DETAILS")
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

    $lines.Add("")
    $lines.Add("MANUAL CHECKS")
    $lines.Add("----------------------------------------------------------------")
    $lines.Add("1. Identify exact mouse brand/model.")
    $lines.Add("2. Open official software and inspect bindings/macros.")
    $lines.Add("3. Repeat button test with software fully CLOSED.")
    $lines.Add("4. Compare physical side/top buttons with actual registered output.")
    $lines.Add("5. Recent modifications near config/database/log files are suspicious.")

    $lines | Set-Content -LiteralPath $ReportTxt -Encoding UTF8
    $sorted | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ReportJson -Encoding UTF8
}

function Show-SummaryPage {
    Header
    Tabs 'Summary'
    $rows = $Analyzed
    $vendors = $rows | Group-Object Vendor | Sort-Object Count -Descending

    Write-Host "SCAN OVERVIEW" -ForegroundColor Cyan
    Line "-" DarkGray
    Write-Host ("Total items      : {0}" -f $rows.Count) -ForegroundColor White
    Write-Host ("Existing         : {0}" -f (($rows | Where-Object Exists -eq 'YES').Count)) -ForegroundColor Green
    Write-Host ("Missing/Deleted  : {0}" -f (($rows | Where-Object Exists -eq 'NO').Count)) -ForegroundColor Yellow
    Write-Host ("Running          : {0}" -f (($rows | Where-Object Running -eq 'YES').Count)) -ForegroundColor Cyan
    Write-Host ("High risk        : {0}" -f (($rows | Where-Object RiskLabel -eq 'HIGH').Count)) -ForegroundColor Red
    Write-Host ("Medium risk      : {0}" -f (($rows | Where-Object RiskLabel -eq 'MEDIUM').Count)) -ForegroundColor Yellow
    Write-Host ("Low risk         : {0}" -f (($rows | Where-Object RiskLabel -eq 'LOW').Count)) -ForegroundColor Green
    Write-Host ""
    Write-Host "TOP VENDORS" -ForegroundColor Cyan
    Line "-" DarkGray
    foreach ($v in $vendors | Select-Object -First 12) {
        Write-Host ("{0,-22} {1,4}" -f $v.Name, $v.Count) -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "Gebruik de tabs bovenaan om te navigeren." -ForegroundColor DarkGray
}

function Show-Cards([string]$ActiveTab, [string]$TitleText, [object[]]$Rows) {
    Header
    Tabs $ActiveTab
    Write-Host $TitleText -ForegroundColor Cyan
    Line "-" DarkGray
    Write-Host ""

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Host "Geen resultaten." -ForegroundColor DarkGray
        Pause-Key
        return
    }

    $i = 0
    foreach ($x in $Rows) {
        $i++
        $riskColor = RiskColor $x.RiskScore
        Write-Host "┌─ Item #$i" -ForegroundColor DarkGray
        Write-Host ("│ Name     : {0}" -f (CutText $x.Name 84)) -ForegroundColor White
        Write-Host ("│ Vendor   : {0}" -f $x.Vendor) -ForegroundColor Gray
        Write-Host ("│ Risk     : {0} ({1})" -f $x.RiskLabel, $x.RiskScore) -ForegroundColor $riskColor
        Write-Host ("│ Type     : {0}" -f $x.Category) -ForegroundColor Gray
        Write-Host ("│ Running  : {0}" -f $x.Running) -ForegroundColor Gray
        Write-Host ("│ Exists   : {0}" -f $x.Exists) -ForegroundColor Gray
        Write-Host ("│ Created  : {0}" -f $x.Created) -ForegroundColor Gray
        Write-Host ("│ Accessed : {0}" -f $x.Accessed) -ForegroundColor Gray
        Write-Host ("│ Modified : {0}" -f $x.Modified) -ForegroundColor Gray
        Write-Host ("│ Size     : {0}" -f $x.Size) -ForegroundColor Gray
        Write-Host ("│ Keywords : {0}" -f (CutText $x.Keywords 84)) -ForegroundColor Gray
        Write-Host ("│ Path     : {0}" -f (CutText $x.Path 84)) -ForegroundColor DarkGray
        Write-Host ("│ Reason   : {0}" -f (CutText $x.Reason 84)) -ForegroundColor DarkGray
        Write-Host "└──────────────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host ""
    }
    Pause-Key
}

function Show-RecentPage {
    $rows = SortRows ($Analyzed | Where-Object { $_.Modified -ne '-' })
    $rows = @($rows | Select-Object -First 25)
    Show-Cards 'Recent' 'RECENTLY MODIFIED ITEMS' $rows
}

function Show-ManualChecks {
    Header
    Tabs 'Manual Checks'
    Write-Host "MANUAL ON-BOARD MACRO CHECKLIST" -ForegroundColor Cyan
    Line "-" DarkGray
    Write-Host "1. Determine the exact mouse brand/model." -ForegroundColor White
    Write-Host "2. Open official software (G HUB / Synapse / Swarm / iCUE / etc.)." -ForegroundColor White
    Write-Host "3. Inspect macro pages and button assignment pages." -ForegroundColor White
    Write-Host "4. Perform a full mouse button test while software is OPEN." -ForegroundColor White
    Write-Host "5. Fully close/kill the vendor software and test again." -ForegroundColor White
    Write-Host "6. If a physical side/top button registers as Left Click repeatedly, that is a red flag." -ForegroundColor Yellow
    Write-Host ""
    Write-Host "This page is intentionally manual because on-board macros can exist without obvious local text configs." -ForegroundColor DarkGray
    Pause-Key
}

function Show-Exports {
    Header
    Tabs 'Exports'
    Write-Host "EXPORT FILES" -ForegroundColor Cyan
    Line "-" DarkGray
    Write-Host "TXT  : $ReportTxt" -ForegroundColor Green
    Write-Host "JSON : $ReportJson" -ForegroundColor Green
    Write-Host ""
    Write-Host "NOTES" -ForegroundColor Yellow
    Line "-" DarkGray
    Write-Host "Last Accessed kan onbetrouwbaar zijn op Windows." -ForegroundColor Gray
    Write-Host "Last Modified is meestal de meest bruikbare tijd." -ForegroundColor Gray
    Write-Host "Gebruik Deep mode alleen als je echt een brede scan wilt." -ForegroundColor Gray
    Pause-Key
}

function Run-Scan {
    Header
    Spinner "TeslaPro scan engine initialiseren" 10

    $steps = @(
        @{P=15; T='Running processes analyseren'; A={ Collect-Processes }},
        @{P=35; T='Registry software controleren'; A={ Collect-RegistryApps }},
        @{P=78; T='Bekende vendor/config paths scannen'; A={ Collect-KnownPaths }},
        @{P=88; T='Deep scan'; A={ Collect-Deep }},
        @{P=96; T='Resultaten analyseren'; A={
            $unique = $Found | Sort-Object Category, Vendor, Name, Path -Unique
            foreach ($u in $unique) { $Analyzed.Add((Analyze-Item $u)) | Out-Null }
        }},
        @{P=100; T='Rapporten exporteren'; A={ Export-Reports }}
    )

    foreach ($s in $steps) {
        Progress $s.P $s.T
        & $s.A
        Start-Sleep -Milliseconds 130
    }

    Write-Host ""
    Write-Host ""
    Write-Host "Scan voltooid." -ForegroundColor Green
    Start-Sleep -Milliseconds 350
}

function Main-Menu {
    while ($true) {
        Show-SummaryPage
        $key = [System.Console]::ReadKey($true).Key
        switch ($key) {
            'D1' { Show-SummaryPage; Pause-Key }
            'NumPad1' { Show-SummaryPage; Pause-Key }
            'D2' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'InstalledApp|Process' })
                Show-Cards 'Software' 'SOFTWARE / PROCESSES' $rows
            }
            'NumPad2' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'InstalledApp|Process' })
                Show-Cards 'Software' 'SOFTWARE / PROCESSES' $rows
            }
            'D3' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'Script|Config/File|DeepScanHit' })
                Show-Cards 'Configs' 'CONFIGS / SCRIPTS / FILES' $rows
            }
            'NumPad3' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'Script|Config/File|DeepScanHit' })
                Show-Cards 'Configs' 'CONFIGS / SCRIPTS / FILES' $rows
            }
            'D4' { Show-RecentPage }
            'NumPad4' { Show-RecentPage }
            'D5' {
                $rows = SortRows ($Analyzed | Where-Object { $_.RiskLabel -eq 'HIGH' })
                Show-Cards 'High-Risk' 'HIGH RISK ITEMS' $rows
            }
            'NumPad5' {
                $rows = SortRows ($Analyzed | Where-Object { $_.RiskLabel -eq 'HIGH' })
                Show-Cards 'High-Risk' 'HIGH RISK ITEMS' $rows
            }
            'D6' { Show-ManualChecks }
            'NumPad6' { Show-ManualChecks }
            'D7' { Show-Exports }
            'NumPad7' { Show-Exports }
            'Q' { break }
            'Escape' { break }
        }
    }
}

try {
    Run-Scan
    Main-Menu
}
finally {
    Header
    Write-Host "Scan afgerond." -ForegroundColor Green
    Write-Host "TXT  : $ReportTxt" -ForegroundColor Gray
    Write-Host "JSON : $ReportJson" -ForegroundColor Gray
    Write-Host ""
}