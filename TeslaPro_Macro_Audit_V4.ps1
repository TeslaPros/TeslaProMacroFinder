param([switch]$Deep)

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "TeslaPro Macro Audit V4"

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($BaseDir)) { $BaseDir = (Get-Location).Path }

$ReportTxt  = Join-Path $BaseDir "TeslaPro_Audit_Report.txt"
$ReportJson = Join-Path $BaseDir "TeslaPro_Audit_Report.json"

$Found = New-Object System.Collections.Generic.List[object]
$Analyzed = New-Object System.Collections.Generic.List[object]

function Clear-UI { Clear-Host }

function Bar {
    param([string]$Char = "═", [ConsoleColor]$Color = [ConsoleColor]::DarkCyan)
    $w = [Math]::Max(78, $Host.UI.RawUI.WindowSize.Width - 1)
    Write-Host ($Char * $w) -ForegroundColor $Color
}

function Title {
    Clear-UI
    Bar
    Write-Host "                    TeslaPro Macro Audit V4" -ForegroundColor Green
    Write-Host "              Local Macro / Mouse / Config Audit" -ForegroundColor Cyan
    Bar
    Write-Host ""
}

function Spinner {
    param([string]$Text = "Loading", [int]$Loops = 12)
    $frames = @('|','/','-','\')
    for ($i = 0; $i -lt $Loops; $i++) {
        foreach ($f in $frames) {
            Write-Host "`r$Text $f" -NoNewline -ForegroundColor Yellow
            Start-Sleep -Milliseconds 55
        }
    }
    Write-Host "`r$Text done.   " -ForegroundColor Green
}

function Progress {
    param([int]$Percent, [string]$Text)
    $w = 32
    $filled = [Math]::Floor(($Percent / 100) * $w)
    $bar = ('█' * $filled).PadRight($w, '░')
    Write-Host "`r[$bar] $Percent%  $Text" -NoNewline -ForegroundColor Green
}

function Pause-Key {
    Write-Host ""
    Write-Host "Druk op een toets om verder te gaan..." -ForegroundColor DarkGray
    [void][System.Console]::ReadKey($true)
}

function SafeText([object]$x) {
    if ($null -eq $x) { return "-" }
    $s = [string]$x
    if ([string]::IsNullOrWhiteSpace($s)) { return "-" }
    return $s
}

function CutText([string]$Text, [int]$Max = 78) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return "-" }
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max - 3) + "..."
}

function Add-Hit {
    param(
        [string]$Category,
        [string]$Vendor,
        [string]$Name,
        [string]$Path,
        [string]$Source,
        [string]$Reason
    )
    $Found.Add([PSCustomObject]@{
        Category = SafeText $Category
        Vendor   = SafeText $Vendor
        Name     = SafeText $Name
        Path     = SafeText $Path
        Source   = SafeText $Source
        Reason   = SafeText $Reason
    }) | Out-Null
}

function GetVendor([string]$Blob) {
    $b = (SafeText $Blob).ToLowerInvariant()
    switch -Regex ($b) {
        'logitech|lghub|logi'        { return 'Logitech' }
        'razer|synapse'              { return 'Razer' }
        'steelseries|gg'             { return 'SteelSeries' }
        'roccat|swarm'               { return 'Roccat' }
        'corsair|icue|cue'           { return 'Corsair' }
        'bloody|a4tech|x7'           { return 'Bloody/A4Tech' }
        'autohotkey|\.ahk'           { return 'AutoHotkey' }
        'x-mouse|xmouse|xmbc'        { return 'X-Mouse' }
        'asus|armoury'               { return 'ASUS' }
        'glorious|by-combo|bycombo'  { return 'Glorious/Ajazz-like' }
        'redragon|m\d{3}'            { return 'Redragon' }
        'lua|\.lua'                  { return 'Lua' }
        default                      { return 'Unknown' }
    }
}

function GetKeywordHits([string]$Path) {
    $keywords = @(
        'macro','macros','autohotkey','ahk','lua','rapidfire','autoclick','auto click',
        'toggle','minecraft','ghub','lghub','synapse','razer','logitech','hotkey',
        'sendinput','click','script','recoil','jitter','bind','loop','repeat','turbo',
        'spamclick','trigger','bhop','autostrafe','autojump','delay','repeat'
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "-" }
        $it = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($it.PSIsContainer) { return "-" }
        if ($it.Length -gt 5MB) { return "-" }

        $ext = $it.Extension.ToLowerInvariant()
        if ($ext -notin '.ahk','.lua','.txt','.cfg','.conf','.ini','.json','.xml','.log','.bat','.cmd','.ps1') {
            return "-"
        }

        $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $hits = foreach ($k in $keywords) {
            if ($content -match [regex]::Escape($k)) { $k }
        }
        $u = $hits | Sort-Object -Unique
        if (-not $u) { return "-" }
        return ($u -join ', ')
    } catch {
        return "-"
    }
}

function GetRiskScore($Name, $Path, $Keywords, $Vendor, $Category) {
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

function GetRiskLabel([int]$Score) {
    if ($Score -ge 70) { return 'HIGH' }
    if ($Score -ge 35) { return 'MEDIUM' }
    return 'LOW'
}

function GetRiskColor([int]$Score) {
    if ($Score -ge 70) { return [ConsoleColor]::Red }
    if ($Score -ge 35) { return [ConsoleColor]::Yellow }
    return [ConsoleColor]::Green
}

function SortRows([object[]]$Rows) {
    @($Rows | Sort-Object @{Expression='RiskScore';Descending=$true}, Vendor, Name)
}

function Collect-Processes {
    $terms = @('lghub','logitech','razer','synapse','steelseries','gg','roccat','swarm','corsair','icue','bloody','xmouse','autohotkey')
    Get-CimInstance Win32_Process | ForEach-Object {
        $name = SafeText $_.Name
        $exe  = SafeText $_.ExecutablePath
        $cmd  = SafeText $_.CommandLine
        $blob = ($name + ' ' + $exe + ' ' + $cmd).ToLowerInvariant()
        if ($terms | Where-Object { $blob.Contains($_) }) {
            Add-Hit 'Process' (GetVendor $blob) $name $exe 'Running Processes' 'Matched process name or command line'
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
            if ($dn -match 'Logitech|G HUB|Razer|Synapse|SteelSeries|GG|Roccat|Swarm|Corsair|iCUE|Bloody|X-Mouse|AutoHotkey|Macro') {
                Add-Hit 'InstalledApp' (GetVendor "$dn $il") $dn $il $key 'Matched uninstall registry entry'
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
        "$env:LOCALAPPDATA\steelseries-engine-3-client\Local Storage\leveldb",
        "$env:APPDATA\ROCCAT\SWARM",
        "$env:APPDATA\BY-COMBO2",
        "$env:APPDATA\BYCOMBO-2",
        "$env:APPDATA\Corsair\CUE",
        "$env:USERPROFILE\Documents\ASUS\ROG\ROG Armoury\common",
        "$env:USERPROFILE\Documents",
        "$env:USERPROFILE\Desktop",
        "$env:USERPROFILE\Downloads",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    $patterns = @('*.ahk','*.lua','*.json','*.ini','*.cfg','*.conf','*.xml','*.txt','*.log','*.db','*.dat','*.bat','*.cmd','*.ps1')

    foreach ($root in $known) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $root -Recurse -File -Filter $pat | ForEach-Object {
                $vendor = GetVendor $_.FullName
                $category = if ($_.Extension -match '\.(ahk|lua|bat|cmd|ps1)$') { 'Script' } else { 'Config/File' }
                $reason = if ($_.Name -match 'macro|minecraft|autohotkey|ahk|lua|rapidfire|toggle|recoil|click|ghub|synapse|swarm|cue|icue') {
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
                Add-Hit 'DeepScanHit' (GetVendor $_.FullName) $_.Name $_.FullName $d.Root 'Deep scan pattern match'
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
            $keywords = GetKeywordHits $Item.Path
        } catch {}
    }

    $score = GetRiskScore $Item.Name $Item.Path $keywords $Item.Vendor $Item.Category
    $label = GetRiskLabel $score

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
    $lines.Add("TeslaPro Macro Audit V4")
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
    $lines.Add("1. Identify exact mouse model.")
    $lines.Add("2. Open official software and check button bindings/macros.")
    $lines.Add("3. Repeat button test with software fully closed.")
    $lines.Add("4. Compare if a physical extra button registers as left click.")
    $lines.Add("5. Review recently modified vendor config/database files.")

    $lines | Set-Content -LiteralPath $ReportTxt -Encoding UTF8
    $sorted | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ReportJson -Encoding UTF8
}

function Show-SummaryPage {
    Title
    $rows = $Analyzed
    $vendors = $rows | Group-Object Vendor | Sort-Object Count -Descending

    Write-Host "SUMMARY" -ForegroundColor Cyan
    Bar "-"
    Write-Host "Total items      : $($rows.Count)" -ForegroundColor White
    Write-Host "Existing         : $(($rows | Where-Object Exists -eq 'YES').Count)" -ForegroundColor Green
    Write-Host "Missing/Deleted  : $(($rows | Where-Object Exists -eq 'NO').Count)" -ForegroundColor Yellow
    Write-Host "Running          : $(($rows | Where-Object Running -eq 'YES').Count)" -ForegroundColor Cyan
    Write-Host "High risk        : $(($rows | Where-Object RiskLabel -eq 'HIGH').Count)" -ForegroundColor Red
    Write-Host "Medium risk      : $(($rows | Where-Object RiskLabel -eq 'MEDIUM').Count)" -ForegroundColor Yellow
    Write-Host "Low risk         : $(($rows | Where-Object RiskLabel -eq 'LOW').Count)" -ForegroundColor Green
    Write-Host ""
    Write-Host "VENDORS" -ForegroundColor Cyan
    Bar "-"
    foreach ($v in $vendors | Select-Object -First 12) {
        Write-Host ("{0,-20} {1,4}" -f $v.Name, $v.Count) -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "[1] Summary  [2] Apps/Processes  [3] Scripts/Configs  [4] High Risk  [5] All  [6] Exports  [Q] Quit" -ForegroundColor DarkGray
}

function Show-Cards($TitleText, [object[]]$Rows) {
    Title
    Write-Host $TitleText -ForegroundColor Cyan
    Bar "-"
    Write-Host ""

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Host "Geen resultaten." -ForegroundColor DarkGray
        Pause-Key
        return
    }

    $i = 0
    foreach ($x in $Rows) {
        $i++
        $riskColor = GetRiskColor $x.RiskScore
        Write-Host "┌─ Item #$i" -ForegroundColor DarkGray
        Write-Host ("│ Name     : {0}" -f (CutText $x.Name 78)) -ForegroundColor White
        Write-Host ("│ Vendor   : {0}" -f $x.Vendor) -ForegroundColor Gray
        Write-Host ("│ Risk     : {0} ({1})" -f $x.RiskLabel, $x.RiskScore) -ForegroundColor $riskColor
        Write-Host ("│ Type     : {0}" -f $x.Category) -ForegroundColor Gray
        Write-Host ("│ Running  : {0}" -f $x.Running) -ForegroundColor Gray
        Write-Host ("│ Exists   : {0}" -f $x.Exists) -ForegroundColor Gray
        Write-Host ("│ Created  : {0}" -f $x.Created) -ForegroundColor Gray
        Write-Host ("│ Accessed : {0}" -f $x.Accessed) -ForegroundColor Gray
        Write-Host ("│ Modified : {0}" -f $x.Modified) -ForegroundColor Gray
        Write-Host ("│ Size     : {0}" -f $x.Size) -ForegroundColor Gray
        Write-Host ("│ Keywords : {0}" -f (CutText $x.Keywords 78)) -ForegroundColor Gray
        Write-Host ("│ Path     : {0}" -f (CutText $x.Path 78)) -ForegroundColor DarkGray
        Write-Host ("│ Reason   : {0}" -f (CutText $x.Reason 78)) -ForegroundColor DarkGray
        Write-Host "└──────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host ""
    }

    Pause-Key
}

function Show-Exports {
    Title
    Write-Host "EXPORTS" -ForegroundColor Cyan
    Bar "-"
    Write-Host "TXT  : $ReportTxt" -ForegroundColor Green
    Write-Host "JSON : $ReportJson" -ForegroundColor Green
    Write-Host ""
    Write-Host "NOTES" -ForegroundColor Yellow
    Bar "-"
    Write-Host "Last Accessed kan onbetrouwbaar zijn op Windows." -ForegroundColor Gray
    Write-Host "Last Modified is meestal de meest bruikbare tijd." -ForegroundColor Gray
    Write-Host "On-board macros moet je handmatig testen via button-test met software open/gesloten." -ForegroundColor Gray
    Pause-Key
}

function Run-Scan {
    Title
    Spinner "TeslaPro modules initialiseren" 10

    $steps = @(
        @{P=15; T='Running processes analyseren'; A={ Collect-Processes }},
        @{P=35; T='Registry apps controleren';   A={ Collect-RegistryApps }},
        @{P=75; T='Bekende vendor paths scannen'; A={ Collect-KnownPaths }},
        @{P=88; T='Deep scan';                   A={ Collect-Deep }},
        @{P=96; T='Resultaten analyseren';       A={
            $unique = $Found | Sort-Object Category, Vendor, Name, Path -Unique
            foreach ($u in $unique) { $Analyzed.Add((Analyze-Item $u)) | Out-Null }
        }},
        @{P=100; T='Rapporten exporteren';       A={ Export-Reports }}
    )

    foreach ($s in $steps) {
        Progress $s.P $s.T
        & $s.A
        Start-Sleep -Milliseconds 130
    }

    Write-Host ""
    Write-Host ""
    Write-Host "Scan voltooid." -ForegroundColor Green
    Start-Sleep -Milliseconds 300
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
                Show-Cards "APPS / PROCESSES" $rows
            }
            'NumPad2' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'InstalledApp|Process' })
                Show-Cards "APPS / PROCESSES" $rows
            }
            'D3' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'Script|Config/File|DeepScanHit' })
                Show-Cards "SCRIPTS / CONFIGS" $rows
            }
            'NumPad3' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'Script|Config/File|DeepScanHit' })
                Show-Cards "SCRIPTS / CONFIGS" $rows
            }
            'D4' {
                $rows = SortRows ($Analyzed | Where-Object { $_.RiskLabel -eq 'HIGH' })
                Show-Cards "HIGH RISK ITEMS" $rows
            }
            'NumPad4' {
                $rows = SortRows ($Analyzed | Where-Object { $_.RiskLabel -eq 'HIGH' })
                Show-Cards "HIGH RISK ITEMS" $rows
            }
            'D5' {
                $rows = SortRows $Analyzed
                Show-Cards "ALL RESULTS" $rows
            }
            'NumPad5' {
                $rows = SortRows $Analyzed
                Show-Cards "ALL RESULTS" $rows
            }
            'D6' { Show-Exports }
            'NumPad6' { Show-Exports }
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
    Title
    Write-Host "Scan afgerond." -ForegroundColor Green
    Write-Host "TXT  : $ReportTxt" -ForegroundColor Gray
    Write-Host "JSON : $ReportJson" -ForegroundColor Gray
    Write-Host ""
}	