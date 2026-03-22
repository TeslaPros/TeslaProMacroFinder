param([switch]$Deep)

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "TeslaProMacroFinder V1.1"

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($BaseDir)) { $BaseDir = (Get-Location).Path }

$ReportTxt  = Join-Path $BaseDir "TeslaProMacroFinder_V1_1_Report.txt"
$ReportJson = Join-Path $BaseDir "TeslaProMacroFinder_V1_1_Report.json"

$Found    = New-Object System.Collections.Generic.List[object]
$Analyzed = New-Object System.Collections.Generic.List[object]

$VendorRules = @(
    @{
        Vendor = 'Logitech'
        ProcessNames = @('lghub','logioptionsplus','logioptions','logi_overlay','lcore')
        PathHints = @('\LGHUB\', '\Logitech\', '\Logishrd\')
        RegistryHints = @('Logitech','G HUB','Logi')
    },
    @{
        Vendor = 'Razer'
        ProcessNames = @('razer synapse service','rzsynapse','rzstats','razerappengine','razercentral','rzchromastreamserver')
        PathHints = @('\Razer\', '\Synapse\')
        RegistryHints = @('Razer','Synapse')
    },
    @{
        Vendor = 'SteelSeries'
        ProcessNames = @('steelseriesengine','steelseriesgg','gg')
        PathHints = @('\SteelSeries\', '\steelseries-engine-3-client\')
        RegistryHints = @('SteelSeries','GG')
    },
    @{
        Vendor = 'Roccat'
        ProcessNames = @('roccatswarm','swarm')
        PathHints = @('\ROCCAT\', '\Swarm\')
        RegistryHints = @('ROCCAT','Swarm')
    },
    @{
        Vendor = 'Corsair'
        ProcessNames = @('icue','corsair.service')
        PathHints = @('\Corsair\', '\CUE\', '\iCUE\')
        RegistryHints = @('Corsair','iCUE','CUE')
    },
    @{
        Vendor = 'Bloody/A4Tech'
        ProcessNames = @('bloody7','bloody')
        PathHints = @('\Bloody7\', '\A4Tech\')
        RegistryHints = @('Bloody','A4Tech')
    },
    @{
        Vendor = 'ASUS'
        ProcessNames = @('armourycrate','armourycrate.user.session.helper','asusnodejswebframework')
        PathHints = @('\ASUS\', '\Armoury\')
        RegistryHints = @('ASUS','Armoury')
    },
    @{
        Vendor = 'AutoHotkey'
        ProcessNames = @('autohotkey','autohotkeyu64','autohotkey32')
        PathHints = @('\AutoHotkey\')
        RegistryHints = @('AutoHotkey')
    },
    @{
        Vendor = 'X-Mouse'
        ProcessNames = @('xmousebuttoncontrol')
        PathHints = @('\X-Mouse Button Control\', '\XMBC\')
        RegistryHints = @('X-Mouse')
    }
)

$RelevantExtensions = @('.ahk','.lua','.json','.ini','.cfg','.conf','.xml','.txt','.log','.db','.dat','.bat','.cmd','.ps1')
$TextLikeExtensions = @('.ahk','.lua','.json','.ini','.cfg','.conf','.xml','.txt','.log','.bat','.cmd','.ps1')

function Clear-UI { Clear-Host }

function Line([string]$Char = "═", [ConsoleColor]$Color = [ConsoleColor]::DarkCyan) {
    $w = [Math]::Max(92, $Host.UI.RawUI.WindowSize.Width - 1)
    Write-Host ($Char * $w) -ForegroundColor $Color
}

function Header {
    Clear-UI
    Line
    Write-Host "                           TeslaProMacroFinder V1.1" -ForegroundColor Green
    Write-Host "                    Precision Macro / Peripheral Audit Console" -ForegroundColor Cyan
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
            Start-Sleep -Milliseconds 55
        }
    }
    Write-Host "`r$Text done.   " -ForegroundColor Green
}

function Progress([int]$Percent, [string]$Text) {
    $width = 36
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

function CutText([string]$Text, [int]$Max = 86) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return "-" }
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max - 3) + "..."
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

function DetectVendorFromPath([string]$Path) {
    $p = SafeText $Path
    foreach ($rule in $VendorRules) {
        foreach ($hint in $rule.PathHints) {
            if ($p -like "*$hint*") { return $rule.Vendor }
        }
    }
    return 'Unknown'
}

function ProcessVendor([string]$Name, [string]$Path, [string]$CommandLine) {
    $n = (SafeText $Name).ToLowerInvariant()
    $p = SafeText $Path
    $c = SafeText $CommandLine

    foreach ($rule in $VendorRules) {
        foreach ($proc in $rule.ProcessNames) {
            if ($n -eq $proc -or $n -eq ($proc + '.exe')) {
                return $rule.Vendor
            }
        }
        foreach ($hint in $rule.PathHints) {
            if ($p -like "*$hint*" -or $c -like "*$hint*") {
                return $rule.Vendor
            }
        }
    }
    return $null
}

function IsSuspiciousFileName([string]$Name) {
    $n = (SafeText $Name).ToLowerInvariant()
    return (
        $n -match '(^|[\W_])(macro|rapidfire|autoclick|autohotkey|ahk|lua|recoil|spamclick|trigger|turbo|bhop|xmouse)([\W_]|$)' -or
        $n -match 'ghub|synapse|swarm|icue|bloody|armoury|x-mouse'
    )
}

function IsRelevantLooseFile([System.IO.FileInfo]$File) {
    $ext = $File.Extension.ToLowerInvariant()
    if ($ext -notin $RelevantExtensions) { return $false }

    if ($ext -in @('.ahk','.lua')) { return $true }

    $nameHit = IsSuspiciousFileName $File.Name
    $pathHit = (DetectVendorFromPath $File.FullName) -ne 'Unknown'

    if ($nameHit -or $pathHit) { return $true }

    return $false
}

function KeywordHits([string]$Path) {
    $keywords = @(
        'macro','macros','autohotkey','ahk','lua','rapidfire','autoclick','auto click',
        'toggle','ghub','lghub','synapse','razer','logitech','hotkey',
        'sendinput','click','script','recoil','jitter','bind','loop','repeat','turbo',
        'spamclick','trigger','bhop','autostrafe','autojump','delay','leftmousebutton','button 4','button4'
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "-" }
        $it = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($it.PSIsContainer) { return "-" }
        if ($it.Length -gt 4MB) { return "-" }

        $ext = $it.Extension.ToLowerInvariant()
        if ($ext -notin $TextLikeExtensions) { return "-" }

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
        if ($blob -match [regex]::Escape($k)) { $score += 12 }
    }
    foreach ($k in @('autohotkey','ahk','lua','.ahk','.lua')) {
        if ($blob -match [regex]::Escape($k)) { $score += 20 }
    }
    foreach ($k in @('logitech','razer','synapse','ghub','steelseries','roccat','swarm','corsair','bloody','x-mouse','armoury')) {
        if ($blob -match [regex]::Escape($k)) { $score += 7 }
    }
    if ($Category -eq 'Process') { $score += 10 }
    if ($Category -eq 'Script') { $score += 14 }
    if ($Category -eq 'VendorConfig') { $score += 8 }

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
    Get-CimInstance Win32_Process | ForEach-Object {
        $name = SafeText $_.Name
        $exe  = SafeText $_.ExecutablePath
        $cmd  = SafeText $_.CommandLine
        $vendor = ProcessVendor $name $exe $cmd
        if ($vendor) {
            Add-Hit 'Process' $vendor $name $exe 'Running Processes' 'Exact vendor process or vendor path match'
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
            $blob = "$dn $il"
            foreach ($rule in $VendorRules) {
                $matched = $false
                foreach ($hint in $rule.RegistryHints) {
                    if ($blob -like "*$hint*") { $matched = $true; break }
                }
                if ($matched) {
                    Add-Hit 'InstalledApp' $rule.Vendor $dn $il $key 'Matched vendor uninstall entry'
                    break
                }
            }
        }
    }
}

function Collect-KnownVendorFiles {
    $known = @(
        "$env:LOCALAPPDATA\Logitech\Logitech Gaming Software",
        "$env:LOCALAPPDATA\LGHUB",
        "$env:PROGRAMDATA\LGHUB",
        "$env:PROGRAMDATA\Logishrd",

        "$env:LOCALAPPDATA\Razer",
        "$env:PROGRAMDATA\Razer\Synapse3\Accounts",
        "$env:LOCALAPPDATA\Razer\Synapse3\Log",

        "$env:LOCALAPPDATA\steelseries-engine-3-client\Local Storage\leveldb",
        "$env:APPDATA\ROCCAT\SWARM",
        "$env:APPDATA\Corsair\CUE",
        "$env:PROGRAMFILES(X86)\Bloody7\Bloody7\Data\Mouse\English\ScriptsMacros\GunLib",
        "$env:USERPROFILE\Documents\ASUS\ROG\ROG Armoury\common",
        "$env:APPDATA\Microsoft\Windows\Start Menu\Programs\Startup"
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -Unique

    $patterns = @('*.ahk','*.lua','*.json','*.ini','*.cfg','*.conf','*.xml','*.txt','*.log','*.db','*.dat','*.bat','*.cmd','*.ps1')

    foreach ($root in $known) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $root -Recurse -File -Filter $pat | ForEach-Object {
                if (-not (IsRelevantLooseFile $_)) { return }
                $vendor = DetectVendorFromPath $_.FullName
                $category = if ($_.Extension -match '\.(ahk|lua|bat|cmd|ps1)$') { 'Script' } else { 'VendorConfig' }
                $reason = if (IsSuspiciousFileName $_.Name) { 'Vendor-area file with suspicious name' } else { "Vendor-area relevant file pattern $pat" }
                Add-Hit $category $vendor $_.Name $_.FullName $root $reason
            }
        }
    }
}

function Collect-UserScripts {
    $roots = @(
        "$env:USERPROFILE\Desktop",
        "$env:USERPROFILE\Documents",
        "$env:USERPROFILE\Downloads"
    ) | Where-Object { $_ -and (Test-Path $_) }

    $patterns = @('*.ahk','*.lua','*.bat','*.cmd','*.ps1','*.txt','*.ini','*.cfg','*.json')

    foreach ($root in $roots) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $root -Recurse -File -Filter $pat | ForEach-Object {
                if (-not (IsRelevantLooseFile $_)) { return }
                $vendor = DetectVendorFromPath $_.FullName
                $category = if ($_.Extension -match '\.(ahk|lua|bat|cmd|ps1)$') { 'Script' } else { 'UserConfig' }
                $reason = if (IsSuspiciousFileName $_.Name) { 'User-area file with suspicious name' } else { 'User-area relevant macro-related file' }
                Add-Hit $category $vendor $_.Name $_.FullName $root $reason
            }
        }
    }
}

function Collect-Deep {
    if (-not $Deep) { return }
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -and (Test-Path $_.Root) }
    $patterns = @('*.ahk','*.lua','*macro*.txt','*macro*.ini','*macro*.cfg','*rapidfire*.txt','*autoclick*.txt')
    foreach ($d in $drives) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $d.Root -Recurse -File -Filter $pat | ForEach-Object {
                if (-not (IsRelevantLooseFile $_)) { return }
                Add-Hit 'DeepScanHit' (DetectVendorFromPath $_.FullName) $_.Name $_.FullName $d.Root 'Deep scan suspicious filename match'
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
    $lines.Add("TeslaProMacroFinder V1.1")
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
    $lines.Add("2. Open official software and inspect macro/button pages.")
    $lines.Add("3. Repeat mouse-button test with software fully CLOSED.")
    $lines.Add("4. If a side/top button acts as left click in both states, investigate on-board macros.")
    $lines.Add("5. Prioritize recent changes in vendor config/db/log files.")

    $lines | Set-Content -LiteralPath $ReportTxt -Encoding UTF8
    $sorted | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $ReportJson -Encoding UTF8
}

function Show-SummaryPage {
    Header
    Tabs 'Summary'
    $rows = $Analyzed
    $vendors = $rows | Group-Object Vendor | Sort-Object Count -Descending

    Write-Host "PRECISION SCAN OVERVIEW" -ForegroundColor Cyan
    Line "-" DarkGray
    Write-Host ("Total items      : {0}" -f $rows.Count) -ForegroundColor White
    Write-Host ("Existing         : {0}" -f (($rows | Where-Object Exists -eq 'YES').Count)) -ForegroundColor Green
    Write-Host ("Missing/Deleted  : {0}" -f (($rows | Where-Object Exists -eq 'NO').Count)) -ForegroundColor Yellow
    Write-Host ("Running          : {0}" -f (($rows | Where-Object Running -eq 'YES').Count)) -ForegroundColor Cyan
    Write-Host ("High risk        : {0}" -f (($rows | Where-Object RiskLabel -eq 'HIGH').Count)) -ForegroundColor Red
    Write-Host ("Medium risk      : {0}" -f (($rows | Where-Object RiskLabel -eq 'MEDIUM').Count)) -ForegroundColor Yellow
    Write-Host ("Low risk         : {0}" -f (($rows | Where-Object RiskLabel -eq 'LOW').Count)) -ForegroundColor Green
    Write-Host ""
    Write-Host "DETECTED VENDORS" -ForegroundColor Cyan
    Line "-" DarkGray
    foreach ($v in $vendors | Select-Object -First 12) {
        Write-Host ("{0,-22} {1,4}" -f $v.Name, $v.Count) -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "V1.1 filters out generic apps like Discord/Opera unless they are actual vendor-linked processes." -ForegroundColor DarkGray
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
    Show-Cards 'Recent' 'RECENTLY MODIFIED RELEVANT ITEMS' $rows
}

function Show-ManualChecks {
    Header
    Tabs 'Manual Checks'
    Write-Host "MANUAL ON-BOARD CHECKLIST" -ForegroundColor Cyan
    Line "-" DarkGray
    Write-Host "1. Confirm exact mouse brand/model." -ForegroundColor White
    Write-Host "2. Open official vendor software and inspect macros/bindings." -ForegroundColor White
    Write-Host "3. Fully close vendor software and repeat button test." -ForegroundColor White
    Write-Host "4. Compare whether extra buttons output left click or repeated actions." -ForegroundColor White
    Write-Host "5. Use recent config changes shown by this scanner as supporting context." -ForegroundColor White
    Write-Host ""
    Write-Host "On-board macros are not always visible in plain files; manual testing remains necessary." -ForegroundColor DarkGray
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
    Write-Host "Last Accessed can be unreliable on Windows." -ForegroundColor Gray
    Write-Host "Last Modified is usually the most useful timestamp." -ForegroundColor Gray
    Write-Host "Deep mode is broader and may still include more noise than normal mode." -ForegroundColor Gray
    Pause-Key
}

function Run-Scan {
    Header
    Spinner "TeslaPro precision scan initialiseren" 10

    $steps = @(
        @{P=20; T='Exacte vendor-processen analyseren'; A={ Collect-Processes }},
        @{P=42; T='Officiële vendor software controleren'; A={ Collect-RegistryApps }},
        @{P=80; T='Vendor configs en user scripts scannen'; A={ Collect-KnownVendorFiles; Collect-UserScripts }},
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
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'Script|VendorConfig|UserConfig|DeepScanHit' })
                Show-Cards 'Configs' 'CONFIGS / SCRIPTS / RELEVANT FILES' $rows
            }
            'NumPad3' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'Script|VendorConfig|UserConfig|DeepScanHit' })
                Show-Cards 'Configs' 'CONFIGS / SCRIPTS / RELEVANT FILES' $rows
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