param([switch]$Deep)

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "TeslaProMacroFinder V3"

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($BaseDir)) { $BaseDir = (Get-Location).Path }

$ReportTxt  = Join-Path $BaseDir "TeslaProMacroFinder_V3_Report.txt"
$ReportJson = Join-Path $BaseDir "TeslaProMacroFinder_V3_Report.json"

$Found    = New-Object System.Collections.Generic.List[object]
$Analyzed = New-Object System.Collections.Generic.List[object]
$Script:DetectedVendors = @()

$VendorRules = @(
    @{
        Vendor = 'Logitech'
        ProcessNames = @('lghub','lcore','logioptionsplus','logioptions')
        PathHints = @('\LGHUB\','\Logitech\','\Logishrd\')
        RegistryHints = @('Logitech','G HUB','Logi')
        Roots = @(
            "$env:LOCALAPPDATA\Logitech\Logitech Gaming Software",
            "$env:LOCALAPPDATA\LGHUB",
            "$env:PROGRAMDATA\LGHUB",
            "$env:PROGRAMDATA\Logishrd"
        )
    },
    @{
        Vendor = 'Razer'
        ProcessNames = @('rzsynapse','razercentral','razerappengine','rzstats','rzchromastreamserver')
        PathHints = @('\Razer\','\Synapse\')
        RegistryHints = @('Razer','Synapse')
        Roots = @(
            "$env:LOCALAPPDATA\Razer",
            "$env:PROGRAMDATA\Razer\Synapse3\Accounts",
            "$env:LOCALAPPDATA\Razer\Synapse3\Log",
            "$env:PROGRAMDATA\Razer"
        )
    },
    @{
        Vendor = 'SteelSeries'
        ProcessNames = @('steelseriesgg','steelseriesengine')
        PathHints = @('\SteelSeries\','\steelseries-engine-3-client\')
        RegistryHints = @('SteelSeries','GG')
        Roots = @(
            "$env:LOCALAPPDATA\steelseries-engine-3-client\Local Storage\leveldb",
            "$env:LOCALAPPDATA\SteelSeries",
            "$env:PROGRAMDATA\SteelSeries"
        )
    },
    @{
        Vendor = 'Roccat'
        ProcessNames = @('roccatswarm','swarm')
        PathHints = @('\ROCCAT\','\Swarm\')
        RegistryHints = @('ROCCAT','Swarm')
        Roots = @(
            "$env:APPDATA\ROCCAT\SWARM"
        )
    },
    @{
        Vendor = 'Corsair'
        ProcessNames = @('icue','corsair.service')
        PathHints = @('\Corsair\','\CUE\','\iCUE\')
        RegistryHints = @('Corsair','iCUE','CUE')
        Roots = @(
            "$env:APPDATA\Corsair\CUE",
            "$env:LOCALAPPDATA\Corsair",
            "$env:PROGRAMDATA\Corsair"
        )
    },
    @{
        Vendor = 'Bloody/A4Tech'
        ProcessNames = @('bloody7','bloody')
        PathHints = @('\Bloody7\','\A4Tech\')
        RegistryHints = @('Bloody','A4Tech')
        Roots = @(
            "$env:PROGRAMFILES(X86)\Bloody7\Bloody7\Data\Mouse\English\ScriptsMacros\GunLib",
            "$env:LOCALAPPDATA\Bloody7",
            "$env:APPDATA\Bloody7"
        )
    },
    @{
        Vendor = 'ASUS'
        ProcessNames = @('armourycrate','asusnodejswebframework','armourycrate.user.session.helper')
        PathHints = @('\ASUS\','\Armoury\')
        RegistryHints = @('ASUS','Armoury')
        Roots = @(
            "$env:USERPROFILE\Documents\ASUS\ROG\ROG Armoury\common"
        )
    },
    @{
        Vendor = 'AutoHotkey'
        ProcessNames = @('autohotkey','autohotkeyu64','autohotkey32')
        PathHints = @('\AutoHotkey\')
        RegistryHints = @('AutoHotkey')
        Roots = @()
    },
    @{
        Vendor = 'X-Mouse'
        ProcessNames = @('xmousebuttoncontrol')
        PathHints = @('\X-Mouse Button Control\','\XMBC\')
        RegistryHints = @('X-Mouse')
        Roots = @()
    },
    @{
        Vendor = 'Glorious/Ajazz-like'
        ProcessNames = @()
        PathHints = @('\BY-COMBO2\','\BYCOMBO-2\')
        RegistryHints = @('BY-COMBO','Glorious')
        Roots = @(
            "$env:APPDATA\BY-COMBO2",
            "$env:APPDATA\BYCOMBO-2"
        )
    },
    @{
        Vendor = 'Redragon/MotoSpeed'
        ProcessNames = @()
        PathHints = @('\MotoSpeed\','\MOTO SPEED\','\Redragon\')
        RegistryHints = @('MotoSpeed','Redragon')
        Roots = @(
            "$env:PROGRAMFILES(X86)\MotoSpeed Gaming Mouse",
            "$env:USERPROFILE\Documents"
        )
    },
    @{
        Vendor = 'Cooler Master'
        ProcessNames = @()
        PathHints = @('\CoolerMaster\')
        RegistryHints = @('Cooler Master','CoolerMaster')
        Roots = @(
            "$env:LOCALAPPDATA\CoolerMaster",
            "$env:APPDATA\CoolerMaster",
            "$env:PROGRAMDATA\CoolerMaster"
        )
    }
)

$TextLikeExtensions = @('.ahk','.lua','.json','.ini','.cfg','.conf','.xml','.txt','.log','.bat','.cmd','.ps1')
$VendorExtensions   = @('.ahk','.lua','.json','.ini','.cfg','.conf','.xml','.txt','.log','.db','.dat','.bat','.cmd','.ps1','.amc2','.bin')
$UserExtensions     = @('.ahk','.lua','.bat','.cmd','.ps1','.txt','.ini','.cfg','.json','.xml')

function Clear-UI { Clear-Host }

function Line([string]$Char = "═", [ConsoleColor]$Color = [ConsoleColor]::DarkCyan) {
    $w = [Math]::Max(100, $Host.UI.RawUI.WindowSize.Width - 1)
    Write-Host ($Char * $w) -ForegroundColor $Color
}

function Header {
    Clear-UI
    Line
    Write-Host "                               TeslaProMacroFinder V3" -ForegroundColor Green
    Write-Host "                      Precision Macro / Vendor Audit Console" -ForegroundColor Cyan
    Line
    Write-Host ""
}

function Tabs([string]$Active) {
    $tabs = @(
        @{K='1';T='Summary'},
        @{K='2';T='Vendor Software'},
        @{K='3';T='Relevant Configs'},
        @{K='4';T='Recent Changes'},
        @{K='5';T='High Risk'},
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
    $width = 40
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

function CutText([string]$Text, [int]$Max = 90) {
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

function IsStrongMacroName([string]$Name) {
    $n = (SafeText $Name).ToLowerInvariant()
    return (
        $n -match '(^|[\W_])(macro|rapidfire|autoclick|autohotkey|ahk|lua|recoil|spamclick|trigger|turbo|bhop|xmouse)([\W_]|$)' -or
        $n -match 'ghub|synapse|swarm|icue|bloody|armoury|macrodb|leftmousebutton|button4'
    )
}

function KeywordHits([string]$Path) {
    $keywords = @(
        'macro','macros','autohotkey','ahk','lua','rapidfire','autoclick','auto click',
        'toggle','ghub','lghub','synapse','razer','logitech','hotkey','button 4','button4',
        'sendinput','click','script','recoil','jitter','bind','loop','repeat','turbo',
        'spamclick','trigger','bhop','autostrafe','autojump','delay','leftmousebutton',
        'profile switch','profile','macro_list','macrodb','custom_macro','recmouseclicksenable'
    )

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "-" }
        $it = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($it.PSIsContainer) { return "-" }
        if ($it.Length -gt 5MB) { return "-" }

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

function FileRelevanceScore([string]$Name, [string]$Path, [string]$Keywords, [string]$Vendor, [string]$Category) {
    $blob = (($Name + ' ' + $Path + ' ' + $Keywords + ' ' + $Vendor + ' ' + $Category).ToLowerInvariant())
    $score = 0

    foreach ($k in @('macro','autoclick','rapidfire','recoil','sendinput','trigger','spamclick','turbo','bhop','autostrafe','autojump','jitter','delay','repeat')) {
        if ($blob -match [regex]::Escape($k)) { $score += 12 }
    }
    foreach ($k in @('autohotkey','ahk','lua','.ahk','.lua')) {
        if ($blob -match [regex]::Escape($k)) { $score += 22 }
    }
    foreach ($k in @('logitech','razer','synapse','ghub','steelseries','roccat','swarm','corsair','bloody','x-mouse','armoury','macrodb')) {
        if ($blob -match [regex]::Escape($k)) { $score += 8 }
    }
    if ($Category -eq 'Process') { $score += 10 }
    if ($Category -eq 'Script') { $score += 15 }
    if ($Category -eq 'VendorConfig') { $score += 8 }
    if ($Category -eq 'InstalledApp') { $score += 6 }

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

function Detect-ActiveVendors {
    $active = New-Object System.Collections.Generic.List[string]

    Get-CimInstance Win32_Process | ForEach-Object {
        $v = ProcessVendor $_.Name $_.ExecutablePath $_.CommandLine
        if ($v -and -not $active.Contains($v)) { $active.Add($v) | Out-Null }
    }

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
                foreach ($hint in $rule.RegistryHints) {
                    if ($blob -like "*$hint*") {
                        if (-not $active.Contains($rule.Vendor)) { $active.Add($rule.Vendor) | Out-Null }
                        break
                    }
                }
            }
        }
    }

    return @($active)
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
                    Add-Hit 'InstalledApp' $rule.Vendor $dn $il $key 'Matched official vendor uninstall entry'
                    break
                }
            }
        }
    }
}

function Collect-VendorFiles([string[]]$ActiveVendors) {
    $patterns = @('*.ahk','*.lua','*.json','*.ini','*.cfg','*.conf','*.xml','*.txt','*.log','*.db','*.dat','*.bat','*.cmd','*.ps1','*.amc2','*.bin')

    foreach ($rule in $VendorRules) {
        if ($rule.Vendor -notin $ActiveVendors) { continue }

        foreach ($root in ($rule.Roots | Where-Object { $_ -and (Test-Path $_) })) {
            foreach ($pat in $patterns) {
                Get-ChildItem -LiteralPath $root -Recurse -File -Filter $pat | ForEach-Object {
                    if ($_.Extension.ToLowerInvariant() -notin $VendorExtensions) { return }

                    $category = if ($_.Extension -match '\.(ahk|lua|bat|cmd|ps1)$') { 'Script' } else { 'VendorConfig' }
                    $reason = if (IsStrongMacroName $_.Name) { 'Vendor-area suspicious file name' } else { "Vendor-area relevant file pattern $pat" }
                    Add-Hit $category $rule.Vendor $_.Name $_.FullName $root $reason
                }
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

    foreach ($root in $roots) {
        Get-ChildItem -LiteralPath $root -Recurse -File | ForEach-Object {
            $ext = $_.Extension.ToLowerInvariant()
            if ($ext -notin $UserExtensions) { return }

            $strongName = IsStrongMacroName $_.Name
            $mustInclude = $false

            if ($ext -in @('.ahk','.lua')) { $mustInclude = $true }
            elseif ($strongName) { $mustInclude = $true }

            if (-not $mustInclude) { return }

            $vendor = DetectVendorFromPath $_.FullName
            $category = if ($ext -match '\.(ahk|lua|bat|cmd|ps1)$') { 'Script' } else { 'UserConfig' }
            $reason = if ($strongName) { 'User-area suspicious macro-related file' } else { 'User-area strong script file' }
            Add-Hit $category $vendor $_.Name $_.FullName $root $reason
        }
    }
}

function Collect-Deep {
    if (-not $Deep) { return }
    $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -and (Test-Path $_.Root) }
    $patterns = @('*.ahk','*.lua','*macro*.txt','*macro*.ini','*macro*.cfg','*rapidfire*.txt','*autoclick*.txt','*macro*.json')

    foreach ($d in $drives) {
        foreach ($pat in $patterns) {
            Get-ChildItem -LiteralPath $d.Root -Recurse -File -Filter $pat | ForEach-Object {
                $ext = $_.Extension.ToLowerInvariant()
                if ($ext -notin $UserExtensions -and $ext -notin $VendorExtensions) { return }
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

    $score = FileRelevanceScore $Item.Name $Item.Path $keywords $Item.Vendor $Item.Category
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
    $lines.Add("TeslaProMacroFinder V3")
    $lines.Add("Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    $lines.Add("")
    $lines.Add("SUMMARY")
    $lines.Add("----------------------------------------------------------------")
    $lines.Add("Detected vendors  : $([string]::Join(', ', $Script:DetectedVendors))")
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
    $lines.Add("2. Open official vendor software and inspect macro/button pages.")
    $lines.Add("3. Use an online mouse button tester and press every button one by one.")
    $lines.Add("4. Repeat the same test with vendor software fully CLOSED and killed.")
    $lines.Add("5. If an extra physical button consistently registers as left click, treat that as a red flag.")
    $lines.Add("6. Prioritize recent changes in vendor config/db/log files for context.")

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
    Write-Host ("Detected vendors : {0}" -f ($(if ($Script:DetectedVendors.Count -gt 0) { [string]::Join(', ', $Script:DetectedVendors) } else { "-" }))) -ForegroundColor White
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
        Write-Host ("{0,-24} {1,4}" -f $v.Name, $v.Count) -ForegroundColor Gray
    }
    Write-Host ""
    Write-Host "V3 only scans official vendor families + strongly relevant user-side files." -ForegroundColor DarkGray
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
        Write-Host ("│ Name     : {0}" -f (CutText $x.Name 90)) -ForegroundColor White
        Write-Host ("│ Vendor   : {0}" -f $x.Vendor) -ForegroundColor Gray
        Write-Host ("│ Risk     : {0} ({1})" -f $x.RiskLabel, $x.RiskScore) -ForegroundColor $riskColor
        Write-Host ("│ Type     : {0}" -f $x.Category) -ForegroundColor Gray
        Write-Host ("│ Running  : {0}" -f $x.Running) -ForegroundColor Gray
        Write-Host ("│ Exists   : {0}" -f $x.Exists) -ForegroundColor Gray
        Write-Host ("│ Created  : {0}" -f $x.Created) -ForegroundColor Gray
        Write-Host ("│ Accessed : {0}" -f $x.Accessed) -ForegroundColor Gray
        Write-Host ("│ Modified : {0}" -f $x.Modified) -ForegroundColor Gray
        Write-Host ("│ Size     : {0}" -f $x.Size) -ForegroundColor Gray
        Write-Host ("│ Keywords : {0}" -f (CutText $x.Keywords 90)) -ForegroundColor Gray
        Write-Host ("│ Path     : {0}" -f (CutText $x.Path 90)) -ForegroundColor DarkGray
        Write-Host ("│ Reason   : {0}" -f (CutText $x.Reason 90)) -ForegroundColor DarkGray
        Write-Host "└──────────────────────────────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host ""
    }
    Pause-Key
}

function Show-RecentPage {
    $rows = SortRows ($Analyzed | Where-Object { $_.Modified -ne '-' -and $_.Category -ne 'InstalledApp' })
    $rows = @($rows | Select-Object -First 25)
    Show-Cards 'Recent Changes' 'RECENTLY MODIFIED RELEVANT ITEMS' $rows
}

function Show-ManualChecks {
    Header
    Tabs 'Manual Checks'
    Write-Host "MANUAL ON-BOARD CHECKLIST" -ForegroundColor Cyan
    Line "-" DarkGray
    Write-Host "1. Confirm exact mouse brand/model." -ForegroundColor White
    Write-Host "2. Open official vendor software and inspect macros/bindings." -ForegroundColor White
    Write-Host "3. Use a proper mouse button test site and press every button one by one." -ForegroundColor White
    Write-Host "4. Fully close vendor software and repeat the exact same test." -ForegroundColor White
    Write-Host "5. Compare whether extra buttons output left click or repeated actions." -ForegroundColor White
    Write-Host "6. Use recent vendor config changes as supporting context, not sole proof." -ForegroundColor White
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
    Write-Host "This scanner is strong on software/config traces; on-board checks remain manual." -ForegroundColor Gray
    Pause-Key
}

function Run-Scan {
    Header
    Spinner "TeslaPro precision scan initialiseren" 10

    $Script:DetectedVendors = Detect-ActiveVendors

    $steps = @(
        @{P=15; T='Exacte vendor-processen analyseren'; A={ Collect-Processes }},
        @{P=35; T='Officiële vendor software controleren'; A={ Collect-RegistryApps }},
        @{P=78; T='Vendor configs en sterke user scripts scannen'; A={ Collect-VendorFiles $Script:DetectedVendors; Collect-UserScripts }},
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
                Show-Cards 'Vendor Software' 'VENDOR SOFTWARE / PROCESSES' $rows
            }
            'NumPad2' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'InstalledApp|Process' })
                Show-Cards 'Vendor Software' 'VENDOR SOFTWARE / PROCESSES' $rows
            }
            'D3' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'Script|VendorConfig|UserConfig|DeepScanHit' })
                Show-Cards 'Relevant Configs' 'RELEVANT CONFIGS / SCRIPTS / FILES' $rows
            }
            'NumPad3' {
                $rows = SortRows ($Analyzed | Where-Object { $_.Category -match 'Script|VendorConfig|UserConfig|DeepScanHit' })
                Show-Cards 'Relevant Configs' 'RELEVANT CONFIGS / SCRIPTS / FILES' $rows
            }
            'D4' { Show-RecentPage }
            'NumPad4' { Show-RecentPage }
            'D5' {
                $rows = SortRows ($Analyzed | Where-Object { $_.RiskLabel -eq 'HIGH' })
                Show-Cards 'High Risk' 'HIGH RISK ITEMS' $rows
            }
            'NumPad5' {
                $rows = SortRows ($Analyzed | Where-Object { $_.RiskLabel -eq 'HIGH' })
                Show-Cards 'High Risk' 'HIGH RISK ITEMS' $rows
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