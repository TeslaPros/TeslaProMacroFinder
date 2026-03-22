param([switch]$Deep)

$ErrorActionPreference = 'SilentlyContinue'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$Host.UI.RawUI.WindowTitle = "TeslaPro Macro Finder Ultra V2"

# ===============================
# UI
# ===============================
function ClearUI { Clear-Host }

function Title {
    ClearUI
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host "        TeslaPro Macro Finder Ultra V2" -ForegroundColor Green
    Write-Host "==================================================" -ForegroundColor Cyan
    Write-Host ""
}

function Loading($text) {
    $frames = @("|","/","-","\")
    for ($i=0;$i -lt 10;$i++){
        foreach($f in $frames){
            Write-Host "`r$text $f" -NoNewline -ForegroundColor Yellow
            Start-Sleep -Milliseconds 60
        }
    }
    Write-Host "`r$text Done      " -ForegroundColor Green
}

# ===============================
# DATA
# ===============================
$Items = @()

function Add-Item {
    param($type,$name,$path,$vendor,$reason)

    $Items += [PSCustomObject]@{
        Type = $type
        Name = $name
        Path = $path
        Vendor = $vendor
        Reason = $reason
    }
}

function Get-Vendor($text){
    $t = $text.ToLower()
    if($t -match "logitech"){return "Logitech"}
    if($t -match "razer"){return "Razer"}
    if($t -match "autohotkey|\.ahk"){return "AutoHotkey"}
    if($t -match "corsair"){return "Corsair"}
    return "Unknown"
}

# ===============================
# SCAN
# ===============================
function Scan-Processes {
    Get-Process | ForEach-Object {
        $n = $_.Name
        if($n -match "logitech|razer|autohotkey|macro"){
            Add-Item "Process" $n "-" (Get-Vendor $n) "Running process"
        }
    }
}

function Scan-Files {
    $paths = @(
        "$env:USERPROFILE\Desktop",
        "$env:USERPROFILE\Documents",
        "$env:USERPROFILE\Downloads"
    )

    $patterns = "*.ahk","*.lua","*.txt","*.cfg","*.ini","*.json"

    foreach($p in $paths){
        if(Test-Path $p){
            foreach($pat in $patterns){
                Get-ChildItem $p -Recurse -Filter $pat | ForEach-Object {
                    $name = $_.Name
                    if($name -match "macro|minecraft|autoclick|ahk|lua"){
                        Add-Item "File" $name $_.FullName (Get-Vendor $_.FullName) "Suspicious name"
                    }
                }
            }
        }
    }
}

# ===============================
# ANALYSE
# ===============================
function Analyse {
    $Results = @()

    foreach($i in $Items){
        $exists = "NO"
        $mod = "-"
        $created = "-"
        $size = "-"

        if(Test-Path $i.Path){
            $exists = "YES"
            $f = Get-Item $i.Path
            $mod = $f.LastWriteTime
            $created = $f.CreationTime
            if(!$f.PSIsContainer){ $size = $f.Length }
        }

        $score = 0
        $blob = ($i.Name + " " + $i.Path).ToLower()

        if($blob -match "macro|autoclick|rapidfire"){ $score += 40 }
        if($blob -match "ahk|lua"){ $score += 40 }
        if($blob -match "logitech|razer"){ $score += 10 }

        if($score -gt 100){$score=100}

        $risk = "LOW"
        if($score -ge 70){$risk="HIGH"}
        elseif($score -ge 35){$risk="MEDIUM"}

        $Results += [PSCustomObject]@{
            Name=$i.Name
            Vendor=$i.Vendor
            Type=$i.Type
            Exists=$exists
            Created=$created
            Modified=$mod
            Size=$size
            Risk=$risk
            Score=$score
            Path=$i.Path
        }
    }

    return $Results
}

# ===============================
# DISPLAY
# ===============================
function Show-Results($data){

    ClearUI
    Title

    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "--------------------------------------"

    Write-Host "Total: $($data.Count)"
    Write-Host "High:  $($data | Where {$_.Risk -eq 'HIGH'}).Count"
    Write-Host "Med:   $($data | Where {$_.Risk -eq 'MEDIUM'}).Count"
    Write-Host "Low:   $($data | Where {$_.Risk -eq 'LOW'}).Count"
    Write-Host ""

    foreach($d in $data){

        $color = "Green"
        if($d.Risk -eq "HIGH"){ $color="Red" }
        elseif($d.Risk -eq "MEDIUM"){ $color="Yellow" }

        Write-Host "[$($d.Risk)] $($d.Name)" -ForegroundColor $color
        Write-Host " Vendor : $($d.Vendor)"
        Write-Host " Type   : $($d.Type)"
        Write-Host " Exists : $($d.Exists)"
        Write-Host " Created: $($d.Created)"
        Write-Host " Mod    : $($d.Modified)"
        Write-Host " Path   : $($d.Path)"
        Write-Host "--------------------------------------"
    }

    pause
}

# ===============================
# RUN
# ===============================
Title
Loading "Initializing"

Scan-Processes
Scan-Files

$results = Analyse

Show-Results $results