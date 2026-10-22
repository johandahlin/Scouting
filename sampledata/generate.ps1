# Build sample scouting rows from real 2026ilpe results and write them into the TestData workbook.
param([string]$Workbook = "D:\JD Projects\DERT\QRCodeTest\DERT_Scouting_REBUILTTestData.xlsx")
$ErrorActionPreference = "Stop"
$dir = $PSScriptRoot
$d = Get-Content "$dir\matches.json" -Raw | ConvertFrom-Json
$rng = New-Object System.Random 2040
function Rnd() { $rng.NextDouble() }
function Pick($arr) { $arr[$rng.Next($arr.Count)] }
function Gauss() { $u1 = 1.0 - (Rnd); $u2 = Rnd; [Math]::Sqrt(-2 * [Math]::Log($u1)) * [Math]::Sin(2 * [Math]::PI * $u2) }
$EN = [char]0x2013   # en dash used by the web form options

# ---------- OPR (least squares) for teleop fuel and auto fuel ----------
$teams = $d | ForEach-Object { $_.red + $_.blue } | Sort-Object -Unique
$idx = @{}; for ($i = 0; $i -lt $teams.Count; $i++) { $idx[[string]$teams[$i]] = $i }
$T = $teams.Count
function Opr($label) {
  $A = New-Object 'double[,]' $T, ($T + 1)
  foreach ($m in $d) {
    for ($s = 0; $s -lt 2; $s++) {
      $al = if ($s -eq 0) { $m.red } else { $m.blue }
      $v = [double](($m.breakdown.$label)[$s])
      foreach ($ta in $al) { foreach ($tb in $al) { $A[$idx["$ta"], $idx["$tb"]] += 1 }; $A[$idx["$ta"], $T] += $v }
    }
  }
  for ($c = 0; $c -lt $T; $c++) {             # Gaussian elimination
    $p = $c; for ($r = $c + 1; $r -lt $T; $r++) { $vr = $A[$r, $c]; $vp = $A[$p, $c]; if ([Math]::Abs($vr) -gt [Math]::Abs($vp)) { $p = $r } }
    for ($k = 0; $k -le $T; $k++) { $tmp = $A[$c, $k]; $A[$c, $k] = $A[$p, $k]; $A[$p, $k] = $tmp }
    for ($r = 0; $r -lt $T; $r++) {
      if ($r -eq $c) { continue }
      $f = $A[$r, $c] / $A[$c, $c]
      if ($f -ne 0) { for ($k = $c; $k -le $T; $k++) { $A[$r, $k] -= $f * $A[$c, $k] } }
    }
  }
  $x = @{}; for ($i = 0; $i -lt $T; $i++) { $x["$($teams[$i])"] = $A[$i, $T] / $A[$i, $i] }
  $x
}
$oprTele = Opr "Teleop Fuel Points"
$oprAuto = Opr "Auto Fuel"
$ranked = $teams | Sort-Object { - $oprTele["$_"] }
$q = @{}; for ($i = 0; $i -lt $ranked.Count; $i++) { $q["$($ranked[$i])"] = 1 - $i / ($ranked.Count - 1) }   # 1 = best

# ---------- team profiles (scout-observed traits, consistent per team) ----------
$climbers = @{}
foreach ($m in $d) { foreach ($p in $m.endgame.PSObject.Properties) { if ($p.Value -ne "None") { $climbers[$p.Name] = $true } } }
$prof = @{}
foreach ($t in $teams) {
  $k = "$t"; $s = $q[$k]
  $prof[$k] = [pscustomobject]@{
    acc       = 0.55 + 0.35 * $s + 0.06 * (Gauss)
    start     = Pick @("Left", "Center", "Right")
    intake    = if ($s -gt 0.5) { @("Neutral Zone floor", "Alliance Zone floor", "Depot") } elseif ($s -gt 0.2) { @("Neutral Zone floor", "Outpost / Human Player") } else { @("Outpost / Human Player", "Depot") }
    shoot     = if ($s -gt 0.7) { @("Close to Hub", "Mid range", "Long range", "While moving") } elseif ($s -gt 0.35) { @("Close to Hub", "Mid range") } else { @("Close to Hub") }
    terrain   = if ($s -gt 0.6) { @("Bump", "Trench") } elseif ($s -gt 0.3) { @(Pick @("Bump", "Trench")) } else { @() }
    inactive  = if ($s -gt 0.6) { @("Collects / stockpiles", "Passes fuel to our zone") } elseif ($s -gt 0.25) { @("Collects / stockpiles") } else { @(Pick @("Plays defense", "Shoots anyway", "Idle")) }
    capacity  = if ($s -gt 0.75) { "40+" } elseif ($s -gt 0.45) { "26${EN}40" } elseif ($s -gt 0.2) { "11${EN}25" } else { "1${EN}10" }
    defender  = ($s -lt 0.3) -and ((Rnd) -lt 0.6)
    preload   = if ((Rnd) -lt 0.8) { 8 } else { 5 + $rng.Next(3) }
    autoRoam  = $oprAuto[$k] -gt 9
  }
}

function Rating($s) { [int][Math]::Max(1, [Math]::Min(5, [Math]::Round(1 + 4 * $s + 0.7 * (Gauss)))) }
function Split([int]$total, $weights) {         # integer split that keeps the exact alliance total
  $sum = ($weights | Measure-Object -Sum).Sum
  $raw = $weights | ForEach-Object { $total * $_ / $sum }
  $res = $raw | ForEach-Object { [int][Math]::Floor($_) }
  $left = $total - ($res | Measure-Object -Sum).Sum
  $order = 0..2 | Sort-Object { - ($raw[$_] - [Math]::Floor($raw[$_])) }
  for ($i = 0; $i -lt $left; $i++) { $res[$order[$i]]++ }
  $res
}

$scouts1 = @("Ava", "Ben", "Chloe", "Dev", "Ella", "Finn")
$scouts2 = @("Grace", "Henry", "Isla", "Jack", "Kai", "Lily")
$rows = New-Object System.Collections.ArrayList
foreach ($m in ($d | Sort-Object num)) {
  # Quals 1-40 on Friday, 41-62 on Saturday, ~8 minute cycle
  if ($m.num -le 40) { $time = [datetime]"2026-03-20 09:00" + [timespan]::FromMinutes(8 * ($m.num - 1) + ($(if ($m.num -gt 22) { 60 } else { 0 }))) }
  else { $time = [datetime]"2026-03-21 09:00" + [timespan]::FromMinutes(8 * ($m.num - 41)) }
  for ($s = 0; $s -lt 2; $s++) {
    $al = if ($s -eq 0) { $m.red } else { $m.blue }
    $alName = if ($s -eq 0) { "Red" } else { "Blue" }
    $died = 0..2 | ForEach-Object { (Rnd) -lt 0.05 }
    $wT = 0..2 | ForEach-Object { [Math]::Max(2, $oprTele["$($al[$_])"]) * [Math]::Exp(0.35 * (Gauss)) * $(if ($died[$_]) { 0.15 } else { 1 }) }
    $wA = 0..2 | ForEach-Object { [Math]::Max(0.5, $oprAuto["$($al[$_])"]) * [Math]::Exp(0.3 * (Gauss)) }
    $tele = Split ([int]$m.breakdown.'Teleop Fuel Points'[$s]) $wT
    $auto = Split ([int]$m.breakdown.'Auto Fuel'[$s]) $wA
    $fouls = ($m.breakdown.'Fouls / Major Fouls'[$s] -split '/') | ForEach-Object { [int]$_.Trim() }
    $minor = @(0, 0, 0); $major = @(0, 0, 0)
    for ($i = 0; $i -lt $fouls[0]; $i++) { $minor[$rng.Next(3)]++ }
    for ($i = 0; $i -lt $fouls[1]; $i++) { $major[$rng.Next(3)]++ }

    for ($i = 0; $i -lt 3; $i++) {
      $t = "$($al[$i])"; $p = $prof[$t]; $sq = $q[$t]
      $acc = [Math]::Min(0.97, [Math]::Max(0.35, $p.acc + 0.05 * (Gauss)))
      $aMiss = [int][Math]::Round($auto[$i] * (1 - $acc) / $acc)
      $tMiss = [int][Math]::Round($tele[$i] * (1 - $acc) / $acc)
      $path = @(); if ($auto[$i] -gt 0 -or (Rnd) -lt 0.4) { $path += "Left start line" }
      if ($auto[$i] -gt $p.preload -or ($p.autoRoam -and (Rnd) -lt 0.5)) { $path += "Went to Neutral Zone"; $path += Pick @("Crossed Bump", "Used Trench") }
      $autoIntake = @(); if ($auto[$i] -gt $p.preload) { $autoIntake += Pick @("Depot", "Neutral Zone") }
      $eg = $m.endgame.$t
      $climb = if ($eg -and $eg -ne "None") { "Level 1" } elseif ($climbers[$t] -and (Rnd) -lt 0.35) { "Attempted" } else { "None" }
      $climbTime = if ($climb -eq "Level 1") { Pick @("5${EN}10 s", "10${EN}20 s") } elseif ($climb -eq "Attempted") { "Over 20 s" } else { "" }
      $defense = if ($p.defender) { Pick @("Some", "A lot") } elseif ((Rnd) -lt 0.08) { "Some" } else { "None" }
      $defended = if ($sq -gt 0.6 -and (Rnd) -lt 0.45) { Pick @("Some", "Heavily") } else { "No" }
      $issues = @(); if ($died[$i]) { $issues += Pick @("Died / disabled", "Lost comms") } elseif ((Rnd) -lt 0.07) { $issues += Pick @("Fuel jam", "Stuck on Bump", "Broke mechanism") }
      $pick = if ($sq -gt 0.7) { "Yes" } elseif ($sq -gt 0.35) { Pick @("Yes", "Maybe", "Maybe", "No") } else { Pick @("Maybe", "No", "No") }
      $notes = @()
      if ($tele[$i] -ge 90) { $notes += "High-volume scorer" } elseif ($tele[$i] -ge 50) { $notes += "Consistent cycles" } elseif ($tele[$i] -lt 10 -and -not $died[$i]) { $notes += "Struggled to score" }
      if ($p.shoot -contains "Long range" -and (Rnd) -lt 0.5) { $notes += "shoots well from long range" }
      if ($climb -eq "Level 1") { $notes += "climbed L1 cleanly" } elseif ($climb -eq "Attempted") { $notes += "climb attempt failed" }
      if ($died[$i]) { $notes += "robot stopped moving mid-match" }
      if ($defense -eq "A lot") { $notes += "aggressive defender" }
      if ($p.inactive -contains "Passes fuel to our zone" -and (Rnd) -lt 0.4) { $notes += "fed fuel to partners during inactive shift" }
      $note = ($notes -join "; ")
      if ($note) { $note = $note.Substring(0, 1).ToUpper() + $note.Substring(1) }
      $autoPts = $auto[$i]; $climbPts = if ($climb -eq "Level 1") { 10 } else { 0 }
      $scout = if ($m.num -le 31) { $scouts1[3 * $s + $i] } else { $scouts2[3 * $s + $i] }
      $stamp = ($time + [timespan]::FromSeconds(200 + $rng.Next(120))).ToString("M/d/yyyy, h:mm:ss tt", [Globalization.CultureInfo]::InvariantCulture)

      [void]$rows.Add(@(
        "Qual", $m.num, [int]$t, $alName, ($i + 1), $scout, "Yes", $p.start, $p.preload,
        $auto[$i], $aMiss, ($path -join "; "), ($autoIntake -join "; "), "None",
        $tele[$i], $tMiss, ($p.intake -join "; "), ($p.shoot -join "; "), ($p.terrain -join "; "), ($p.inactive -join "; "),
        $p.capacity, $defense, $defended, $climb, $climbTime,
        (Rating $sq), (Rating $sq), (Rating (($acc - 0.35) / 0.6)), ($issues -join "; "), $minor[$i], $major[$i], "None", $pick, $note,
        $autoPts, $tele[$i], $climbPts, ($autoPts + $tele[$i] + $climbPts), $stamp
      ))
    }
  }
}
"Generated $($rows.Count) rows"

# Sanity: per-alliance fuel adds up to the real totals
$chk = 0; foreach ($m in $d) { for ($s = 0; $s -lt 2; $s++) {
  $al = if ($s -eq 0) { "Red" } else { "Blue" }
  $sum = ($rows | Where-Object { $_[1] -eq $m.num -and $_[3] -eq $al } | ForEach-Object { $_[9] + $_[14] } | Measure-Object -Sum).Sum
  if ($sum -ne ([int]$m.breakdown.'Auto Fuel'[$s] + [int]$m.breakdown.'Teleop Fuel Points'[$s])) { $chk++ } } }
"Alliance fuel mismatches: $chk"
"Top teleop fuel OPR: " + (($ranked | Select-Object -First 5 | ForEach-Object { "$_=$([Math]::Round($oprTele["$_"],1))" }) -join ', ')
"2040 OPR teleop $([Math]::Round($oprTele['2040'],1)) auto $([Math]::Round($oprAuto['2040'],1))"

# ---------- write into the workbook ----------
# Same stamp index.html puts in column A: "S<columns after it>-<FNV-1a hash of their headers>"
function Fnv1a([string]$s) {
  [uint64]$h = 2166136261
  foreach ($b in [Text.Encoding]::UTF8.GetBytes($s)) { $h = (($h -bxor $b) * [uint64]16777619) -band [uint64]4294967295 }
  '{0:x8}' -f $h
}

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Open($Workbook)
  $ws = $wb.Worksheets.Item("Scans"); $lo = $ws.ListObjects.Item("Scans")
  $hdr = @($lo.HeaderRowRange.Value2 | ForEach-Object { "$_" })
  $lead = @(); if ($hdr[0] -eq "Schema") { $lead = @("S$($hdr.Count - 1)-$(Fnv1a ($hdr[1..($hdr.Count - 1)] -join "`t"))") }
  $nCols = $lead.Count + $rows[0].Count
  $arr = New-Object 'object[,]' $rows.Count, $nCols
  for ($r = 0; $r -lt $rows.Count; $r++) { $row = $lead + $rows[$r]; for ($c = 0; $c -lt $nCols; $c++) {
    $v = $row[$c]; if ($v -is [string] -and $v -eq "") { $v = $null }
    $arr[$r, $c] = $v } }
  if ($lo.DataBodyRange -ne $null) { $lo.DataBodyRange.ClearContents() | Out-Null }
  $lastCol = ($ws.Cells.Item(1, $nCols).Address($false, $false)) -replace '\d', ''
  $lo.Resize($ws.Range("A1:$lastCol$($rows.Count + 1)"))
  $ws.Range("A2:$lastCol$($rows.Count + 1)").Value2 = $arr
  $wb.Worksheets.Item("Team Lookup").Range("B3").Value2 = 2040
  # The real TBA Flat Schedule, pasted the way the data lead would (one line per cell from A5)
  $flat = @(Get-Content "$dir\tba_flat_schedule_2026ilpe.csv" | Where-Object { $_ })
  $schedWs = $wb.Worksheets | Where-Object { $_.Name -eq "TBA Schedule" }
  if ($schedWs -and $flat.Count) {
    $lines = New-Object 'object[,]' $flat.Count, 1
    # [string]: Get-Content lines carry PowerShell extras that Excel can't take ("insufficient memory")
    for ($i = 0; $i -lt $flat.Count; $i++) { $lines[$i, 0] = [string]$flat[$i] }
    $schedWs.Range("A5:A$($flat.Count + 4)").Value2 = $lines
  }
  $x.CalculateFull()
  $sum = $wb.Worksheets.Item("Team Summary")
  "Table rows: $($lo.ListRows.Count); summary teams: " + $x.WorksheetFunction.CountA($sum.Range("A5:A304"))
  "Pick list top 5: " + ((5..9 | ForEach-Object { "$($sum.Range("U$_").Text) ($($sum.Range("V$_").Text))" }) -join ', ')
  if ($schedWs) {
    $chk = $wb.Worksheets.Item("Scouting Check")
    "Scouting Check: " + ((5..11 | ForEach-Object { $chk.Range("A$_").Text + "=" + $chk.Range("B$_").Text }) -join ", ") + " | " + $chk.Range("A15").Text
  }
  $sum.Activate(); $sum.Range("A5").Select() | Out-Null
  $wb.Save(); $wb.Close($false); "Saved $Workbook"
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
