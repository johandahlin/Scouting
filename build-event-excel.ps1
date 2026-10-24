# Build the event workbook from The Blue Alliance: TBA Flat Schedule and Team List pasted in, plus one Scans row per robot
# for every qual match already played, ESTIMATED from TBA's results (scout name "TBA estimate"):
#   - real per robot:  Tower climb in auto, Tower climb (endgame)
#   - estimated:       Auto / Teleop fuel scored = the alliance's real fuel split between its 3 robots by OPR
#   - left blank:      everything TBA can't know (misses, intake, ratings, pick, fouls per robot, notes...)
# The dashboard's "Use event data" button loads this file.
# Usage: powershell -ExecutionPolicy Bypass -File build-event-excel.ps1 [-EventKey 2026iltrr] [-Out <file.xlsx>]
# Without -EventKey it uses the event in data\event.js. Re-run during the event to add newly played matches (it starts a new workbook).
param([string]$EventKey = "", [string]$Out = (Join-Path $PSScriptRoot "DERT_Scouting_REBUILTEventData.xlsx"))
$ErrorActionPreference = "Stop"
if (-not $EventKey) { $EventKey = [regex]::Match((Get-Content (Join-Path $PSScriptRoot "data\event.js") -Raw), 'DERT_EVENT\s*=\s*"([^"]+)"').Groups[1].Value }
$EventKey = $EventKey.ToLower()

# The Scouting tab on the event page holds the same CSVs as its "Copy to Clipboard" buttons
$h = (Invoke-WebRequest "https://www.thebluealliance.com/event/$EventKey" -UseBasicParsing -TimeoutSec 60).Content
function CsvLines($id) {
  $m = [regex]::Match($h, "(?s)<textarea id=""$id""[^>]*>(.*?)</textarea>")
  if (-not $m.Success) { return @() }
  @([Net.WebUtility]::HtmlDecode($m.Groups[1].Value).Trim() -split "\r?\n" | Where-Object { $_ })
}
[string[]]$schedule = CsvLines "flat-schedule-csv"   # one robot per line, what the TBA Schedule sheet reads
[string[]]$teamList = CsvLines "team-list-csv"
if ($schedule.Count -lt 2) { throw "TBA has no Flat Schedule yet for $EventKey" }
if ($teamList.Count -lt 2) { throw "TBA has no Team List yet for $EventKey" }

# ---------- played qual matches ----------
# Robots in schedule order (red 1-3, blue 1-3). Team ids stay text: off-season events have B robots such as 1706B.
$quals = [ordered]@{}
foreach ($s in ($schedule | ConvertFrom-Csv | Where-Object { $_.comp_level -eq "qm" })) {
  $k = [string]$s.match_number   # text key: an int would index the ordered dictionary by position
  if (-not $quals.Contains($k)) { $quals[$k] = @{ num = [int]$k; red = @(); blue = @(); time = "$($s.scheduled_date) $($s.scheduled_time)" } }
  $quals[$k][$s.color] += [string]$s.team
}
function CellText($s) { (($s -replace '<[^>]+>', ' ') -replace '&nbsp;', ' ' -replace '\s+', ' ').Trim() }
$played = @()
foreach ($m in $quals.Values) {
  $p = (Invoke-WebRequest "https://www.thebluealliance.com/match/$($EventKey)_qm$($m.num)" -UseBasicParsing -TimeoutSec 60).Content
  $bd = @{}; $autoTower = @{ red = @(); blue = @() }; $endgame = @{ red = @("", "", ""); blue = @("", "", "") }
  $trs = [regex]::Matches($p, '(?s)<tr[^>]*>\s*<td class="(red|redScore)" colspan="2">(.*?)</td>\s*<t[dh]>(.*?)</t[dh]>\s*<td class="(blue|blueScore)" colspan="2">(.*?)</td>')
  foreach ($t in $trs) {
    $label = CellText $t.Groups[3].Value
    $cells = @{ red = $t.Groups[2].Value; blue = $t.Groups[5].Value }
    if ($label -eq "Auto Tower") {
      # one ok/remove icon per robot, in robot order
      foreach ($side in "red", "blue") { $autoTower[$side] = @([regex]::Matches($cells[$side], 'glyphicon-(ok|remove)') | ForEach-Object { $_.Groups[1].Value -eq "ok" }) }
    } elseif ($label -match '^Robot (\d) Endgame$') {
      # by position: a B robot carries its team's number in the tooltip, so the title can't tell 2338 and 2338B apart
      $pos = [int]$Matches[1] - 1
      foreach ($side in "red", "blue") { $endgame[$side][$pos] = CellText $cells[$side] }
    } else {
      $bd[$label] = @((CellText $cells.red), (CellText $cells.blue))
    }
  }
  if (-not $bd["Total Score"] -or $bd["Total Score"][0] -notmatch '^\d+$') { continue }   # not played yet
  $m.bd = $bd; $m.autoTower = $autoTower; $m.endgame = $endgame
  $played += $m
}
"$EventKey : $($quals.Count) qual matches on the schedule, $($played.Count) played"

# ---------- OPR: each robot's share of its alliance's fuel ----------
# Ridge least squares (pulled toward the average robot): early in an event there are more robots than matches.
$robots = @($played | ForEach-Object { $_.red + $_.blue } | Sort-Object -Unique)
$idx = @{}; for ($i = 0; $i -lt $robots.Count; $i++) { $idx[$robots[$i]] = $i }
function Opr($label) {
  $T = $robots.Count; $A = New-Object 'double[,]' $T, ($T + 1)
  $sum = 0.0; $cnt = 0
  foreach ($m in $played) { for ($s = 0; $s -lt 2; $s++) { $sum += [double]$m.bd[$label][$s]; $cnt += 3 } }
  $avg = $sum / [Math]::Max(1, $cnt); $lambda = 0.5
  for ($i = 0; $i -lt $T; $i++) { $A[$i, $i] += $lambda; $A[$i, $T] += $lambda * $avg }
  foreach ($m in $played) {
    for ($s = 0; $s -lt 2; $s++) {
      $al = if ($s -eq 0) { $m.red } else { $m.blue }; $v = [double]$m.bd[$label][$s]
      foreach ($ta in $al) { foreach ($tb in $al) { $A[$idx[$ta], $idx[$tb]] += 1 }; $A[$idx[$ta], $T] += $v }
    }
  }
  for ($c = 0; $c -lt $T; $c++) {   # Gauss-Jordan; the ridge term keeps every pivot positive
    for ($r = 0; $r -lt $T; $r++) {
      if ($r -eq $c) { continue }
      $f = $A[$r, $c] / $A[$c, $c]
      if ($f -ne 0) { for ($k = $c; $k -le $T; $k++) { $A[$r, $k] -= $f * $A[$c, $k] } }
    }
  }
  $res = @{}; for ($i = 0; $i -lt $T; $i++) { $res[$robots[$i]] = $A[$i, $T] / $A[$i, $i] }
  $res
}
function Split([int]$total, $weights) {   # integer split that keeps the exact alliance total
  $sum = ($weights | Measure-Object -Sum).Sum
  $raw = @($weights | ForEach-Object { $total * $_ / $sum })
  $res = @($raw | ForEach-Object { [int][Math]::Floor($_) })
  $left = $total - ($res | Measure-Object -Sum).Sum
  $order = 0..2 | Sort-Object { - ($raw[$_] - [Math]::Floor($raw[$_])) }
  for ($i = 0; $i -lt $left; $i++) { $res[$order[$i]]++ }
  $res
}
$oprAuto = Opr "Auto Fuel"; $oprTele = Opr "Teleop Fuel Points"

# ---------- one Scans row per robot, by column name ----------
$TELE_CLIMB = @{ "Level 1" = 10; "Level 2" = 20; "Level 3" = 30 }
function Level($s) { if (-not $s -or $s -eq "None") { "None" } elseif ($s -match '(\d)') { "Level $($Matches[1])" } else { "Level 1" } }
$scanRows = @()
foreach ($m in $played) {
  for ($s = 0; $s -lt 2; $s++) {
    $side = if ($s -eq 0) { "red" } else { "blue" }; $al = $m[$side]
    $auto = Split ([int]$m.bd["Auto Fuel"][$s]) @($al | ForEach-Object { [Math]::Max(1, $oprAuto[$_]) })
    $tele = Split ([int]$m.bd["Teleop Fuel Points"][$s]) @($al | ForEach-Object { [Math]::Max(1, $oprTele[$_]) })
    $when = [datetime]::ParseExact($m.time, "yyyy-MM-dd hh:mm tt", [Globalization.CultureInfo]::InvariantCulture)
    for ($i = 0; $i -lt 3; $i++) {
      $autoClimb = if ($m.autoTower[$side][$i]) { "Level 1" } else { "None" }
      $climb = Level $m.endgame[$side][$i]
      $autoPts = $auto[$i] + $(if ($autoClimb -eq "Level 1") { 15 } else { 0 })
      $climbPts = [int]$TELE_CLIMB[$climb]
      $team = $al[$i]
      $scanRows += @{
        "Match type" = "Qual"; "Match #" = $m.num; "Team #" = $(if ($team -match '^\d+$') { [double]$team } else { $team })
        "Alliance" = (Get-Culture).TextInfo.ToTitleCase($side); "Station" = $i + 1; "Scout name" = "TBA estimate"; "Robot on field?" = "Yes"
        "Auto fuel scored" = $auto[$i]; "Tower climb in auto" = $autoClimb; "Teleop fuel scored" = $tele[$i]; "Tower climb" = $climb
        "Observations" = "Estimated from TBA results: alliance fuel split by OPR, climbs are real. Not scouted."
        "Auto pts" = $autoPts; "Teleop fuel pts" = $tele[$i]; "Climb pts" = $climbPts; "Total pts" = $autoPts + $tele[$i] + $climbPts
        "Submitted" = $when.ToString("M/d/yyyy, h:mm tt", [Globalization.CultureInfo]::InvariantCulture)
      }
    }
  }
}

# Stamp in column A, same as the web app (see build-excel.ps1 / index.html)
function Fnv1a([string]$s) {
  [uint64]$f = 2166136261
  foreach ($b in [Text.Encoding]::UTF8.GetBytes($s)) { $f = (($f -bxor $b) * [uint64]16777619) -band [uint64]4294967295 }
  '{0:x8}' -f $f
}

# One CSV line per cell from A5, the way it lands when the data lead presses Ctrl+V
function PasteLines($ws, [string[]]$csv) {
  $cells = New-Object 'object[,]' $csv.Count, 1
  for ($i = 0; $i -lt $csv.Count; $i++) { $cells[$i, 0] = $csv[$i] }
  $ws.Range("A5:A" + ($csv.Count + 4)).Value2 = $cells
}

# Separate process: after build-excel.ps1 has run in this session, Excel rejects the paste below ("Unable to cast Object[,]")
powershell -STA -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot "build-excel.ps1") -Out $Out
if (-not (Test-Path $Out)) { throw "build-excel.ps1 did not create $Out" }

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false
try {
  $wb = $x.Workbooks.Open($Out)
  $x.Calculation = -4135   # manual while writing, so Excel isn't busy recalculating between calls
  PasteLines $wb.Worksheets.Item("TBA Schedule") $schedule
  PasteLines $wb.Worksheets.Item("TBA Teams") $teamList
  if ($scanRows.Count) {
    $ws = $wb.Worksheets.Item("Scans"); $lo = $ws.ListObjects.Item("Scans")
    $hdr = @($lo.HeaderRowRange.Value2 | ForEach-Object { "$_" })
    $stamp = "S$($hdr.Count - 1)-$(Fnv1a ($hdr[1..($hdr.Count - 1)] -join "`t"))"
    $arr = New-Object 'object[,]' $scanRows.Count, $hdr.Count
    for ($r = 0; $r -lt $scanRows.Count; $r++) {
      $arr[$r, 0] = $stamp
      for ($c = 1; $c -lt $hdr.Count; $c++) { $arr[$r, $c] = $scanRows[$r][$hdr[$c]] }
    }
    $lastCol = ($ws.Cells.Item(1, $hdr.Count).Address($false, $false)) -replace '\d', ''
    $lo.Resize($ws.Range("A1:$lastCol$($scanRows.Count + 1)"))
    $ws.Range("A2:$lastCol$($scanRows.Count + 1)").Value2 = $arr
  }
  $x.Calculation = -4105; $x.CalculateFull()
  $chk = $wb.Worksheets.Item("Scouting Check")
  "$($schedule.Count - 1) schedule lines (one per robot), $($teamList.Count - 1) teams, $($scanRows.Count) estimated scans"
  "Scouting Check: " + ((5..13 | ForEach-Object { $chk.Range("A$_").Text + "=" + $chk.Range("B$_").Text }) -join ", ") + " | " + $chk.Range("A17").Text
  $info = $wb.Worksheets.Item("Instructions"); $info.Activate(); [void]$info.Range("A1").Select()
  $wb.Save(); $wb.Close($false); "Saved $Out"
} finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }
