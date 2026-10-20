# Save a snapshot of a The Blue Alliance event (teams, schedule + results, rankings) for matches.html.
# Reads the public event page, so no TBA API key is needed. Re-run during the event to refresh.
# Usage: powershell -ExecutionPolicy Bypass -File fetch-tba.ps1 [-EventKey 2026iltrr]
# Writes data\tba-<event>.js
param([string]$EventKey = "2026iltrr")
$ErrorActionPreference = "Stop"
$EventKey = $EventKey.ToLower()
$h = (Invoke-WebRequest "https://www.thebluealliance.com/event/$EventKey" -UseBasicParsing -TimeoutSec 60).Content

function Text($s) { [Net.WebUtility]::HtmlDecode((($s -replace '<br\s*/?>', ' ' -replace '<[^>]+>', ' ') -replace '\s+', ' ').Trim()) }
function Num($s) {
  $n = 0.0
  if ([double]::TryParse((Text $s), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) { $n } else { $null }
}

$name = $EventKey
$og = [regex]::Match($h, '<meta property="og:title" content="([^"]*)"')
if ($og.Success) { $name = Text $og.Groups[1].Value }

# Team list tab: <a href="/team/648/2026">648<br>Nickname</a>
$teams = [ordered]@{}
foreach ($t in [regex]::Matches($h, '<a href="/team/(\d+)/\d{4}">\d+<br>([^<]*)</a>')) { $teams[$t.Groups[1].Value] = Text $t.Groups[2].Value }

# Match tables: key, red 1-3, blue 1-3, scores (missing until the match is played)
$byKey = [ordered]@{}
foreach ($r in [regex]::Matches($h, '(?s)<tr class="visible-lg">(.*?)</tr>')) {
  $row = $r.Groups[1].Value
  # Keys look like 2026ilpe_qm12 (quals) or 2026ilpe_sf5m1 (playoff set 5, match 1)
  $k = [regex]::Match($row, "/match/($($EventKey)_(qm|ef|qf|sf|f)(\d+)(?:m(\d+))?)""")
  if (-not $k.Success -or $byKey.Contains($k.Groups[1].Value)) { continue }
  $ts = @([regex]::Matches($row, '<a href="/team/(\d+)/') | ForEach-Object { [int]$_.Groups[1].Value })
  if ($ts.Count -ne 6) { continue }
  $score = @{}
  foreach ($side in "red", "blue") {
    $cell = [regex]::Match($row, "(?s)<td class=""$($side)Score[^""]*"">(.*?)</td>")
    $s = if ($cell.Success) { [regex]::Match($cell.Groups[1].Value, '<span[^>]*>(-?\d+)</span>') }
    $score[$side] = if ($s -and $s.Success -and [int]$s.Groups[1].Value -ge 0) { [int]$s.Groups[1].Value } else { $null }
  }
  $qual = $k.Groups[2].Value -eq "qm"
  $set = if ($qual) { 1 } else { [int]$k.Groups[3].Value }
  $num = if ($qual) { [int]$k.Groups[3].Value } else { [int]$k.Groups[4].Value }
  $byKey[$k.Groups[1].Value] = [ordered]@{
    key = $k.Groups[1].Value; level = $k.Groups[2].Value; set = $set; num = $num
    red = @($ts[0..2]); blue = @($ts[3..5]); redScore = $score.red; blueScore = $score.blue; time = $null
  }
}

# Rankings tab. Columns between "Team" and "Record" change every season, so keep them by name.
$rankCols = @(); $rankings = @()
$rt = [regex]::Match($h, '(?s)id="rankingsTable">(.*?)</table>')
if ($rt.Success) {
  $head = @([regex]::Matches($rt.Groups[1].Value, '<th>([^<]*)</th>') | ForEach-Object { Text $_.Groups[1].Value })
  $iRank = [array]::IndexOf($head, "Rank"); $iTeam = [array]::IndexOf($head, "Team")
  $iRec = (0..($head.Count - 1) | Where-Object { $head[$_] -like "Record*" } | Select-Object -First 1)
  $iDq = [array]::IndexOf($head, "DQ"); $iPlayed = [array]::IndexOf($head, "Played")
  $iEnd = if ($null -ne $iRec) { $iRec } else { $head.Count }
  if ($iEnd - 1 -gt $iTeam) { $rankCols = @($head[($iTeam + 1)..($iEnd - 1)]) }
  $body = [regex]::Match($rt.Groups[1].Value, '(?s)<tbody>(.*?)</tbody>').Groups[1].Value
  foreach ($tr in [regex]::Matches($body, '(?s)<tr>(.*?)</tr>')) {
    $cells = @([regex]::Matches($tr.Groups[1].Value, '(?s)<td>(.*?)</td>') | ForEach-Object { $_.Groups[1].Value })
    if ($cells.Count -lt $head.Count) { continue }
    $rec = if ($null -ne $iRec) { (Text $cells[$iRec]) -split "-" } else { @() }
    $rankings += [ordered]@{
      rank = [int](Num $cells[$iRank]); team = [int](Num $cells[$iTeam])
      vals = @(for ($i = $iTeam + 1; $i -lt $iEnd; $i++) { Num $cells[$i] })
      w = if ($rec.Count -eq 3) { [int]$rec[0] } else { $null }
      l = if ($rec.Count -eq 3) { [int]$rec[1] } else { $null }
      t = if ($rec.Count -eq 3) { [int]$rec[2] } else { $null }
      dq = if ($iDq -ge 0) { [int](Num $cells[$iDq]) } else { 0 }
      played = if ($iPlayed -ge 0) { [int](Num $cells[$iPlayed]) } else { $null }
    }
  }
}

$data = [ordered]@{
  key = $EventKey; name = $name; fetched = (Get-Date).ToUniversalTime().ToString("o")
  teams = $teams; rankCols = $rankCols; rankings = $rankings; matches = @($byKey.Values)
}
$json = ConvertTo-Json -InputObject $data -Depth 6 -Compress
$js = "// Snapshot of https://www.thebluealliance.com/event/$EventKey saved by fetch-tba.ps1 on $(Get-Date -Format 'yyyy-MM-dd HH:mm')`n" +
      "window.TBA_EVENTS = window.TBA_EVENTS || {};`nTBA_EVENTS[""$EventKey""] = $json;`n"
$out = Join-Path $PSScriptRoot "data\tba-$EventKey.js"
[IO.File]::WriteAllText($out, $js, (New-Object Text.UTF8Encoding $false))
"$name : $($teams.Count) teams, $($byKey.Count) matches ($(@($byKey.Values | Where-Object { $null -ne $_.redScore }).Count) played), $($rankings.Count) ranked"
"Saved $out"
