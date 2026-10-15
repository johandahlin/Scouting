# Download real 2026ilpe qualification results from The Blue Alliance website (HTML) into matches.json
$ErrorActionPreference = "Stop"
$out = $PSScriptRoot
$ev = "2026ilpe"
$h = [IO.File]::ReadAllText("$out\tba_event.html")

# Match rows on the event page: key + 6 teams in order (red1-3, blue1-3)
$matches = @()
$rows = [regex]::Matches($h, '(?s)<tr class="visible-lg">(.*?)</tr>')
foreach ($r in $rows) {
  $m = [regex]::Match($r.Groups[1].Value, "/match/($($ev)_qm(\d+))""")
  if (-not $m.Success) { continue }
  $teams = [regex]::Matches($r.Groups[1].Value, '<a href="/team/(\d+)/2026">') | ForEach-Object { [int]$_.Groups[1].Value }
  if ($teams.Count -ne 6) { continue }
  $matches += [pscustomobject]@{ key = $m.Groups[1].Value; num = [int]$m.Groups[2].Value; red = $teams[0..2]; blue = $teams[3..5] }
}
$matches = $matches | Sort-Object num -Unique
"Found $($matches.Count) qual matches"

function CellText($s) { (($s -replace '<[^>]+>', ' ') -replace '&nbsp;', ' ' -replace '\s+', ' ').Trim() }

$result = @()
foreach ($m in $matches) {
  $p = (Invoke-WebRequest "https://www.thebluealliance.com/match/$($m.key)" -UseBasicParsing -TimeoutSec 60).Content
  $bd = @{}
  $autoTower = @{}
  $endgame = @{}
  $trs = [regex]::Matches($p, '(?s)<tr[^>]*>\s*<td class="(red|redScore)" colspan="2">(.*?)</td>\s*<t[dh]>(.*?)</t[dh]>\s*<td class="(blue|blueScore)" colspan="2">(.*?)</td>')
  foreach ($t in $trs) {
    $label = CellText $t.Groups[3].Value
    $redHtml = $t.Groups[2].Value; $blueHtml = $t.Groups[5].Value
    if ($label -eq "Auto Tower") {
      foreach ($side in @($redHtml, $blueHtml)) {
        foreach ($g in [regex]::Matches($side, 'glyphicon-(ok|remove)" rel="tooltip" title="(\d+)"')) { $autoTower[$g.Groups[2].Value] = ($g.Groups[1].Value -eq "ok") }
      }
    } elseif ($label -like "Robot * Endgame") {
      foreach ($side in @($redHtml, $blueHtml)) {
        $g = [regex]::Match($side, 'title="(\d+)">([^<]*)<')
        if ($g.Success) { $endgame[$g.Groups[1].Value] = $g.Groups[2].Value.Trim() }
      }
    } else {
      $bd[$label] = @((CellText $redHtml), (CellText $blueHtml))
    }
  }
  $result += [pscustomobject]@{ key = $m.key; num = $m.num; red = $m.red; blue = $m.blue; breakdown = $bd; autoTower = $autoTower; endgame = $endgame }
  Write-Host "$($m.key): red $($bd['Total Score'][0]) blue $($bd['Total Score'][1])"
}
$result | ConvertTo-Json -Depth 6 | Out-File "$out\matches.json" -Encoding utf8
"Saved $($result.Count) matches"
