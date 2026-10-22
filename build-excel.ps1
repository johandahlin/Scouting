param([string]$Out = "D:\JD Projects\DERT\QRCodeTest\DERT_Scouting_REBUILT.xlsx", [switch]$Test)
$ErrorActionPreference = "Stop"

# Must match the column order produced by index.html (Schema stamp + FIELDS + computed columns)
$headers = @(
  "Schema","Match type","Match #","Team #","Alliance","Station","Scout name","Robot on field?","Starting position","Fuel preloaded",
  "Auto fuel scored","Auto fuel missed","Auto path","Picked up fuel from","Tower climb in auto",
  "Teleop fuel scored","Teleop fuel missed","Intake from","Shoots from","Drives over / through","Inactive Hub behavior",
  "Fuel it can carry","Played defense","Was defended","Tower climb","Climb time",
  "Driver skill","Robot speed","Shooting accuracy","Problems","Minor fouls","Major fouls","Card","Pick?","Observations",
  "Auto pts","Teleop fuel pts","Climb pts","Total pts","Submitted"
)
$n = $headers.Count

# Stamp the web app puts in column A: "S<columns after it>-<FNV-1a hash of their headers>" (same as index.html)
function Fnv1a([string]$s) {
  [uint64]$h = 2166136261
  foreach ($b in [Text.Encoding]::UTF8.GetBytes($s)) { $h = (($h -bxor $b) * [uint64]16777619) -band [uint64]4294967295 }
  '{0:x8}' -f $h
}
$schema = "S$($n - 1)-$(Fnv1a ($headers[1..($n - 1)] -join "`t"))"

# Excel rejects COM calls while it is busy calculating ("Call was rejected by callee"); this filter makes them wait and retry
if (-not ("ExcelMessageFilter" -as [type])) {
  Add-Type @"
using System;
using System.Runtime.InteropServices;
[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public interface IOleMessageFilter {
  [PreserveSig] int HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr interfaceInfo);
  [PreserveSig] int RetryRejectedCall(IntPtr callee, int tickCount, int rejectType);
  [PreserveSig] int MessagePending(IntPtr callee, int tickCount, int pendingType);
}
public class ExcelMessageFilter : IOleMessageFilter {
  [DllImport("ole32.dll")] static extern int CoRegisterMessageFilter(IOleMessageFilter filter, out IOleMessageFilter old);
  public static void Register() { IOleMessageFilter old; CoRegisterMessageFilter(new ExcelMessageFilter(), out old); }
  public static void Revoke() { IOleMessageFilter old; CoRegisterMessageFilter(null, out old); }
  public int HandleInComingCall(int callType, IntPtr caller, int tickCount, IntPtr interfaceInfo) { return 0; }
  // 2 = SERVERCALL_RETRYLATER: try again after 100 ms, give up after 2 minutes
  public int RetryRejectedCall(IntPtr callee, int tickCount, int rejectType) { return rejectType == 2 && tickCount < 120000 ? 100 : -1; }
  public int MessagePending(IntPtr callee, int tickCount, int pendingType) { return 2; }
}
"@
}
[ExcelMessageFilter]::Register()

function RGB($r,$g,$b) { $r + 256*$g + 65536*$b }
$TEAL = RGB 18 180 175; $DARK = RGB 28 28 28; $LIGHT = RGB 249 249 249; $YELLOW = RGB 255 242 153; $WHITE = RGB 255 255 255

$x = New-Object -ComObject Excel.Application
$x.Visible = $false; $x.DisplayAlerts = $false; $x.ScreenUpdating = $false
try {
  $wb = $x.Workbooks.Add()
  # Manual calculation while building: otherwise Excel is still recalculating when the next call arrives and rejects it
  $x.Calculation = -4135
  while ($wb.Worksheets.Count -lt 6) { [void]$wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
  $wb.Styles.Item("Normal").Font.Name = "Arial"
  $wb.Styles.Item("Normal").Font.Size = 10

  $info  = $wb.Worksheets.Item(1); $info.Name  = "Instructions"
  $scan  = $wb.Worksheets.Item(2); $scan.Name  = "Scans"
  $sum   = $wb.Worksheets.Item(3); $sum.Name   = "Team Summary"
  $look  = $wb.Worksheets.Item(4); $look.Name  = "Team Lookup"
  $check = $wb.Worksheets.Item(5); $check.Name = "Scouting Check"
  $sched = $wb.Worksheets.Item(6); $sched.Name = "TBA Schedule"

  # Long formulas are written over several lines below for reading; Excel gets them on one line
  function OneLine([string]$f) { $f -replace '\r?\n\s*', '' }

  function Title($ws, $text, $sub, $lastCol) {
    $r = $ws.Range("A1:$($lastCol)1"); $r.Interior.Color = $DARK
    $ws.Range("A1").Value2 = $text; $ws.Range("A1").Font.Size = 16; $ws.Range("A1").Font.Bold = $true; $ws.Range("A1").Font.Color = $LIGHT
    $ws.Rows.Item(1).RowHeight = 28
    $r2 = $ws.Range("A2:$($lastCol)2"); $r2.Interior.Color = $TEAL
    $ws.Range("A2").Value2 = $sub; $ws.Range("A2").Font.Color = $DARK; $ws.Range("A2").Font.Bold = $true
  }
  function HeaderStyle($r) {
    $r.Interior.Color = $TEAL; $r.Font.Color = $DARK; $r.Font.Bold = $true; $r.WrapText = $true; $r.VerticalAlignment = -4108
  }

  # ---------------- TBA Schedule ----------------
  # The data lead pastes TBA's Flat Schedule CSV into A5 (one robot per line: match_key,date,time,comp_level,
  # match_number,set_number,color,team). Excel either keeps each line in one cell or splits it over A:H;
  # the formula in R5 reads both and returns the qual matches as Match # | Alliance | Team #.
  Title $sched "DERT #2040 - TBA Schedule" "Paste TBA's Flat Schedule here so the Scouting Check can find unscouted robots, team-number typos and wrong alliances." "T"
  $sched.Range("A4").Value2 = "Paste the Flat Schedule here: click the yellow cell, then press Ctrl+V"
  $sched.Range("A4").Font.Bold = $true
  $sched.Range("A5").Interior.Color = $YELLOW; $sched.Range("A5").Borders.Weight = 2
  $sched.Columns.Item(1).ColumnWidth = 62
  $howTo = @(
    "HOW TO GET THE SCHEDULE",
    "1. On thebluealliance.com open your event (for example thebluealliance.com/event/2026iltrr) and click the Scouting tab.",
    "2. Under Flat Schedule, click Copy to Clipboard.",
    "3. Click the yellow cell A5 on this sheet and press Ctrl+V.",
    "4. Done: the Scouting Check sheet now compares every scan with the schedule.",
    "Paste again (over the old one) if the schedule changes. TBA usually has it the day before quals.",
    "To remove it: select columns A to H from row 5 down and press Delete."
  )
  for ($i = 0; $i -lt $howTo.Count; $i++) { $sched.Cells.Item(4 + $i, 10).Value2 = $howTo[$i] }
  $sched.Range("J4").Font.Bold = $true; $sched.Columns.Item(10).ColumnWidth = 3
  $sched.Range("R3").Value2 = "Read from the paste (don't type here)"; $sched.Range("R3").Font.Italic = $true; $sched.Range("R3").Font.Color = RGB 120 120 120
  $sched.Range("R4").Value2 = "Match #"; $sched.Range("S4").Value2 = "Alliance"; $sched.Range("T4").Value2 = "Team #"
  HeaderStyle $sched.Range("R4:T4")
  foreach ($c in 18..20) { $sched.Columns.Item($c).ColumnWidth = 10 }
  # Built before the other sheets: adding it after Team Summary / Team Lookup crashes Excel, and so does a LAMBDA over 3000 rows.
  $sched.Range("R5").Formula2 = OneLine @'
=LET(src,A1:H2000,a,INDEX(src,,1),
line,IF(INDEX(src,,2)="",a,a&","&INDEX(src,,2)&","&INDEX(src,,3)&","&INDEX(src,,4)&","&INDEX(src,,5)&","&INDEX(src,,6)&","&INDEX(src,,7)&","&INDEX(src,,8)),
t,","&line&",",
lvl,IFERROR(TEXTBEFORE(TEXTAFTER(t,",",4),","),""),
num,IFERROR(--TEXTBEFORE(TEXTAFTER(t,",",5),","),""),
side,LOWER(IFERROR(TEXTBEFORE(TEXTAFTER(t,",",7),","),"")),
tm,IFERROR(--TEXTBEFORE(TEXTAFTER(t,",",8),","),""),
keep,ISNUMBER(num)*ISNUMBER(tm)*(lvl="qm")*((side="red")+(side="blue")),
IFERROR(FILTER(HSTACK(num,PROPER(side),tm),keep),HSTACK("","","")))
'@

  # ---------------- Scans ----------------
  for ($i = 0; $i -lt $n; $i++) { $scan.Cells.Item(1, $i + 1).Value2 = $headers[$i] }
  $lastColLetter = ($scan.Cells.Item(1, $n).Address($false, $false)) -replace '\d',''
  $lo = $scan.ListObjects.Add(1, $scan.Range("A1:$($lastColLetter)2"), [Type]::Missing, 1)
  $lo.Name = "Scans"
  $lo.TableStyle = "TableStyleLight1"
  HeaderStyle $lo.HeaderRowRange
  $scan.Rows.Item(1).RowHeight = 42
  $widths = @{ "Schema"=13; "Scout name"=14; "Auto path"=30; "Picked up fuel from"=22; "Intake from"=30; "Shoots from"=26; "Inactive Hub behavior"=30;
               "Problems"=24; "Observations"=50; "Submitted"=20; "Drives over / through"=14; "Starting position"=11 }
  for ($i = 0; $i -lt $n; $i++) {
    $w = 10; if ($widths.ContainsKey($headers[$i])) { $w = $widths[$headers[$i]] }
    $scan.Columns.Item($i + 1).ColumnWidth = $w
  }
  # Points columns highlighted
  foreach ($h in "Auto pts","Teleop fuel pts","Climb pts","Total pts") {
    $lo.ListColumns.Item($h).DataBodyRange.EntireColumn.Interior.Color = $LIGHT
    $lo.ListColumns.Item($h).DataBodyRange.EntireColumn.Font.Bold = $true
  }
  HeaderStyle $lo.HeaderRowRange
  $scan.Activate(); $x.ActiveWindow.SplitRow = 1; $x.ActiveWindow.SplitColumn = 4; $x.ActiveWindow.FreezePanes = $true
  $scan.Range("A2").Select()
  # A row whose stamp doesn't match these headers came from another version of the web form: show it red.
  # No functions or argument separators in the formula, so it works in any Excel language.
  $fc = $scan.Range("A2:$($lastColLetter)5000").FormatConditions.Add(2, [Type]::Missing, "=(`$A2<>`"`")*(`$A2<>`"$schema`")")
  $fc.Interior.Color = RGB 255 199 206; $fc.Font.Color = RGB 156 0 6

  # ---------------- Team Summary ----------------
  Title $sum "DERT #2040 - REBUILT Team Summary" "Updates automatically from the Scans sheet. One row per team, sorted by team number. Rank 1 = highest average total points." "R"
  $sumCols = @(
    @("Team #",              '=IFERROR(SORT(UNIQUE(FILTER(Scans[Team ''#],Scans[Team ''#]<>""))),"")', "0"),
    @("Matches",             '=IF(A5#="","",COUNTIFS(Scans[Team ''#],A5#))', "0"),
    @("Avg auto pts",        '=IF(A5#="","",AVERAGEIFS(Scans[Auto pts],Scans[Team ''#],A5#))', "0.0"),
    @("Avg teleop fuel",     '=IF(A5#="","",AVERAGEIFS(Scans[Teleop fuel pts],Scans[Team ''#],A5#))', "0.0"),
    @("Avg climb pts",       '=IF(A5#="","",AVERAGEIFS(Scans[Climb pts],Scans[Team ''#],A5#))', "0.0"),
    @("Avg total pts",       '=IF(A5#="","",AVERAGEIFS(Scans[Total pts],Scans[Team ''#],A5#))', "0.0"),
    @("Best total",          '=IF(A5#="","",MAXIFS(Scans[Total pts],Scans[Team ''#],A5#))', "0"),
    @("Shot accuracy",       '=IF(A5#="","",IFERROR((SUMIFS(Scans[Auto fuel scored],Scans[Team ''#],A5#)+SUMIFS(Scans[Teleop fuel scored],Scans[Team ''#],A5#))/(SUMIFS(Scans[Auto fuel scored],Scans[Team ''#],A5#)+SUMIFS(Scans[Teleop fuel scored],Scans[Team ''#],A5#)+SUMIFS(Scans[Auto fuel missed],Scans[Team ''#],A5#)+SUMIFS(Scans[Teleop fuel missed],Scans[Team ''#],A5#)),""))', "0%"),
    @("Auto climb rate",     '=IF(A5#="","",COUNTIFS(Scans[Team ''#],A5#,Scans[Tower climb in auto],"Level 1")/B5#)', "0%"),
    @("Endgame climb rate",  '=IF(A5#="","",COUNTIFS(Scans[Team ''#],A5#,Scans[Tower climb],"Level*")/B5#)', "0%"),
    @("Level 3 climbs",      '=IF(A5#="","",COUNTIFS(Scans[Team ''#],A5#,Scans[Tower climb],"Level 3"))', "0"),
    @("Avg driver skill",    '=IF(A5#="","",IFERROR(AVERAGEIFS(Scans[Driver skill],Scans[Team ''#],A5#),""))', "0.0"),
    @("Defense matches",     '=IF(A5#="","",COUNTIFS(Scans[Team ''#],A5#,Scans[Played defense],"Some")+COUNTIFS(Scans[Team ''#],A5#,Scans[Played defense],"A lot"))', "0"),
    @("Matches w/ problems", '=IF(A5#="","",COUNTIFS(Scans[Team ''#],A5#,Scans[Problems],"<>"))', "0"),
    @("Total fouls",         '=IF(A5#="","",SUMIFS(Scans[Minor fouls],Scans[Team ''#],A5#)+SUMIFS(Scans[Major fouls],Scans[Team ''#],A5#))', "0"),
    @("Pick: Yes votes",     '=IF(A5#="","",COUNTIFS(Scans[Team ''#],A5#,Scans[Pick?],"Yes"))', "0"),
    @("No-shows",            '=IF(A5#="","",COUNTIFS(Scans[Team ''#],A5#,Scans[Robot on field?],"No show"))', "0"),
    @("Rank",                '=IF(A5#="","",RANK.EQ(F5#,F5#))', "0")
  )
  for ($i = 0; $i -lt $sumCols.Count; $i++) {
    $c = $i + 1
    $sum.Cells.Item(4, $c).Value2 = $sumCols[$i][0]
    $sum.Range($sum.Cells.Item(5, $c), $sum.Cells.Item(304, $c)).NumberFormat = $sumCols[$i][2]
    $sum.Cells.Item(5, $c).Formula2 = $sumCols[$i][1]
    $sum.Columns.Item($c).ColumnWidth = 11
  }
  HeaderStyle $sum.Range("A4:R4"); $sum.Rows.Item(4).RowHeight = 30
  $sum.Range("A5:A304").Font.Bold = $true
  $sum.Range("A4:R304").HorizontalAlignment = -4108
  # Heat-map the key columns
  foreach ($col in "F","C","D","E","H","J") {
    $cs = $sum.Range("$($col)5:$($col)304").FormatConditions.AddColorScale(3)
    $cs.ColorScaleCriteria.Item(1).FormatColor.Color = RGB 248 105 107
    $cs.ColorScaleCriteria.Item(2).FormatColor.Color = RGB 255 235 132
    $cs.ColorScaleCriteria.Item(3).FormatColor.Color = RGB 99 190 123
  }
  $sum.Activate(); $x.ActiveWindow.SplitRow = 4; $x.ActiveWindow.SplitColumn = 1; $x.ActiveWindow.FreezePanes = $true

  # Pick list: sorted by average total, to the right
  $sum.Range("T3").Value2 = "PICK LIST - sorted by average total points"
  $sum.Range("T3").Font.Bold = $true
  $pick = @("Rank","Team #","Avg total pts","Avg auto pts","Avg teleop fuel","Avg climb pts","Endgame climb rate","Shot accuracy","Matches w/ problems","Pick: Yes votes")
  for ($i = 0; $i -lt $pick.Count; $i++) { $sum.Cells.Item(4, 20 + $i).Value2 = $pick[$i]; $sum.Columns.Item(20 + $i).ColumnWidth = 11 }
  HeaderStyle $sum.Range("T4:AC4")
  $sum.Range("T5").Formula2 = '=IF(A5="","",SEQUENCE(ROWS(A5#)))'
  $sum.Range("U5").Formula2 = '=IF(A5="","",SORTBY(HSTACK(A5#,F5#,C5#,D5#,E5#,J5#,H5#,N5#,P5#),F5#,-1))'
  $sum.Range("T5:AC304").HorizontalAlignment = -4108
  $sum.Range("V5:Y304").NumberFormat = "0.0"; $sum.Range("Z5:AA304").NumberFormat = "0%"
  $sum.Range("U5:U304").Font.Bold = $true
  $sum.Columns.Item(19).ColumnWidth = 3

  # ---------------- Team Lookup ----------------
  Title $look "DERT #2040 - Team Lookup" "Type a team number in the yellow cell to see every scan of that team (including observations)." "H"
  $look.Range("A3").Value2 = "Team #:"; $look.Range("A3").Font.Bold = $true; $look.Range("A3").HorizontalAlignment = -4152
  $look.Range("B3").Interior.Color = $YELLOW; $look.Range("B3").Font.Bold = $true; $look.Range("B3").Font.Size = 14
  $look.Range("B3").Borders.Weight = 2
  $look.Range("C3").Formula2 = '=IF(B3="","",IFERROR(INDEX(''Team Summary''!F5#,MATCH(B3,''Team Summary''!A5#,0)),""))'
  $look.Range("C3").NumberFormat = '"Avg total: "0.0'
  $look.Range("D3").Formula2 = '=IF(B3="","",IFERROR("Rank "&INDEX(''Team Summary''!R5#,MATCH(B3,''Team Summary''!A5#,0))&" of "&ROWS(''Team Summary''!A5#),"No scans yet"))'
  for ($i = 0; $i -lt $n; $i++) { $look.Cells.Item(5, $i + 1).Value2 = $headers[$i] }
  HeaderStyle $look.Range("A5:$($lastColLetter)5"); $look.Rows.Item(5).RowHeight = 42
  $look.Range("A6").Formula2 = '=IF(B3="","",IFERROR(SORTBY(FILTER(Scans,Scans[Team ''#]=B3),FILTER(Scans[Match ''#],Scans[Team ''#]=B3),1),"No scans for this team yet"))'
  for ($i = 0; $i -lt $n; $i++) {
    $w = 10; if ($widths.ContainsKey($headers[$i])) { $w = $widths[$headers[$i]] }
    $look.Columns.Item($i + 1).ColumnWidth = $w
  }
  $look.Columns.Item(3).ColumnWidth = 14
  $look.Activate(); $x.ActiveWindow.SplitRow = 5; $x.ActiveWindow.FreezePanes = $true

  # ---------------- Scouting Check ----------------
  # One list of problems (A15) built from the Scans table and, once pasted, the TBA schedule; the counts above it read that list.
  Title $check "DERT #2040 - Scouting Check" "Problems to fix in the Scans sheet. Updates by itself; paste the TBA schedule on the TBA Schedule sheet for the full check." "G"
  $hasSched = 'COUNT(INDEX(''TBA Schedule''!$R$5#,,1))>0'
  function CountOf($label) { 'COUNTIF(INDEX($A$15#,,1),"' + $label + '")' }
  $summary = @(
    @("Scans", '=COUNTIF(Scans[Team ''#],"<>")', ""),
    @("Duplicate scans (rows)", ('=' + (CountOf "Duplicate scan")), "Same match + team scanned more than once: delete the extra row on the Scans sheet."),
    @("Matches without 6 scans", ('=IF(' + $hasSched + ',"–",' + (CountOf "Not 6 scans") + ')'), "Find the missing scout, or a duplicate. (Replaced by the checks below once the schedule is pasted.)"),
    @("Robots not scouted", ('=IF(' + $hasSched + ',' + (CountOf "Not scouted") + ',"needs schedule")'), "Ask that robot's scout to show the QR code again: Submitted matches in the app."),
    @("Team not in that match", ('=IF(' + $hasSched + ',' + (CountOf "Team not in that match") + ',"needs schedule")'), "Usually a typo in Team #: fix that cell on the Scans sheet."),
    @("Wrong alliance", ('=IF(' + $hasSched + ',' + (CountOf "Wrong alliance") + ',"needs schedule")'), "Fix the Alliance cell on the Scans sheet."),
    @("TBA schedule", ('=IF(' + $hasSched + ',COUNT(INDEX(''TBA Schedule''!$R$5#,,1))/6&" qual matches","not pasted yet")'), "See the TBA Schedule sheet.")
  )
  $check.Range("A4").Value2 = "Check"; $check.Range("B4").Value2 = "Found"; $check.Range("C4").Value2 = "What to do"
  HeaderStyle $check.Range("A4:G4")
  $check.Range("A13").Value2 = "PROBLEMS TO FIX"; $check.Range("A13").Font.Bold = $true; $check.Range("A13").Font.Size = 12
  $listHead = @("Problem", "Match #", "Team #", "Alliance", "Excel row", "Scout", "What it means")
  for ($i = 0; $i -lt $listHead.Count; $i++) { $check.Cells.Item(14, $i + 1).Value2 = $listHead[$i] }
  HeaderStyle $check.Range("A14:G14")
  $cw = @(24, 9, 9, 9, 10, 14, 70)
  for ($i = 0; $i -lt $cw.Count; $i++) { $check.Columns.Item($i + 1).ColumnWidth = $cw[$i] }
  $check.Range("B15:E3000").HorizontalAlignment = -4108
  $check.Range("A15").Formula2 = OneLine @'
=LET(ty,Scans[Match type],mt,Scans[Match '#],tm,Scans[Team '#],al,Scans[Alliance],sc,Scans[Scout name],rw,ROW(Scans[Team '#]),
nT,SEQUENCE(ROWS(tm)),
has,tm<>"",
qual,has*((ty="Qual")+(ty="")),
s,'TBA Schedule'!$R$5#,sm,INDEX(s,,1),sa,INDEX(s,,2),st,INDEX(s,,3),nS,SEQUENCE(ROWS(sm)),
hasSched,COUNT(sm)>0,
dupN,COUNTIFS(Scans[Team '#],tm,Scans[Match '#],mt,Scans[Match type],ty),
dup,FILTER(HSTACK(IF(nT,"Duplicate scan"),mt,tm,al,rw,sc,"Scanned "&dupN&" times: delete the extra row"),has*(dupN>1),""),
qm,UNIQUE(FILTER(mt,qual,"")),nQ,SEQUENCE(ROWS(qm)),
qc,COUNTIFS(Scans[Match '#],qm,Scans[Team '#],"<>",Scans[Match type],"Qual")+COUNTIFS(Scans[Match '#],qm,Scans[Team '#],"<>",Scans[Match type],""),
cnt,IF(hasSched,"",FILTER(HSTACK(IF(nQ,"Not 6 scans"),qm,IF(nQ,""),IF(nQ,""),IF(nQ,""),IF(nQ,""),qc&" scans in this match, expected 6"),(qm<>"")*(qc<>6),"")),
last,MAX(FILTER(mt,qual,0)),
seen,COUNTIFS(Scans[Match '#],sm,Scans[Team '#],st,Scans[Match type],"Qual")+COUNTIFS(Scans[Match '#],sm,Scans[Team '#],st,Scans[Match type],""),
miss,IF(hasSched,FILTER(HSTACK(IF(nS,"Not scouted"),sm,st,sa,IF(nS,""),IF(nS,""),IF(nS,"Nobody scanned this robot: ask its scout to show the QR code again (Submitted matches)")),ISNUMBER(sm)*(sm<=last)*(seen=0),""),""),
inM,COUNTIFS(sm,mt,st,tm),
wt,IF(hasSched,FILTER(HSTACK(IF(nT,"Team not in that match"),mt,tm,al,rw,sc,IF(COUNTIFS(sm,mt)>0,"Typo in Team #? Check the schedule for this match","There is no such qual match on the TBA schedule")),qual*(inM=0),""),""),
inA,COUNTIFS(sm,mt,st,tm,sa,al),
wa,IF(hasSched,FILTER(HSTACK(IF(nT,"Wrong alliance"),mt,tm,al,rw,sc,"The schedule says "&XLOOKUP(mt&"|"&tm,sm&"|"&st,sa,"?")),qual*(inM>0)*(inA=0),""),""),
all,VSTACK(dup,cnt,miss,wt,wa),
IFERROR(FILTER(all,IFERROR(INDEX(all,,1)<>"",FALSE)),"✓ Nothing to fix"))
'@
  # The counts read the list in A15, so they go in after it
  for ($i = 0; $i -lt $summary.Count; $i++) {
    $check.Cells.Item(5 + $i, 1).Value2 = $summary[$i][0]; $check.Cells.Item(5 + $i, 1).Font.Bold = $true
    $check.Cells.Item(5 + $i, 2).Formula2 = $summary[$i][1]; $check.Cells.Item(5 + $i, 2).HorizontalAlignment = -4108
    $check.Cells.Item(5 + $i, 3).Value2 = $summary[$i][2]
  }
  # Counts: red when above 0, green at 0 (cell-value rules, so no function names that differ by Excel language)
  $fc = $check.Range("B6:B10").FormatConditions.Add(1, 1, "=1", "=1E+307"); $fc.Interior.Color = RGB 255 199 206; $fc.Font.Color = RGB 156 0 6
  $fc = $check.Range("B6:B10").FormatConditions.Add(1, 3, "=0"); $fc.Interior.Color = RGB 198 239 206; $fc.Font.Color = RGB 0 97 0

  # ---------------- Instructions ----------------
  Title $info "DERT #2040 - REBUILT Scouting Workbook" "Import data from the DERT Scout web app (index.html) by scanning its QR codes." "J"
  $info.Columns.Item(1).ColumnWidth = 4; $info.Columns.Item(2).ColumnWidth = 110
  $lines = @(
    @("h","ONE-TIME SETUP (each laptop)"),
    @("1","Turn OFF Excel AutoComplete: File > Options > Advanced > uncheck 'Enable AutoComplete for cell values'."),
    @("", "   Otherwise Excel may 'finish' a scanned value for you (e.g. 'Depot' becomes 'Depot; Outpost') and the data will be wrong."),
    @("2","Use a USB/Bluetooth QR scanner in keyboard mode (most scanners default to this). Set its keyboard language to match Windows."),
    @("3","Test: open Notepad and scan a code - you should see values separated by tabs and the cursor should jump to a new line at the end."),
    @("h","IMPORTING MATCHES"),
    @("1","Go to the Scans sheet and click the first EMPTY cell in column A (Schema)."),
    @("2","Scan the student's QR code. The whole match fills one row and the cursor moves to the next row automatically."),
    @("3","Repeat for every QR code. Keep scanning - the table grows by itself."),
    @("4","Made a mistake? Delete the whole row (right-click row number > Delete) and scan again."),
    @("h","CHECKING THE DATA"),
    @("1","Once TBA has the match schedule (usually the day before quals), paste it on the TBA Schedule sheet - the steps are on that sheet."),
    @("2","Every few matches, open the Scouting Check sheet. Red numbers = something to fix; the list below them says what and in which Excel row."),
    @("h","SHEETS"),
    @("-","Scans - raw data, one row per robot per match. Do not change the column headers or their order."),
    @("-","Team Summary - averages per team (updates automatically). Pick list on the right is sorted by average total points."),
    @("-","Team Lookup - type a team number in the yellow cell to see every match and all observations for that team."),
    @("-","Scouting Check - duplicate scans, matches without 6 scans and, with the TBA schedule, unscouted robots, team-number typos and wrong alliances."),
    @("-","TBA Schedule - paste TBA's Flat Schedule here. Nothing else to type."),
    @("h","LEGEND"),
    @("-","Teal header = column names (must match the web app).   Yellow cell = type here.   Everything else is calculated or scanned - don't edit."),
    @("-","Points: fuel in active Hub = 1 pt. Tower Level 1 = 15 pts in auto / 10 pts in teleop, Level 2 = 20, Level 3 = 30 (2026 REBUILT Game Manual, Table 6-4)."),
    @("-","Ratings (Driver skill, Robot speed, Shooting accuracy) are 1 = poor to 5 = great. Multi-choice answers are separated by '; '."),
    @("-","Requires Microsoft 365 / Excel 2021 or newer (uses UNIQUE, SORTBY, FILTER, HSTACK)."),
    @("-","Schema (column A) is filled in by every QR code ($schema for this workbook). A RED row came from a different version"),
    @("", "   of the scouting form, so its answers are in the wrong columns: delete it, update the app on that phone (open it on Wi-Fi), scan again."),
    @("h","IF YOU CHANGE THE WEB FORM"),
    @("-","Adding, removing or renaming fields in index.html changes the QR columns and the Schema stamp. Update the header list in"),
    @("", "   build-excel.ps1 to match (the web app's QR screen has an 'Excel column headers' section you can copy) and rebuild this workbook,"),
    @("", "   otherwise every new scan shows up red. Check the Team Summary formulas too."),
    @("h","EXAMPLE ROW (what one scanned match looks like)")
  )
  $r = 4
  foreach ($l in $lines) {
    if ($l[0] -eq "h") {
      $r++; $info.Cells.Item($r, 1).Value2 = $l[1]; $info.Cells.Item($r, 1).Font.Bold = $true; $info.Cells.Item($r, 1).Font.Size = 12
      $info.Cells.Item($r, 1).Font.Color = RGB 14 143 139
    } else {
      $info.Cells.Item($r, 1).Value2 = $l[0]; $info.Cells.Item($r, 2).Value2 = $l[1]; $info.Cells.Item($r, 1).HorizontalAlignment = -4108
    }
    $r++
  }
  $example = @($schema,"Qual",12,2040,"Blue",2,"Alex",'Yes',"Center",8,8,2,"Left start line; Went to Neutral Zone","Depot","Level 1",
               47,9,"Neutral Zone floor; Depot","Close to Hub; Mid range","Bump","Collects / stockpiles","26–40","None","Some",
               "Level 3","5–10 s",4,4,4,"",0,0,"None","Yes","Fast shooter, very consistent climb",23,47,30,100,"10/6/2026, 3:45:12 PM")
  $exStart = $r
  # Example as a two-column list (header -> value) so it is readable
  for ($i = 0; $i -lt $n; $i++) {
    $info.Cells.Item($exStart + $i, 2).Value2 = "$($headers[$i]):  $($example[$i])"
    $info.Cells.Item($exStart + $i, 2).Font.Color = RGB 90 90 90
  }

  if ($Test) {
    # Fill test data to verify formulas
    $rows = @(
      @($schema,"Qual",1,2040,"Blue",1,"A","Yes","Left",8,8,2,"","Depot","Level 1",40,10,"Depot","Close to Hub","Bump","","26–40","None","No","Level 3","5–10 s",4,4,4,"",0,0,"None","Yes","",23,40,30,93,"x"),
      @($schema,"Qual",5,2040,"Blue",2,"B","Yes","Left",8,6,0,"","","None",30,10,"","","","","26–40","Some","No","Level 2","5–10 s",2,4,4,"Fuel jam",1,1,"None","Maybe","",6,30,20,56,"x"),
      @($schema,"Qual",2,111,"Red",3,"C","Yes","Left",8,2,2,"","","None",10,10,"","","","","1–10","None","No","None","",3,3,3,"",0,0,"None","No","",2,10,0,12,"x")
    )
    for ($ri = 0; $ri -lt $rows.Count; $ri++) {
      for ($ci = 0; $ci -lt $n; $ci++) {
        $v = $rows[$ri][$ci]; $cell = $scan.Cells.Item(2 + $ri, 1 + $ci)
        if ($v -is [string]) { if ($v -ne "") { $cell.Value2 = [string]$v } } else { $cell.Value2 = [double]$v }
      }
    }
    $look.Range("B3").Value2 = 2040
    $x.CalculateFull()
    "Summary rows:"
    for ($rr = 5; $rr -le 6; $rr++) { ($sum.Range("A$($rr):R$($rr)").Value2 | ForEach-Object { $_ }) -join " | " }
    "Pick list:"
    for ($rr = 5; $rr -le 6; $rr++) { ($sum.Range("T$($rr):AC$($rr)").Value2 | ForEach-Object { $_ }) -join " | " }
    "Lookup: " + $look.Range("C3").Text + " / " + $look.Range("D3").Text + " / Match # col C: " + $look.Range("C6").Text + ", " + $look.Range("C7").Text + " / A8: " + $look.Range("A8").Text
    "Table rows: " + $lo.ListRows.Count
    "Scouting Check: " + ((5..11 | ForEach-Object { $check.Range("A$_").Text + "=" + $check.Range("B$_").Text }) -join ", ")
    "  first problem: " + (($check.Range("A15:G15").Value2 | ForEach-Object { $_ }) -join " | ")
  }

  # Back to automatic so the saved workbook updates by itself
  $x.Calculation = -4105
  $x.CalculateFull()
  $info.Activate(); $info.Range("A1").Select()
  if (Test-Path $Out) { Remove-Item $Out }
  $wb.SaveAs($Out, 51)
  "Saved $Out"
  $wb.Close($false)
} catch {
  "ERROR at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
  $_.InvocationInfo.Line
} finally {
  try { $x.Quit() } catch {}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)
  [ExcelMessageFilter]::Revoke()
}
