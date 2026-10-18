# Verifies an MME export that contains the Messages table.
# Must run under 32-bit PowerShell (Jet 4.0 is 32-bit only):
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File docs\verify-mme-messages.ps1 -New "C:\path\new-export.mdb" -Old "G:\My Drive\mud\dats\data-Paradigm-1.9.1x.mdb"
param(
  [Parameter(Mandatory=$true)][string]$New,
  [string]$Old = "G:\My Drive\mud\dats\data-Paradigm-1.9.1x.mdb"
)
$ErrorActionPreference = "Continue"
$fail = 0
function Pass($m){ Write-Output ("PASS  " + $m) }
function Fail($m){ $script:fail++; Write-Output ("FAIL  " + $m) }

function OpenDb($path){ $cn = New-Object -ComObject ADODB.Connection; $cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$path"); return $cn }
function Schema($cn){
  $out = @{}
  $tabs = $cn.OpenSchema(20)
  while (-not $tabs.EOF) {
    if ($tabs.Fields.Item("TABLE_TYPE").Value -eq "TABLE") {
      $t = $tabs.Fields.Item("TABLE_NAME").Value
      $rs = New-Object -ComObject ADODB.Recordset
      $rs.Open("SELECT TOP 1 * FROM [$t]", $cn)
      $cols = New-Object System.Collections.Specialized.OrderedDictionary
      for ($i=0;$i -lt $rs.Fields.Count;$i++){ $cols[$rs.Fields.Item($i).Name] = $rs.Fields.Item($i).Type }
      $rs.Close()
      $out[$t] = $cols
    }
    $tabs.MoveNext()
  }
  return $out
}
function Indexes($cn){
  $out = @{}
  $ix = $cn.OpenSchema(12)
  while (-not $ix.EOF) { $k = $ix.Fields.Item("TABLE_NAME").Value + "." + $ix.Fields.Item("INDEX_NAME").Value; $out[$k] = $ix.Fields.Item("COLUMN_NAME").Value; $ix.MoveNext() }
  return $out
}
function Q($cn,$sql){ $rs = New-Object -ComObject ADODB.Recordset; $rs.Open($sql,$cn); $list=@(); while(-not $rs.EOF){ $h=@{}; for($i=0;$i -lt $rs.Fields.Count;$i++){ $v=$rs.Fields.Item($i).Value; if($v -is [string]){ $v=$v.TrimEnd([char]0).Trim() }; $h[$rs.Fields.Item($i).Name]=$v }; $list += ,$h; $rs.MoveNext() }; $rs.Close(); return $list }
function Count($cn,$sql){ (Q $cn $sql)[0].Values | Select-Object -First 1 }

$cnNew = OpenDb $New
$sNew = Schema $cnNew
$iNew = Indexes $cnNew

# ---- 1. schema diff against the previous-version export
if (Test-Path $Old) {
  $cnOld = OpenDb $Old
  $sOld = Schema $cnOld
  $iOld = Indexes $cnOld
  foreach ($t in $sOld.Keys) {
    if (-not $sNew.ContainsKey($t)) { Fail "table missing in new export: $t"; continue }
    $oldCols = @($sOld[$t].Keys); $newCols = @($sNew[$t].Keys)
    for ($i=0; $i -lt $oldCols.Count; $i++) {
      if ($i -ge $newCols.Count -or $newCols[$i] -ne $oldCols[$i]) { Fail "$t column order/name differs at position $i (old '$($oldCols[$i])' new '$($newCols[$i])')"; break }
      if ($sOld[$t][$oldCols[$i]] -ne $sNew[$t][$oldCols[$i]]) { Fail "$t.$($oldCols[$i]) type changed $($sOld[$t][$oldCols[$i]]) -> $($sNew[$t][$oldCols[$i]])" }
    }
    $extra = $newCols | Where-Object { $oldCols -notcontains $_ }
    if ($t -eq "Items") {
      if (($extra -join ",") -eq "HitVerbs,HitVerbsOther,MissVerbs") { Pass "Items has exactly the three new verb columns appended" } else { Fail "Items extra columns: '$($extra -join ',')'" }
    } elseif ($extra.Count -gt 0) { Fail "$t has unexpected extra columns: $($extra -join ',')" }
  }
  foreach ($k in $iOld.Keys) { if (-not $iNew.ContainsKey($k)) { Fail "index missing: $k" } }
  $newTabs = $sNew.Keys | Where-Object { -not $sOld.ContainsKey($_) }
  if (($newTabs -join ",") -eq "Messages") { Pass "only new table is Messages" } else { Fail "new tables: $($newTabs -join ',')" }
  $cnOld.Close()
  if ($fail -eq 0) { Pass "all pre-existing tables, columns, types and indexes unchanged" }
} else { Write-Output "SKIP  old export not found: $Old" }

# ---- 2. Messages table shape
if (-not $sNew.ContainsKey("Messages")) { Fail "Messages table missing"; $cnNew.Close(); exit 1 }
$want = "ID,Name,Kind,Text,EndText,Effects,Duration"
$have = (@($sNew["Messages"].Keys) -join ",")
if ($have -eq $want) { Pass "Messages columns = $want" } else { Fail "Messages columns = $have" }
if ($iNew.ContainsKey("Messages.pkMessages")) { Pass "pkMessages exists" } else { Fail "pkMessages missing" }
$n = Count $cnNew "SELECT COUNT(*) AS N FROM Messages"
Write-Output "INFO  Messages rows: $n"
foreach ($k in "Status","Target","Self","Fail","Shield","ShieldOwner","Hit","HitOther","Death","Trap","TrapFail") { $c = Count $cnNew "SELECT COUNT(*) AS N FROM Messages WHERE Kind='$k'"; Write-Output ("INFO  Kind {0,-12} {1}" -f $k, $c) }

# ---- 3. negative checks
$bad = Count $cnNew "SELECT COUNT(*) AS N FROM Messages WHERE Text LIKE '*%s*' OR Text LIKE '*%d*' OR EndText LIKE '*%s*' OR EndText LIKE '*%d*'"
if ($bad -eq 0) { Pass "no raw %s/%d left" } else { Fail "$bad rows still contain raw %s/%d" }
$bad = Count $cnNew "SELECT COUNT(*) AS N FROM Messages WHERE Kind NOT IN ('Status','Target','Self','Fail','Shield','ShieldOwner','Hit','HitOther','Death','Trap','TrapFail')"
if ($bad -eq 0) { Pass "all Kind values known" } else { Fail "$bad rows with unknown Kind" }
$bad = Count $cnNew "SELECT COUNT(*) AS N FROM Messages WHERE Kind IN ('Status','Target','Self','Fail','Death') AND (Name IS NULL OR Len(Name)<=1)"
if ($bad -eq 0) { Pass "names present on named kinds" } else { Fail "$bad named-kind rows have an empty Name" }
$bad = Count $cnNew "SELECT COUNT(*) AS N FROM Messages WHERE Text LIKE '*{?}*'"
Write-Output "INFO  rows with unresolved {?} placeholder: $bad (should match the warning log)"
$dup = Count $cnNew "SELECT COUNT(*) AS N FROM (SELECT Kind,Name,Text,EndText,Effects FROM Messages GROUP BY Kind,Name,Text,EndText,Effects HAVING COUNT(*)>1)"
if ($dup -eq 0) { Pass "no duplicate rows" } else { Fail "$dup duplicate key groups" }
# standard-form lines that should have been excluded
$rows = Q $cnNew "SELECT Kind, Text FROM Messages WHERE Kind IN ('Self','Target','Hit','HitOther')"
$rx = @{
  Self   = '^You cast \{spell\} (on|at|upon) \{target\}(,| for| and)? .*\{dmg\} (damage|healing|hit points)!?$'
  Target = '^\{caster\} casts \{spell\} (on|at|upon) you(,| for| and)? .*\{dmg\} damage!?$'
  Hit    = '^(The |An? )?\{monster\} [a-z-]+ you for \{dmg\} damage!$'
  HitOther = '^(The |An? )?\{monster\} [a-z-]+ \{target\} for \{dmg\} damage!$'
}
$leak = 0
foreach ($r in $rows) { if ($r.Text -cmatch $rx[$r.Kind]) { $leak++; if ($leak -le 5) { Write-Output ("      standard-form row exported: [{0}] {1}" -f $r.Kind, $r.Text) } } }
if ($leak -eq 0) { Pass "no standard-form lines exported" } else { Fail "$leak standard-form lines exported" }

# ---- 4. spot checks (stock data; skip any that your realm renamed)
function Expect($desc, $sql, $pred) {
  $rows = Q $cnNew $sql
  $ok = $false; foreach ($r in $rows) { if (& $pred $r) { $ok = $true; break } }
  if ($ok) { Pass $desc } else { Fail ($desc + "  (got " + $rows.Count + " rows" + $(if($rows.Count){": " + (($rows | ForEach-Object { $_.Text + " / " + $_.EndText + " / " + $_.Effects }) -join " || ")}) + ")") }
}
Expect "bless Status row"      "SELECT * FROM Messages WHERE Name='bless' AND Kind='Status'"      { param($r) $r.Text -eq "You feel lucky!" -and $r.EndText -eq "The effects of bless wear off!" -and $r.Effects -match "Timed" }
Expect "acid rain Status row"  "SELECT * FROM Messages WHERE Name='acid rain' AND Kind='Status'"  { param($r) $r.Text -eq "You are covered in acid!" -and $r.EndText -eq "The acid covering you dries up." -and $r.Effects -match "DamageOverTime" -and $r.Effects -match "Timed" }
Expect "poison cloud Poisoned" "SELECT * FROM Messages WHERE Name='poison cloud' AND Kind='Status'" { param($r) $r.Text -eq "You feel ill." -and $r.Effects -match "Poisoned" }
Expect "bites Poisoned"        "SELECT * FROM Messages WHERE Name='bites' AND Kind='Status'"      { param($r) $r.Effects -match "Poisoned" }
Expect "magic missile Self"    "SELECT * FROM Messages WHERE Name='magic missile' AND Kind='Self'" { param($r) $r.Text -eq "You fire a {spell} at {target} for {dmg} damage!" }
Expect "magic missile Target"  "SELECT * FROM Messages WHERE Name='magic missile' AND Kind='Target'" { param($r) $r.Text -eq "{caster} fires a {spell} at you for {dmg} damage!" }
Expect "confusion Fail"        "SELECT * FROM Messages WHERE Name='confusion' AND Kind='Fail'"    { param($r) $r.Text -eq "You fumble in confusion!" }
Expect "hold person Status"    "SELECT * FROM Messages WHERE Name='hold person' AND Kind='Status'" { param($r) $r.Text -eq "Your legs are paralyzed!" -and $r.Effects -match "Held" }
Expect "giant rat Death"       "SELECT * FROM Messages WHERE Name='giant rat' AND Kind='Death'"   { param($r) $r.Text -like "The giant rat falls to the ground*" }
Expect "lashworm Hit"          "SELECT * FROM Messages WHERE Kind='Hit' AND Text LIKE '*darts forward and bites you*'" { param($r) $r.Text -eq "{monster} darts forward and bites you for {dmg} damage!" }
Expect "lashworm HitOther"     "SELECT * FROM Messages WHERE Kind='HitOther' AND Text LIKE '*darts forward and bites*'" { param($r) $r.Text -eq "{monster} darts forward and bites {target} for {dmg} damage!" }
Expect "arrow trap row"        "SELECT * FROM Messages WHERE Kind IN ('Trap','TrapFail') AND Text LIKE 'An arrow shoots out of the wall*'" { param($r) $r.Effects -eq "Damage,Trap" }
Expect "longsword verbs"       "SELECT Name, HitVerbs AS Text, HitVerbsOther AS EndText, MissVerbs AS Effects FROM Items WHERE Name='longsword'" { param($r) $r.Text -eq "slash|impale|hack" -and $r.EndText -eq "slashes|impales|hacks" -and $r.Effects -eq "swing at" }
$cnNew.Close()
Write-Output ""
if ($fail -eq 0) { Write-Output "ALL CHECKS PASSED" } else { Write-Output "$fail CHECK(S) FAILED" }
