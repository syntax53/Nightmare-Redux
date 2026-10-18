# Verifies an MME export that contains the Messages table (Revision 2 schema).
# Must run under 32-bit PowerShell (Jet 4.0 is 32-bit only):
#   C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File docs\verify-mme-messages.ps1 -New "C:\path\new-export.mdb" [-Old "C:\path\previous-version-export.mdb"]
# Jet SQL through ADO uses % and _ as LIKE wildcards; reserved names ([Name], [Msg]) are bracketed.
param(
  [Parameter(Mandatory=$true)][string]$New,
  [string]$Old = ""
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
function Q($cn,$sql){ $rs = New-Object -ComObject ADODB.Recordset; try { $rs.Open($sql,$cn) } catch { Write-Output ("ERROR running: " + $sql + " :: " + $_.Exception.Message); return @() }; $list=@(); while(-not $rs.EOF){ $h=@{}; for($i=0;$i -lt $rs.Fields.Count;$i++){ $v=$rs.Fields.Item($i).Value; if($v -is [string]){ $v=$v.TrimEnd([char]0).Trim() }; $h[$rs.Fields.Item($i).Name]=$v }; $list += ,$h; $rs.MoveNext() }; $rs.Close(); return $list }
function Count($cn,$sql){ $r = @(Q $cn $sql); if ($r.Count -eq 0) { return -1 }; return ($r[0].Values | Select-Object -First 1) }

$cnNew = OpenDb $New
$sNew = Schema $cnNew
$iNew = Indexes $cnNew

# ---- 1. schema diff against a previous-version export (optional)
if ($Old -ne "" -and (Test-Path $Old)) {
  $cnOld = OpenDb $Old
  $sOld = Schema $cnOld
  $iOld = Indexes $cnOld
  foreach ($t in $sOld.Keys) {
    if ($t -eq "Messages") { continue }
    if (-not $sNew.ContainsKey($t)) { Fail "table missing in new export: $t"; continue }
    $oldCols = @($sOld[$t].Keys); $newCols = @($sNew[$t].Keys)
    for ($i=0; $i -lt $oldCols.Count; $i++) {
      if ($i -ge $newCols.Count -or $newCols[$i] -ne $oldCols[$i]) { Fail "$t column order/name differs at position $i (old '$($oldCols[$i])' new '$($newCols[$i])')"; break }
      if ($sOld[$t][$oldCols[$i]] -ne $sNew[$t][$oldCols[$i]]) { Fail "$t.$($oldCols[$i]) type changed $($sOld[$t][$oldCols[$i]]) -> $($sNew[$t][$oldCols[$i]])" }
    }
    $extra = $newCols | Where-Object { $oldCols -notcontains $_ }
    if ($t -eq "Items") {
      if (($extra -join ",") -eq "HitVerbs,HitVerbsOther,MissVerbs" -or $extra.Count -eq 0) { Pass "Items columns: only the verb columns are new" } else { Fail "Items extra columns: '$($extra -join ',')'" }
    } elseif ($extra.Count -gt 0) { Fail "$t has unexpected extra columns: $($extra -join ',')" }
  }
  foreach ($k in $iOld.Keys) { if (-not $iNew.ContainsKey($k) -and $k -notlike "Messages.*") { Fail "index missing: $k" } }
  $newTabs = $sNew.Keys | Where-Object { -not $sOld.ContainsKey($_) }
  if ($newTabs.Count -eq 0 -or ($newTabs -join ",") -eq "Messages") { Pass "no unexpected new tables" } else { Fail "new tables: $($newTabs -join ',')" }
  $cnOld.Close()
} else { Write-Output "SKIP  no previous export given for the schema diff" }

# ---- 2. Messages table shape
if (-not $sNew.ContainsKey("Messages")) { Fail "Messages table missing"; $cnNew.Close(); exit 1 }
$want = "ID,Name,Kind,Msg,EndMsg,Effects"
$have = (@($sNew["Messages"].Keys) -join ",")
if ($have -eq $want) { Pass "Messages columns = $want" } else { Fail "Messages columns = $have" }
if ($iNew.ContainsKey("Messages.pkMessages")) { Pass "pkMessages exists" } else { Fail "pkMessages missing" }
$n = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages"
Write-Output "INFO  Messages rows: $n"
foreach ($k in "Cast","Self","Other") { $c = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Kind='$k'"; Write-Output ("INFO  Kind {0,-6} {1}" -f $k, $c) }
Write-Output ("INFO  Self rows with EndMsg (tracked effects): " + (Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Kind='Self' AND Len([EndMsg])>1"))
Write-Output ("INFO  Fail rows: " + (Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Effects LIKE '%Fail%'"))
Write-Output ("INFO  Monster hit rows: " + (Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Effects LIKE '%Monster%'"))
Write-Output ("INFO  Death rows: " + (Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Effects LIKE '%Death%'"))
Write-Output ("INFO  Trap rows: " + (Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Effects LIKE '%Trap%'"))

# ---- 3. negative checks
$bad = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg] LIKE '%[%]s%' OR [Msg] LIKE '%[%]d%' OR [EndMsg] LIKE '%[%]s%' OR [EndMsg] LIKE '%[%]d%'"
if ($bad -eq 0) { Pass "no raw %s/%d left" } else { Fail "$bad rows still contain raw %s/%d" }
$bad = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Kind NOT IN ('Cast','Self','Other')"
if ($bad -eq 0) { Pass "all Kind values known" } else { Fail "$bad rows with unknown Kind" }
$bad = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg] LIKE '%{caster}%' OR [Msg] LIKE '%{monster}%' OR [Msg] LIKE '%{spell}%'"
if ($bad -eq 0) { Pass "no retired tokens ({caster}/{monster}/{spell})" } else { Fail "$bad rows use retired tokens" }
$bad = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Effects LIKE '%Timed%'"
if ($bad -eq 0) { Pass "no Timed effect" } else { Fail "$bad rows carry Timed" }
$bad = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Kind='Cast' AND (Len([Name])<=1)"
if ($bad -eq 0) { Pass "every Cast row is named" } else { Fail "$bad Cast rows have no Name" }
$bad = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg] LIKE '%' + Chr(27) + '%' OR [EndMsg] LIKE '%' + Chr(27) + '%'"
if ($bad -eq 0) { Pass "no ANSI escape codes" } else { Fail "$bad rows contain ANSI escape codes" }
$q = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg] LIKE '%{?}%' OR [EndMsg] LIKE '%{?}%'"
Write-Output "INFO  rows with unresolved {?} placeholder: $q (should match the warning log)"
$allRows = Q $cnNew "SELECT Kind,[Name],[Msg],[EndMsg],Effects FROM Messages"
$seen = @{}; $dup = 0; $sameText = 0; $seenText = @{}; $youForms = New-Object System.Collections.Hashtable; $twins = 0; $noName = 0
foreach ($r in $allRows) {
  $k = $r.Kind + "|" + $r.Name + "|" + $r.Msg + "|" + $r.EndMsg
  if ($seen.ContainsKey($k)) { $dup++ } else { $seen[$k] = 1 }
  if ($r.Kind -ne "Cast") { $t = $r.Kind + "|" + $r.Msg + "|" + $r.EndMsg; if ($seenText.ContainsKey($t)) { $sameText++ } else { $seenText[$t] = 1 } }
  if ($r.Kind -ne "Self" -and $r.EndMsg -eq "" -and $r.Msg.Contains("{target}")) { $youForms[$r.Msg.Replace("{target}'s","your").Replace("{target}","you")] = 1 }
  if ($r.Name -eq "") { $noName++ }
}
foreach ($r in $allRows) { if ($r.Kind -eq "Self" -and $r.EndMsg -eq "" -and $youForms.ContainsKey($r.Msg)) { $twins++ } }
if ($dup -eq 0) { Pass "no duplicate rows" } else { Fail "$dup duplicate rows" }
if ($sameText -eq 0) { Pass "Self/Other rows unique by kind, text and end text (effects merged)" } else { Fail "$sameText Self/Other rows repeat another row's text" }
if ($twins -eq 0) { Pass "no Self row is the you-form of a Cast/Other row" } else { Fail "$twins Self rows duplicate a {target} row with you/your" }
if ($noName -eq 0) { Pass "every row has a Name (the Msg text when no source name applies)" } else { Fail "$noName rows have an empty Name" }
# standard-form lines that should have been excluded
$rows = Q $cnNew "SELECT Kind, [Msg] AS Text, Effects FROM Messages"
$rx = @(
  '^You cast .+ (on|at|upon) .*\{dmg\} (damage|healing|hit points)!?$',
  '^\{source\} casts .+ (on|at|upon) you.*\{dmg\} damage!?$',
  '^You are .* for \{dmg\} damage!?$',
  '^You take {dmg} damage',
  'you for {dmg} damage',
  '^{target} .*{dmg}.* damage',
  '^{attacker} .*{dmg}.* damage',
  '^(The |An? )?\{source\} [a-z-]+ you for \{dmg\} damage!$',
  '{target} for {dmg} damage'
)
$leak = 0
foreach ($r in $rows) { foreach ($p in $rx) { if ($r.Text -cmatch $p) { $leak++; if ($leak -le 8) { Write-Output ("      standard-form row exported: [{0}] {1}" -f $r.Kind, $r.Text) }; break } } }
if ($leak -eq 0) { Pass "no standard-form lines exported" } else { Fail "$leak standard-form lines exported" }

# ---- 4. spot checks (stock data; a check fails harmlessly if your realm renamed the record)
function Expect($desc, $sql, $pred) {
  $rows = Q $cnNew $sql
  $ok = $false; foreach ($r in $rows) { if (& $pred $r) { $ok = $true; break } }
  if ($ok) { Pass $desc } else { Fail ($desc + "  (got " + $rows.Count + " rows" + $(if($rows.Count){": " + (($rows | ForEach-Object { $_.Msg + " / " + $_.EndMsg + " / " + $_.Effects }) -join " || ")}) + ")") }
}
function ExpectNone($desc, $sql) { $c = Count $cnNew $sql; if ($c -eq 0) { Pass $desc } else { Fail "$desc  ($c rows)" } }
Expect "bless effect row (shared by several sources, so Name = Msg)" "SELECT * FROM Messages WHERE [Msg]='You feel lucky!' AND Kind='Self'" { param($r) $r.EndMsg -eq "The effects of bless wear off!" -and ($r.Name -eq "bless" -or $r.Name -eq $r.Msg) }
Expect "acid rain effect row (hostile spell: Name = Msg)" "SELECT * FROM Messages WHERE [Msg]='You are covered in acid!' AND Kind='Self'" { param($r) $r.Name -eq $r.Msg -and $r.EndMsg -eq "The acid covering you dries up." -and $r.Effects -match "DamageOverTime" }
Expect "poison Poisoned (Name = Msg)" "SELECT * FROM Messages WHERE [Msg]='You feel ill.' AND Kind='Self'" { param($r) $r.Name -eq $r.Msg -and $r.Effects -match "Poisoned" }
Expect "magic missile Cast"      "SELECT * FROM Messages WHERE [Name]='magic missile' AND Kind='Cast'" { param($r) $r.Msg -eq "You fire a magic missile at {target} for {dmg} damage!" }
ExpectNone "magic missile incoming line is standard (contains 'you for {dmg} damage')" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='{source} fires a magic missile at you for {dmg} damage!'"
ExpectNone "greater healing rain cast line is standard (You cast X on ..., healing N damage!)" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Name]='greater healing rain'"
Expect "fumble once, Name = Msg"  "SELECT * FROM Messages WHERE [Msg]='You fumble in confusion!'"      { param($r) $r.Effects -eq "Fail" -and $r.Name -eq $r.Msg }
$fumbles = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='You fumble in confusion!'"
if ($fumbles -eq 1) { Pass "fumble appears only once" } else { Fail "fumble appears $fumbles times" }
$flat = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='You are flat on your back!'"
if ($flat -eq 1) { Pass "knockdown status appears once (effects merged)" } else { Fail "knockdown status appears $flat times" }
$gaze = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='The gaze of luck is upon you.'"
if ($gaze -le 1) { Pass "gaze of luck variants merged into one row" } else { Fail "gaze of luck appears $gaze times" }
ExpectNone "lightning hit line is standard (contains 'you for {dmg} damage')" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='Lightning strikes you for {dmg} damage!'"
Expect "stunned merged Fail (Name = Msg)" "SELECT * FROM Messages WHERE Kind='Self' AND [Msg]='You are stunned!'" { param($r) $r.Name -eq $r.Msg -and $r.Effects -match "Fail" -and $r.Effects -match "Confused" }
$stun = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='You are stunned!'"
if ($stun -le 2) { Pass "stunned text appears at most twice (one with an EndMsg, one without)" } else { Fail "stunned text appears $stun times" }
ExpectNone "dark force has no Cast row (monster-only)" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Name]='dark force' AND Kind='Cast'"
ExpectNone "dark force standard incoming line absent" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='You are struck by a dark force for {dmg} damage!'"
Expect "hold person Held (Name = Msg)" "SELECT * FROM Messages WHERE [Msg]='Your legs are paralyzed!' AND Kind='Self'" { param($r) $r.Name -eq $r.Msg -and $r.Effects -match "Held" }
Expect "giant rat death text (Other, Death)" "SELECT * FROM Messages WHERE Kind='Other' AND [Msg] LIKE 'The giant rat falls to the ground%'" { param($r) $r.Effects -eq "Death" -and ($r.Name -eq "giant rat" -or $r.Name -eq $r.Msg) }
ExpectNone "no death text for monsters worth 1 exp or less (practice dummy)" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Name]='practice dummy'"
ExpectNone "lashworm hit on me is standard (contains 'you for {dmg} damage')" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg] LIKE '%darts forward and bites you for {dmg} damage!'"
Expect "filthbug hit on me (non-standard phrase)" "SELECT * FROM Messages WHERE Kind='Self' AND [Msg] LIKE '%claws you with its pincers%'" { param($r) $r.Msg -eq "The {source} claws you with its pincers for {dmg} damage!" -and $r.Effects -eq "Damage,Monster" }
ExpectNone "lashworm hit on other is standard shape" "SELECT COUNT(*) AS Cnt FROM Messages WHERE Kind='Other' AND [Msg] LIKE '%darts forward and bites {target} for {dmg} damage!'"
Expect "arrow trap row"          "SELECT * FROM Messages WHERE Effects='Damage,Trap' AND [Msg] LIKE 'An arrow shoots out of the wall%'" { param($r) $r.Kind -eq "Self" -and $r.Name -eq $r.Msg }
ExpectNone "you-form of a monster hit on another is dropped" "SELECT COUNT(*) AS Cnt FROM Messages WHERE Kind='Self' AND [Msg]='The {source} pounds you with its fist for {dmg} damage!'"
Expect "longsword verbs"         "SELECT [Name], HitVerbs AS Msg, HitVerbsOther AS EndMsg, MissVerbs AS Effects FROM Items WHERE [Name]='longsword'" { param($r) $r.Msg -eq "slash|impale|hack" -and $r.EndMsg -eq "slashes|impales|hacks" -and $r.Effects -eq "swing at" }
$cnNew.Close()
Write-Output ""
if ($fail -eq 0) { Write-Output "ALL CHECKS PASSED" } else { Write-Output "$fail CHECK(S) FAILED" }
