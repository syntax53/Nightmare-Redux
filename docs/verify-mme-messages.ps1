# Verifies an MME export that contains the Messages table (Revision 2 schema) and the Titles and Abilities tables (NMR 1.8.4).
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
    if ($t -eq "Messages" -or $t -eq "Titles" -or $t -eq "Abilities") { continue }
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
  foreach ($k in $iOld.Keys) { if (-not $iNew.ContainsKey($k) -and $k -notlike "Messages.*" -and $k -notlike "Titles.*" -and $k -notlike "Abilities.*") { Fail "index missing: $k" } }
  # hashtable key order is arbitrary, so test each new table against the allowed set
  $newTabs = @($sNew.Keys | Where-Object { -not $sOld.ContainsKey($_) })
  $unexpected = @($newTabs | Where-Object { @("Messages","Titles","Abilities") -notcontains $_ })
  if ($unexpected.Count -eq 0) { Pass "no unexpected new tables (new: $($newTabs -join ','))" } else { Fail "unexpected new tables: $($unexpected -join ',')" }
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
Write-Output ("INFO  Chase rows: " + (Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Effects='Chase'"))
Write-Output ("INFO  CheckRoom rows: " + (Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE Effects='CheckRoom'"))

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
Write-Output "INFO  Self/Other rows whose text another row (a differently named spell or item) also has: $sameText"
if ($twins -eq 0) { Pass "no Self row is the you-form of a Cast/Other row" } else { Fail "$twins Self rows duplicate a {target} row with you/your" }
if ($noName -eq 0) { Pass "every row has a Name (the cleaned Msg text when no source name applies)" } else { Fail "$noName rows have an empty Name" }
$lists = @($allRows | Where-Object { $_.Name.Contains("|") }).Count
if ($lists -eq 0) { Pass "every Name is a single spell, item, monster or text (no lists)" } else { Fail "$lists rows carry a |-separated Name list" }
# every learnable spell aimed at the caster or a friend keeps its name on the row its own wear-off line belongs to
$friendly = Q $cnNew "SELECT Number,[Name] FROM Spells WHERE Learnable=1 AND Targets IN (0,1,2,6,13)"
$lost = 0; $checked = 0
foreach ($s in $friendly) {
  $sn = $s.Name; if ($sn -eq "") { continue }
  $own = @($allRows | Where-Object { $_.Kind -eq "Self" -and $_.EndMsg -like ("*effects of " + $sn + " wear*") })
  if ($own.Count -eq 0) { continue }
  $checked++
  $hit = $false; foreach ($r in $own) { if ($r.Name -eq $sn) { $hit = $true } }
  if (-not $hit) { $lost++; if ($lost -le 8) { Write-Output ("      spell lost its name: {0} -> row '{1}'" -f $sn, $own[0].Name) } }
}
if ($lost -eq 0) { Pass "learnable friendly spells keep their name on their own wear-off row ($checked checked)" } else { Fail "$lost of $checked learnable friendly spells are missing from their wear-off row's Name" }
$dirty = 0
foreach ($r in $allRows) { if ($r.Name.StartsWith("The {source} ") -or $r.Name.StartsWith("{source} ") -or $r.Name.StartsWith("The {source}'s ") -or $r.Name.StartsWith("{source}'s ") -or $r.Name.EndsWith("{dmg} damage!") -or $r.Name.EndsWith("{dmg} hit points!") -or $r.Name.Contains("{target} ") -or $r.Name.Contains("{target}'s ")) { $dirty++ } }
if ($dirty -eq 0) { Pass "text-derived names are cleaned of {source}/{target}/{dmg} boilerplate" } else { Fail "$dirty names still carry {source}/{target}/{dmg} boilerplate" }
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
# Cast rows are exempt: every player spell and casting item keeps its cast line, standard form or not
foreach ($r in $rows) { if ($r.Kind -eq "Cast") { continue }; foreach ($p in $rx) { if ($r.Text -cmatch $p) { $leak++; if ($leak -le 8) { Write-Output ("      standard-form row exported: [{0}] {1}" -f $r.Kind, $r.Text) }; break } } }
if ($leak -eq 0) { Pass "no standard-form lines exported" } else { Fail "$leak standard-form lines exported" }
$plain = 0
foreach ($r in $rows) { if ($r.Kind -ne "Cast" -and $r.Text -cmatch '^You (cast|invoke|sing) ' -and -not $r.Text.Contains("{dmg}")) { $plain++; if ($plain -le 5) { Write-Output ("      standard cast line exported: [{0}] {1}" -f $r.Kind, $r.Text) } } }
if ($plain -eq 0) { Pass "no standard 'You cast/invoke/sing ...' line without an amount exported" } else { Fail "$plain standard cast lines without an amount exported" }

# ---- 4. spot checks (stock data; a check fails harmlessly if your realm renamed the record)
function Expect($desc, $sql, $pred) {
  $rows = @(Q $cnNew $sql)
  $ok = $false; foreach ($r in $rows) { if (& $pred $r) { $ok = $true; break } }
  if ($ok) { Pass $desc } else { Fail ($desc + "  (got " + $rows.Count + " rows" + $(if($rows.Count){": " + (($rows | ForEach-Object { $_.Msg + " / " + $_.EndMsg + " / " + $_.Effects }) -join " || ")}) + ")") }
}
function ExpectNone($desc, $sql) { $c = Count $cnNew $sql; if ($c -eq 0) { Pass $desc } else { Fail "$desc  ($c rows)" } }
Expect "bless effect row named bless" "SELECT * FROM Messages WHERE [Msg]='You feel lucky!' AND Kind='Self'" { param($r) $r.EndMsg -eq "The effects of bless wear off!" -and $r.Name -eq "bless" }
Expect "resist cold effect row named resist cold" "SELECT * FROM Messages WHERE [Msg]='You feel resistant to cold!' AND Kind='Self'" { param($r) $r.EndMsg -eq "The effects of resist cold wear off!" -and $r.Name -eq "resist cold" }
ExpectNone "ice shield (monster-only, magery type) has no named row" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Name]='ice shield'"
Expect "acid rain effect row (hostile spell: Name = Msg)" "SELECT * FROM Messages WHERE [Msg]='You are covered in acid!' AND Kind='Self'" { param($r) ($r.Name -eq $r.Msg -or $r.Msg.EndsWith($r.Name + " for {dmg} damage!") -or $r.Msg.EndsWith($r.Name + " {dmg} damage!")) -and $r.EndMsg -eq "The acid covering you dries up." -and $r.Effects -match "DamageOverTime" }
Expect "poison Poisoned (Name = Msg)" "SELECT * FROM Messages WHERE [Msg]='You feel ill.' AND Kind='Self'" { param($r) ($r.Name -eq $r.Msg -or $r.Msg.EndsWith($r.Name + " for {dmg} damage!") -or $r.Msg.EndsWith($r.Name + " {dmg} damage!")) -and $r.Effects -match "Poisoned" }
Expect "magic missile Cast"      "SELECT * FROM Messages WHERE [Name]='magic missile' AND Kind='Cast'" { param($r) $r.Msg -eq "You fire a magic missile at {target} for {dmg} damage!" }
ExpectNone "magic missile incoming line is standard (contains 'you for {dmg} damage')" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='{source} fires a magic missile at you for {dmg} damage!'"
Expect "greater healing rain keeps its standard cast line (player spell)" "SELECT * FROM Messages WHERE Kind='Cast' AND [Name]='greater healing rain'" { param($r) $r.Msg -like "You cast greater healing rain on*" }
Expect "fumble once, Name = Msg"  "SELECT * FROM Messages WHERE [Msg]='You fumble in confusion!'"      { param($r) $r.Effects -eq "Fail" -and ($r.Name -eq $r.Msg -or $r.Msg.EndsWith($r.Name + " for {dmg} damage!") -or $r.Msg.EndsWith($r.Name + " {dmg} damage!")) }
$fumbles = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='You fumble in confusion!'"
if ($fumbles -eq 1) { Pass "fumble appears only once" } else { Fail "fumble appears $fumbles times" }
$flat = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='You are flat on your back!'"
if ($flat -eq 1) { Pass "knockdown status appears once (effects merged)" } else { Fail "knockdown status appears $flat times" }
# gaze of luck: one row per casting item/spell name (duplicates are covered by the duplicate-row check)
ExpectNone "lightning hit line is standard (contains 'you for {dmg} damage')" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='Lightning strikes you for {dmg} damage!'"
Expect "stunned merged Fail (Name = Msg)" "SELECT * FROM Messages WHERE Kind='Self' AND [Msg]='You are stunned!'" { param($r) ($r.Name -eq $r.Msg -or $r.Msg.EndsWith($r.Name + " for {dmg} damage!") -or $r.Msg.EndsWith($r.Name + " {dmg} damage!")) -and $r.Effects -match "Fail" -and $r.Effects -match "Confused" }
$stun = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='You are stunned!'"
if ($stun -le 2) { Pass "stunned text appears at most twice (one with an EndMsg, one without)" } else { Fail "stunned text appears $stun times" }
ExpectNone "dark force has no Cast row (monster-only)" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Name]='dark force' AND Kind='Cast'"
ExpectNone "dark force standard incoming line absent" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='You are struck by a dark force for {dmg} damage!'"
Expect "hold person Held (Name = Msg)" "SELECT * FROM Messages WHERE [Msg]='Your legs are paralyzed!' AND Kind='Self'" { param($r) ($r.Name -eq $r.Msg -or $r.Msg.EndsWith($r.Name + " for {dmg} damage!") -or $r.Msg.EndsWith($r.Name + " {dmg} damage!")) -and $r.Effects -match "Held" }
Expect "a death text row exists (Other, Death)" "SELECT TOP 1 * FROM Messages WHERE Kind='Other' AND Effects='Death'" { param($r) $r.Msg -ne "" }
ExpectNone "no death text for monsters worth 1 exp or less (practice dummy)" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Name]='practice dummy'"
ExpectNone "no death text for monsters under 101 exp (giant rat, 9 exp)" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg] LIKE 'The giant rat falls to the ground%'"
ExpectNone "lashworm hit on me is standard (contains 'you for {dmg} damage')" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg] LIKE '%darts forward and bites you for {dmg} damage!'"
Expect "filthbug hit on me (non-standard phrase)" "SELECT * FROM Messages WHERE Kind='Self' AND [Msg] LIKE '%claws you with its pincers%'" { param($r) $r.Msg -eq "The {source} claws you with its pincers for {dmg} damage!" -and $r.Name -eq "claws you with its pincers" -and $r.Effects -eq "Damage,Monster" }
ExpectNone "lashworm hit on other is standard shape" "SELECT COUNT(*) AS Cnt FROM Messages WHERE Kind='Other' AND [Msg] LIKE '%darts forward and bites {target} for {dmg} damage!'"
Expect "arrow trap row"          "SELECT * FROM Messages WHERE Effects='Damage,Trap' AND [Msg] LIKE 'An arrow shoots out of the wall%'" { param($r) $r.Kind -eq "Self" -and $r.Name -ne "" }
ExpectNone "you-form of a monster hit on another is dropped" "SELECT COUNT(*) AS Cnt FROM Messages WHERE Kind='Self' AND [Msg]='The {source} pounds you with its fist for {dmg} damage!'"
Expect "black wind cast line without an amount (Cast, named)" "SELECT * FROM Messages WHERE Kind='Cast' AND [Name]='black wind'" { param($r) $r.Msg -eq "A cool wind blows throughout the room!" }
Expect "annointed hands start line stands in for its standard cast line" "SELECT * FROM Messages WHERE Kind='Cast' AND [Name]='annointed hands'" { param($r) $r.Msg -eq "You lay hands on {target}." }
Expect "spiked shield damage shield (item, owner line)" "SELECT * FROM Messages WHERE Kind='Self' AND [Name]='spiked shield'" { param($r) $r.Msg -eq "The shield spike stabs {attacker} for {dmg} damage!" -and $r.Effects -eq "Damage,DamageShield" }
ExpectNone "item shield attacker line not exported (enigma gloves line 2)" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='A counterstrike at you does {dmg} damage!'"
Expect "desert damage (textblock-cast, player is caster)" "SELECT * FROM Messages WHERE Kind='Self' AND [Msg]='You suffer in the desert heat... you need water, soon!'" { param($r) $r.Effects -match "Damage" }
Expect "agony keeps its standard cast line" "SELECT * FROM Messages WHERE Kind='Cast' AND [Name]='agony'" { param($r) $r.Msg -eq "You cast agony on {target} for {dmg} damage!" }
Expect "bane dart incoming line (status message has no status text)" "SELECT * FROM Messages WHERE Kind='Self' AND [Msg]='{source} hurls a bane dart at you!'" { param($r) $r.Msg -ne "" }
ExpectNone "flay incoming line stays out (status and wear-off follow)" "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='{source} flays you viciously!'"
Expect "chase row (static)" "SELECT * FROM Messages WHERE [Msg]='{target} slips into the dark alley.'" { param($r) $r.Kind -eq "Other" -and $r.Effects -eq "Chase" }
Expect "monster entrance row (static)" "SELECT * FROM Messages WHERE [Msg]='A moss zombie rises from the ground!'" { param($r) $r.Kind -eq "Other" -and $r.Effects -eq "CheckRoom" }
$heads = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Messages WHERE [Msg]='A horde of severed heads descend upon you!'"
if ($heads -eq 1) { Pass "static rows de-duplicated (severed heads once)" } else { Fail "severed heads appears $heads times" }
ExpectNone "no swing verbs on non-weapons" "SELECT COUNT(*) AS Cnt FROM Items WHERE ItemType<>1 AND (Len(HitVerbs)>1 OR Len(HitVerbsOther)>1 OR Len(MissVerbs)>1)"
# WeaponType 0 is 1H Blunt (club, mace...), a weapon class like any other
$blunt = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Items WHERE ItemType=1 AND WeaponType=0 AND Len(HitVerbs)>1"
if ($blunt -gt 0) { Pass "1H blunt weapons carry swing verbs ($blunt)" } else { Fail "no 1H blunt weapon (ItemType 1, WeaponType 0) has swing verbs" }
Expect "longsword verbs"         "SELECT [Name], HitVerbs AS Msg, HitVerbsOther AS EndMsg, MissVerbs AS Effects FROM Items WHERE [Name]='longsword'" { param($r) $r.Msg -eq "slash|impale|hack" -and $r.EndMsg -eq "slashes|impales|hacks" -and $r.Effects -eq "swing at" }

# ---- 5. Titles table (one row per class: "<first>-<last>:<title>" runs from level 1, in level order)
if (-not $sNew.ContainsKey("Titles")) { Fail "Titles table missing" } else {
  $have = (@($sNew["Titles"].Keys) -join ",")
  if ($have -eq "Class,Titles") { Pass "Titles columns = Class,Titles" } else { Fail "Titles columns = $have" }
  if ($iNew.ContainsKey("Titles.pkTitles")) { Pass "pkTitles exists" } else { Fail "pkTitles missing" }
  $nT = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Titles"
  $nC = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Classes"
  if ($nT -eq $nC) { Pass "one Titles row per class ($nT)" } else { Fail "Titles has $nT rows, Classes has $nC" }
  $bad = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Titles WHERE [Class] NOT IN (SELECT [Number] FROM Classes)"
  if ($bad -eq 0) { Pass "every Titles row belongs to a class" } else { Fail "$bad Titles rows have no matching class" }
  $bad = 0; $empty = 0
  foreach ($r in (Q $cnNew "SELECT [Class], [Titles] FROM Titles")) {
    $t = [string]$r.Titles
    if ($t -eq "") { $empty++; continue }
    $ok = -not ($t.IndexOf([char]27) -ge 0); $next = 1
    # a "," inside a title belongs to the entry before it: split only where a new "<digits>-<digits>:" starts
    foreach ($piece in [regex]::Split($t, ',(?=\d+-\d+:)')) {
      if (-not $ok) { break }
      $m = [regex]::Match($piece, '^(\d+)-(\d+):(.*)$')
      if (-not $m.Success) { $ok = $false; break }
      $a = [int]$m.Groups[1].Value; $b = [int]$m.Groups[2].Value
      if ($a -ne $next -or $b -lt $a) { $ok = $false; break }
      $next = $b + 1
    }
    if (-not $ok) { $bad++; if ($bad -le 5) { Write-Output ("      malformed Titles for class {0}: {1}" -f $r.Class, $t) } }
  }
  if ($bad -eq 0) { Pass "every Titles string is contiguous level ranges from 1 with no colour codes" } else { Fail "$bad Titles strings are malformed" }
  if ($empty -gt 0) { Write-Output "INFO  classes without a title textblock: $empty" }
  Expect "class 1 titles start at level 1 (stock: Apprentice)" "SELECT [Class], [Titles] AS Msg FROM Titles WHERE [Class]=1" { param($r) $r.Msg -like "1-1:Apprentice,*" }
}

# ---- 6. Abilities table (NMR's loaded ability list: one row per ability number)
if (-not $sNew.ContainsKey("Abilities")) { Fail "Abilities table missing" } else {
  $have = (@($sNew["Abilities"].Keys) -join ",")
  if ($have -eq "Number,Name") { Pass "Abilities columns = Number,Name" } else { Fail "Abilities columns = $have" }
  if ($iNew.ContainsKey("Abilities.pkAbilities")) { Pass "pkAbilities exists" } else { Fail "pkAbilities missing" }
  $nA = Count $cnNew "SELECT COUNT(*) AS Cnt FROM Abilities"
  if ($nA -gt 0) { Pass "Abilities has $nA rows" } else { Fail "Abilities is empty" }
  ExpectNone "no blank ability names" "SELECT COUNT(*) AS Cnt FROM Abilities WHERE Len([Name] & '')=0 OR [Name]=Chr(0)"
  Expect "ability 1 is named (stock: Damage)" "SELECT [Number], [Name] AS Msg FROM Abilities WHERE [Number]=1" { param($r) $r.Msg -ne "" }
  # the game data can use ability numbers NMR's list does not name; report them, they are not an export fault
  $unnamed = @()
  foreach ($t in "Items","Monsters","Spells","Classes","Races") {
    $n = if ($t -eq "Items") { 19 } else { 9 }
    for ($x = 0; $x -le $n; $x++) {
      foreach ($r in (Q $cnNew "SELECT DISTINCT [Abil-$x] AS A FROM [$t] WHERE [Abil-$x]>0 AND [Abil-$x] NOT IN (SELECT [Number] FROM Abilities)")) { if ($unnamed -notcontains $r.A) { $unnamed += $r.A } }
    }
  }
  if ($unnamed.Count -eq 0) { Pass "every Abil-# value in the export has an Abilities row" } else { Write-Output ("INFO  ability numbers used but not in the list: " + (($unnamed | Sort-Object) -join ",")) }
}
$cnNew.Close()
Write-Output ""
if ($fail -eq 0) { Write-Output "ALL CHECKS PASSED" } else { Write-Output "$fail CHECK(S) FAILED" }
