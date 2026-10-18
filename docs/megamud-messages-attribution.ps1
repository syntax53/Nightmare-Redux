$f="P:\MajorMUD\_Euphoria\NMR-EuphStockishDataExport\NMR-EuphStockishDataExport.mdb"
$md="C:\Users\synta\GitHub\MegaMud2\MegaMudSource\Debug\Default\MESSAGES.md"
$out="C:\Users\synta\AppData\Local\Temp\claude\C--Users-synta-GitHub-Nightmare-Redux\af1c8898-08bc-4273-9327-abc1358799a3\scratchpad\attribution2.txt"

$cn = New-Object -ComObject ADODB.Connection
$cn.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$f")
function Rows($sql){ $rs = New-Object -ComObject ADODB.Recordset; $rs.Open($sql,$cn); $list = New-Object System.Collections.ArrayList; while(-not $rs.EOF){ $h=@{}; for($i=0;$i -lt $rs.Fields.Count;$i++){ $v=$rs.Fields.Item($i).Value; if($v -is [string]){$v=$v.Trim()}; $h[$rs.Fields.Item($i).Name]=$v }; [void]$list.Add($h); $rs.MoveNext() }; $rs.Close(); return $list }

# messages
$msgs = @{}
foreach ($r in (Rows "SELECT Number,[Line 1],[Line 2],[Line 3] FROM Messages")) { $msgs[[int]$r["Number"]] = @($r["Line 1"],$r["Line 2"],$r["Line 3"]) }

# reverse refs
$refs = @{}
function AddRef($n,$s){ if($n -le 1){return}; if(-not $refs.ContainsKey($n)){ $refs[$n]=New-Object System.Collections.ArrayList }; if($refs[$n].Count -lt 6){ [void]$refs[$n].Add($s) } }
$spellNames=@{}
foreach ($s in (Rows "SELECT * FROM Spells")) {
  $n=[int]$s["Number"]; $spellNames[$n]=$s["Name"]
  AddRef ([int]$s["Cast MSG A"]) "Spell#$n($($s["Name"])).CastA"
  AddRef ([int]$s["Cast MSG B"]) "Spell#$n($($s["Name"])).CastB"
  for($i=0;$i -le 9;$i++){ $a=[int]$s["Ability $i"]; $v=[int]$s["Ability Value $i"]; if($a -eq 101 -or $a -eq 115 -or $a -eq 120){ AddRef $v "Spell#$n($($s["Name"])).Ab$a" } }
}
foreach ($m in (Rows "SELECT * FROM Monsters")) {
  $n=[int]$m["Number"]
  AddRef ([int]$m["Move Msg"]) "Mon#$n($($m["Name"])).MoveMsg"
  AddRef ([int]$m["Death Msg"]) "Mon#$n($($m["Name"])).DeathMsg"
  for($i=0;$i -le 4;$i++){ AddRef ([int]$m["Attack Hit Msg $i"]) "Mon#$n($($m["Name"])).AttHit$i"; AddRef ([int]$m["Attack Miss Msg $i"]) "Mon#$n.AttMiss$i"; AddRef ([int]$m["Attack Dodge Msg $i"]) "Mon#$n.AttDodge$i" }
}
foreach ($it in (Rows "SELECT Number,Name,[Hit Msg],[Miss Msg],[Distruct Msg] FROM Items")) {
  $n=[int]$it["Number"]
  AddRef ([int]$it["Hit Msg"]) "Item#$n($($it["Name"])).HitMsg"; AddRef ([int]$it["Miss Msg"]) "Item#$n.MissMsg"; AddRef ([int]$it["Distruct Msg"]) "Item#$n($($it["Name"])).DestructMsg"
}
foreach ($r in (Rows "SELECT * FROM Rooms WHERE [Type 0]>0 OR [Type 1]>0 OR [Type 2]>0 OR [Type 3]>0 OR [Type 4]>0 OR [Type 5]>0 OR [Type 6]>0 OR [Type 7]>0 OR [Type 8]>0 OR [Type 9]>0")) {
  $rm="Room $($r["Map Number"])/$($r["Room Number"])"
  for($i=0;$i -le 9;$i++){ $t=[int]$r["Type $i"]; if($t -eq 0){continue}
    $p1=[int]$r["Para1 $i"];$p2=[int]$r["Para2 $i"];$p3=[int]$r["Para3 $i"];$p4=[int]$r["Para4 $i"]
    switch($t){ 9 { AddRef $p3 "$rm.TrapPassMsg"; AddRef $p4 "$rm.TrapFailMsg" }
                24 { AddRef $p3 "$rm.SpellTrapFailMsg"; AddRef $p4 "$rm.SpellTrapPassMsg" }
                22 { AddRef $p3 "$rm.CastPassMsg"; AddRef $p4 "$rm.CastLookMsg" }
                10 { AddRef $p1 "$rm.TextMsg"; AddRef $p2 "$rm.TextFailMsg" }
                12 { AddRef $p1 "$rm.RemoteActionMsg" }
                5  { AddRef $p1 "$rm.ActionMsg" }
                default { } }
  }
}
$cn.Close()

# textblocks could reference messages too but skip

# parse MESSAGES.md
$lines = [System.IO.File]::ReadAllLines($md)
$recs = New-Object System.Collections.ArrayList
for($i=0; $i+2 -lt $lines.Count+1 -and $i -lt $lines.Count; $i+=3){
  $hdr=$lines[$i]; $start= if($i+1 -lt $lines.Count){$lines[$i+1].Trim()}else{""}; $end= if($i+2 -lt $lines.Count){$lines[$i+2].Trim()}else{""}
  $parts=$hdr.Split(":",4); if($parts.Count -lt 3){continue}
  [void]$recs.Add(@{name=$parts[0];flags=$parts[1];act=$parts[2];reply=$(if($parts.Count -gt 3){$parts[3]}else{""});start=$start;end=$end})
}

function ToRegex($t){
  $t2 = [regex]::Escape($t)
  $t2 = [regex]::Replace($t2, '\\\{[a-z0-9]+\}', '.*')
  return $t2
}
function LitChunk($t){ $k=$t.IndexOf("{"); if($k -lt 0){ return $t }; if($k -eq 0){ $j=$t.IndexOf("}"); $rest=$t.Substring($j+1); $k2=$rest.IndexOf("{"); if($k2 -ge 0){ return $rest.Substring(0,$k2) } else { return $rest } }; return $t.Substring(0,$k) }
function GameRx($l){ if($grx.ContainsKey($l)){ return $grx[$l] }; $e=[regex]::Escape($l.TrimEnd(" !.")); $e=$e -replace "%s",".*" -replace "%d",".*"; $r=New-Object System.Text.RegularExpressions.Regex("^"+$e); $grx[$l]=$r; return $r }
$grx=@{}
function FindMatches($text){
  $res=New-Object System.Collections.ArrayList
  if([string]::IsNullOrWhiteSpace($text)){ return $res }
  $rx = New-Object System.Text.RegularExpressions.Regex((ToRegex $text))
  $mdX = [regex]::Replace($text, "{[a-z0-9]+}", "X").TrimEnd(" !.")
  foreach($k in $msgs.Keys){ $ls=$msgs[$k]; for($li=0;$li -lt 3;$li++){ $l=$ls[$li]; if($l.Length -lt 6){continue}; $ok=$false; if($rx.IsMatch($l)){ $ok=$true } elseif($l.IndexOf("%") -ge 0){ $g=GameRx $l; if($g.IsMatch($mdX) -and $mdX.Length -ge 6){ $ok=$true } }; if($ok){ $rr= if($refs.ContainsKey($k)){ ($refs[$k] -join ";") } else { "NOREF" }; [void]$res.Add("msg$k/L$($li+1) <$rr>"); if($res.Count -ge 4){ return $res } } } }
  return $res
}

$sb = New-Object System.Text.StringBuilder
$i=0
foreach($r in $recs){
  $i++
  $sm = FindMatches $r.start
  $em = FindMatches $r.end
  [void]$sb.AppendLine(("REC {0}|{1}|{2}|{3}" -f $r.name,$r.flags,$r.act,$r.reply))
  [void]$sb.AppendLine(("  START: {0}" -f $r.start))
  [void]$sb.AppendLine(("    -> {0}" -f $(if($sm.Count){$sm -join " || "}else{"(no match)"})))
  if($r.end.Length){ [void]$sb.AppendLine(("  END: {0}" -f $r.end)); [void]$sb.AppendLine(("    -> {0}" -f $(if($em.Count){$em -join " || "}else{"(no match)"}))) }
}
[System.IO.File]::WriteAllText($out,$sb.ToString())
Write-Output "records=$($recs.Count) msgs=$($msgs.Count) refs=$($refs.Count)"
