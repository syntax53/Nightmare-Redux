# Plan: game-message export for the MME export

Source prompt: `PROMPT-mme-messages-export.md`. Status: **approved and implemented in the working tree (not compiled, not committed)**. Implementation notes: the message logic lives in `modMMEMessages.bas`; `CalcTotalRecords` was left unchanged (the message pass updates the status labels only); `OpenTables` probes the schema before opening `Messages`; the "Export Only" filler row sets the three new Items columns. Verification script: `docs/verify-mme-messages.ps1`.

All counts below come from the stock-ish sample export (`NMR-EuphStockishDataExport.mdb`: 1,429 spells, 1,111 monsters, 1,979 items, 297 trap exits, 3,891 messages) and from the engine sources cited in the prompt. I re-read the two display routines that matter most (`display_spell_success` in the stock DLL and `Spell.cs:1216-1300` in Paradigm) and Paradigm's `Message.cs` role comments; they agree with the prompt's role table.

---

## 1. Schema

### 1.1 New table `Messages` (one row per client-relevant line, keyed by surrogate id)

| Column | Jet type | Size | Null | Meaning |
|---|---|---|---|---|
| `ID` | adInteger | | no | Surrogate row id, 1-based in export order. Primary key `pkMessages`. |
| `Name` | adVarWChar | 50 | no | Spell name or monster name as the client sees it in play; `Chr(0)` for hit and trap rows (see 2.4) |
| `Kind` | adVarWChar | 12 | no | Role of `Text`: `Status`, `Target`, `Self`, `Fail`, `Shield`, `ShieldOwner`, `Hit`, `HitOther`, `Death`, `Trap`, `TrapFail`. The kind also tells the source: spell kinds, monster kinds (`Hit`, `HitOther`, `Death`), trap kinds. |
| `Text` | adVarWChar | 100 | no | The normalized line (tokens, see 2.2) |
| `EndText` | adVarWChar | 100 | no | Wear-off line for `Status` rows; `Chr(0)` otherwise |
| `Effects` | adVarWChar | 255 | no | CSV of effect identifiers (see 2.3); `Chr(0)` when none |
| `Duration` | adInteger | | no | Spell duration field (0 for instant and for non-spell rows) |

No source, record-number or attack-slot columns: a row can come from several records, and nothing a client does needs the originating record. `Kind` + `Name` is the whole key a client uses.

Why rows rather than one wide row per spell: a spell can yield up to five different lines with different roles, and monster/trap rows have no spell at all. Narrow rows dedupe cleanly and let a client index by `Name` + `Kind`. Clients that want "one record per spell" group by `Name`.

Why no message numbers: they would let anyone join the rows back to a sysop's full message file. `ID` is meaningless outside this export.

### 1.2 Appended columns on `Items` (after `References`)

| Column | Jet type | Size | Meaning |
|---|---|---|---|
| `HitVerbs` | adVarWChar | 74 | Item `HitMsg` line 1: the attacker's own verbs, `slash\|impale\|hack` |
| `HitVerbsOther` | adVarWChar | 74 | Item `HitMsg` line 2: third-person forms, `slashes\|impales\|hacks` (what the room sees when another player hits) |
| `MissVerbs` | adVarWChar | 74 | Item `MissMsg` line 1: `swing at`, `lunge at`, `thrust at` |

Line 3 of `HitMsg` equals line 2 in every sampled record, and `MissMsg` lines 2-3 are the third-person forms of line 1; they are not exported. 296 of 304 weapons carry verbs; 96 non-weapons carry them too (shields and odd items) and are exported as-is. Empty → `Chr(0)`.

### 1.3 Untouched contracts

- Table names `Items, Classes, Races, Spells, Monsters, Shops, Rooms, Info, TBInfo, Lairs`; all existing columns, types and order; indexes `pkItems, pkSpells, pkMonsters, pkClasses, pkRaces, pkShops, pkTBInfo, pkLairs, idxRooms`; the single `Info` row and its columns; `NMR Version` stays the current numeric string (no bump).
- No new columns on `Spells`, `Monsters`, `Rooms`. `Monsters` stays at its current width.

---

## 2. Rules

### 2.1 Which lines become rows (per source)

**Spells** (iterate every spell record; inclusion is "everything referenced", so simply every spell with a usable message):

| Kind | Taken from | Exported when | Text / EndText |
|---|---|---|---|
| `Status` | DescMsg (ability 115) line 3 / line 1 | DescMsg exists and line 3 or line 1 is non-blank | `Text` = line 3 (stat text, printed to the target when the spell lands; Paradigm also prints it on each poison tick), `EndText` = line 1 (wear-off) |
| `Target` | `CastMsgB` line 2 | line contains `%d` **and** is not a hard-coded form (2.5); **or** the spell has a status ability (19/71/107/74/75/60/76/170) but no DescMsg and line 2 is non-blank (28 spells; the cast line is then the only signal) | `Text` = normalized line 2 |
| `Self` | `CastMsgB` line 1 | line contains `%d` and is not a hard-coded form | `Text` = normalized line 1 (your own spell/item damage or heal: "You fire a {spell} at {target} for {dmg} damage!") |
| `Fail` | ConfuseMsg (ability 101) line 1 | non-blank | `Text` = line 1 ("You fumble in confusion!", "You are too afraid to do that!") |
| `Shield` | ShockMsg (ability 137) line 2 | non-blank | the attacker's line: "You are shocked for {dmg} damage!" |
| `ShieldOwner` | ShockMsg line 1 | non-blank | the owner's line: "{attacker} is shocked for {dmg} damage!" |

Not exported from spells: cast line 3 (room), StartMsg (ability 120, all lines: flavor gestures), DescMsg line 2 (look text), ConfuseMsg lines 2-3, ShockMsg line 3, `CastMsgA` (never read by any engine). Cast lines without `%d` are not exported because the status line already identifies the effect, and they are the most content-revealing lines.

Sample counts: 439 spells have a DescMsg (383 distinct name+text, 231 distinct text pairs); 349 target lines and 233 self lines carry damage in a non-hard-coded form; 77 ConfuseMsg; 6 ShockMsg.

**Monsters** (iterate every monster; attack slots 0-4 with `AttackType = 1`; type-2 attacks are spells and are covered by spell rows):

| Kind | Taken from | Exported when | Text |
|---|---|---|---|
| `Hit` | `AttackHitMsg(i)` line 1 | line is not the standard single-verb form (2.5) | "{monster} darts forward and bites you for {dmg} damage!" |
| `HitOther` | `AttackHitMsg(i)` line 2 | the same attack produced a `Hit` row, and line 2 is not the standard form `^(The \|An? )?\{monster\} [a-z-]+ \{target\} for \{dmg\} damage!$`; deduped by text | "{monster} darts forward and bites {target} for {dmg} damage!" (what you see when a party member is hit) |
| `Death` | `DeathMsg` line 3 (fallback: `AttackHitMsg(0)` line 3 when `DeathMsg` ≤ 1) | non-blank | "The giant rat falls to the ground with a tortured squeak." |

Sample: 1,745 melee attacks → 241 distinct hit lines, of which 181 are non-standard (multi-word verbs, "with its pincers", "all-out slashes"); roughly the same number of distinct `HitOther` lines; 585 distinct death texts.

**Traps** (iterate rooms; exit types 9 and 24):

| Kind | Taken from | Text |
|---|---|---|
| `Trap` | Para3 line 1 (type 9), Para4 line 1 (type 24, stock trigger; Paradigm uses Para3) | "An arrow shoots out of the wall and strikes you!" |
| `TrapFail` | Para4 line 1 (type 9), Para3 line 1 (type 24) | "You try and disarm the trap, but trigger it!" |

Because stock and Paradigm swap Para3/Para4 for type 24, both messages of a type-24 exit are exported and the `Kind` is assigned by the stock rule; the consumer doc says the two kinds are interchangeable for parsing. Excluded rooms are **not** skipped (inclusion is everything), but buffer rooms are. Sample: 297 exits → 27 + 20 distinct texts.

Not exported: monster move/miss/dodge messages, item destruct messages, exit command words, text/action/hidden exit narrative, textblocks.

### 2.2 Normalization (tokens)

Replace the n-th `%s`/`%d` with the n-th token of the role's argument list. Lines are read with `ClipNull` (stop at the first `Chr(0)`), never `RTrim` — the full-dump export used `RTrim` and shows garbage after the terminator ("You can move again! person wear off!").

| Role | `MsgStyle` bit 0 = 0 | bit 0 = 1 |
|---|---|---|
| Cast line 1 (`Self`) | `{spell}`, `{target}`, `{dmg}` | `{target}`, `{dmg}` |
| Cast line 2 (`Target`) | `{caster}`, `{spell}`, `{dmg}` | `{dmg}` |
| DescMsg line 1 (`EndText`) | `{spell}` (stock passes the spell name; no sampled line uses it) | same |
| DescMsg line 3 (`Status`) | none (no sampled line has a placeholder) | same |
| ConfuseMsg line 1 (`Fail`) | none | same |
| ShockMsg line 1 (`ShieldOwner`) | `{attacker}`, `{dmg}` | same |
| ShockMsg line 2 (`Shield`) | `{dmg}` | same |
| Monster hit line 1 (`Hit`) | `{monster}`, `{dmg}` | n/a |
| Monster hit line 2 (`HitOther`) | `{monster}`, `{target}`, `{dmg}` (the damage arrives as `%s` here; the third placeholder is always mapped to `{dmg}`) | n/a |
| Death line 3 | none | n/a |
| Trap lines | `{dmg}` (stock passes damage for line 1; sampled texts have none) | n/a |

Rules: `%d` always maps to `{dmg}` regardless of position (the engines only ever pass one number; on room-perspective lines such as hit line 2 it arrives as the last `%s`, which the role table maps to `{dmg}`). A `%s` beyond the role's list becomes `{?}` and is logged. Monster-cast spells use the same cast-line table (the monster name fills `{caster}`; the data already contains "The " where needed). 462 sampled spells have `MsgStyle` bit 0 set, so both tables are exercised. Area/room spells pass `""` for `{target}`; the token stays in the text and the doc says it may be empty. Tokens are literal `{name}` so they survive Jet and read naturally in MegaMud-style templates.

Unmatched or odd lines are written to `NMR-Log_MMEMessages.txt` beside the MDB (one line each: spell/monster number, role, raw text, reason) and the completion dialog reports the count. Nothing is silently dropped: a flagged line is still exported with `{?}`.

### 2.3 Effects vocabulary (computed from the spell's ability slots, duration and target)

| Effect | From | Evidence |
|---|---|---|
| `Poisoned` | 19 | ticks "You feel ill." in stock, DescMsg L3 in Paradigm |
| `Confused` | 71 | `check_confusion`; `Fail` row carries ConfuseMsg |
| `Blinded` | 107 | "You are blind!" hard-coded on actions |
| `Held` | 74 or 75 | move attempts fail; Freedom (81) ends it |
| `Fear` | 60 | random flee, "You are too afraid!" |
| `Mute` | 76 | |
| `Sleep` | 170 | |
| `Damage` | 1 or 17 with duration 0 | |
| `DamageOverTime` | 1 or 17 with duration > 0 (73 spells) | silent ticks every 3 s |
| `Drain` | 8, 175 | |
| `Heal` | 18 with duration 0; `HealOverTime` with duration > 0 (23 spells) | |
| `HPRegen` | 123 | |
| `ManaRegen` | 145 | |
| `ManaHeal` | 150 | |
| `ManaSteal` | 174, 176 | |
| `DamageShield` | 72 | ShockMsg rows |
| `Cure` | 20, 73, 81, 84, 122 | |
| `Timed` | duration > 0 | |
| `RoomSpell` | spell number appears in any room's `Spell` field (collected during the room pass, so the spell pass runs after rooms) | stock prints only cast line 2 with "The room" as caster |
| `Trap` | trap rows | |
| `Death` | death rows | |

Ambiguities found in the sample and how they are handled: 182 timed spells with a DescMsg have none of the abilities above (pure stat buffs such as bless, magic armour, way of the owl) → `Timed` only; the ability columns already in `Spells` give the detail. 7 spells have 71 without a ConfuseMsg → `Confused` with no `Fail` row. No spell combines 74 and 60. Charm (6), stat modifiers, light, teleport are not effects a client acts on and get no identifier. "Diseased" and "can't attack" have no engine ability; the doc tells clients to map `DamageOverTime` + stat drains, and `Sleep`/`Held`, respectively.

Monster `Hit`/`HitOther` rows get `Damage`; `Death` rows `Death`; trap rows `Damage,Trap`; `Shield` rows `Damage,DamageShield`.

### 2.4 Dedupe

Key = `Kind|Name|Text|EndText|Effects` (case-sensitive, after normalization). First occurrence wins. `Hit`, `HitOther` and `Trap`/`TrapFail` rows use `Name = Chr(0)` so identical lines from different monsters/rooms collapse (the `{monster}` token already carries the identity); `Death` rows keep the monster name. Expected total ≈ 2,100 rows in the sample.

### 2.5 Exclusion patterns (lines covered by client hard-coded parsing; applied after normalization)

- Cast line 1: `^You cast \{spell\} (on|at|upon) \{target\}(,| for| and)? .*\{dmg\} (damage|healing|hit points)!?$` — 25 sampled lines.
- Cast line 2: `^\{caster\} casts \{spell\} (on|at|upon) you(,| for| and)? .*\{dmg\} damage!?$` — 24 sampled lines.
- Monster hit line 1: `^(The |An? )?\{monster\} [a-z-]+ you for \{dmg\} damage!$` (exactly one verb word) — 860 attacks / 60 distinct lines. Custom single-word verbs ("chomps", "whomps", "headbutts") are excluded by this rule although MegaMud's verb list may not know them; if that proves to matter, the rule drops to "verb in a fixed stock list" in a follow-up. Lines with "all-out", "critically" etc. are multi-word and stay exported.
- Everything else is exported. `You take {dmg} damage ...` (hard-coded in MegaMud) stays exported because it is cheap and other clients may not have the rule.

---

## 3. File-by-file changes (execution order)

1. **`modMMEMessages.bas` (new module; registered in `Nightmare Redux.vbp` as `Module=modMMEMessages; modMMEMessages.bas`)** — all logic lives here so `frmMME_Export.frm` grows by a few lines only:
   - `Public Sub MME_ExportMessages(tabMessages As Recordset, ByRef bStop As Boolean, sLogPath As String)` — driver: room pass (collect `RoomSpell` numbers in a `Dictionary`, emit trap rows), spell pass (BGETFIRST/BGETNEXT on `SpellPosBlock`; `GetMessages` on `MessagePosBlock` is safe to call inside because it uses a different position block; **never** call `GetSpellName` inside the spell loop, it repositions `SpellPosBlock`), monster pass. Progress via a callback-free pattern: the form passes a label and calls happen through `frmMME_Export.lblPanel` as the other scans do, respecting `bUseCPU`/`DoEvents` and `bStopExport`.
   - `Private Function NormalizeLine(sLine, sTokens() As String) As String` (2.2), `Private Function SpellEffects(...) As String` (2.3), `Private Function IsStandardCastLine1/2`, `IsStandardHitLine` (2.5), `Private Sub AddRow(...)` with the dedupe `Dictionary`, `Private Sub LogLine(...)`.
   - Message reads go through `GetMessages(n, line)`; treat return values `"none"` and `"MSG NOT IN DATABASE!"` as blank (number ≤ 1 or missing record).
   - `Option Explicit`, `On Error GoTo error:` + `HandleError` in every procedure.
2. **`frmMME_Export.frm`**
   - Declarations (≈605-616): `Dim tabMessages As Recordset`.
   - `CreateTables` (6363-6886): `tabNewMessages` with the columns of 1.1 and `pkMessages` on `ID`; three `Items` columns appended after `References` (1.2); object cleanup in both the normal and error paths.
   - `OpenTables` (7015) / `CloseDB` (7034): open/close `Messages`.
   - `cmdGo_Click` (1444): after `Call OpenTables` and the existing `CheckVersion` block, when `bUpdateExistingADB` verify `DB.TableDefs("Messages")` and `DB.TableDefs("Items").Fields("HitVerbs")` exist (error-trapped); if either is missing: `MsgBox` "This export was made by an older version and has no Messages table. Create a new export instead." then `CloseDB(True)` and `GoTo ReEnable`. Add `Call MME_ExportMessages(...)` after `ExportMonsters` and before `ExportVersionInfo`, gated like the others: `If (chkOnly(0).Value = 0) Or (chkOnly(0).Value = 1 And chkOnly(8).Value = 1)`. Show the log count in the completion `MsgBox` when > 0.
   - `ExportItems` (5309): fill the three verb columns from `GetMessages(Itemrec.HitMsg, 1/2)` and `GetMessages(Itemrec.MissMsg, 1)`; blank/`none`/not-found → `Chr(0)`.
   - `CalcTotalRecords` (6935): add one spell-count and one monster-count unit for the message pass so the progress bar scales.
   - "Export Only" group: new `chkOnly(8)` "Messages" in the form layout (same frame, below "Shops"), `chkOnly_Click` loop bound `1 To 7` → `1 To 8`. Not persisted to the config file, matching the other `chkOnly` boxes.
   - `VerifyOneRecordInDBs`: no filler row for `Messages` (no consumer requires a non-empty table).
3. **`modMMUDFunctions.bas` `LookupAbility` (≈899)** and **`frmSpell.frm` (≈2074)**: add `137` beside `101, 115, 120`.
4. **`docs/mme_export_notes.txt`**: append the consumer section (section 5 below).
5. **`README.md`** "Unreleased": `-NEW: MME export: Messages table with the game text clients need for status effects, custom damage lines, monster death text and trap text (normalized {tokens}, Effects CSV); weapon verb columns on Items. Older MME versions ignore the additions; updating an export made by an older NMR is refused.`

### 3.1 Subagent work packages (Opus 5.5, one each, no shared procedures)

- **WP-A — `modMMEMessages.bas` + `.vbp` entry.** Brief carries sections 2.1-2.5 verbatim, the record-type field names (`Spellrec.CastMsgB`, `AbilityA/B`, `MsgStyle`, `duration`; `Monsterrec.AttackType/AttackHitMsg/DeathMsg`; `Roomrec.RoomType/Para3/Para4/Spell`), the Btrieve loop idiom copied from `ExportSpells`, and the `GetMessages` contract. Output: the module, no form edits.
- **WP-B — `frmMME_Export.frm` plumbing** (items 2 above) plus `LookupAbility`/`frmSpell` (item 3). Brief carries the exact column list of 1.1-1.2, the call signature of `MME_ExportMessages` fixed in advance so A and B do not depend on each other's output, and the `.frm` control-layout snippet for `chkOnly(7)` to clone.
- **WP-C — docs** (items 4-5) after A and B are reviewed, so the doc text matches the final column names.

Review (me): read every diff against this plan; check `Option Explicit`/error handlers; check no existing line in `CreateTables` moved; check CRLF preserved (`git diff --stat` should show only additions in the form, and `file` reports CRLF).

---

## 4. Edge cases

- Message number 0 or 1: blank record in every sampled file → no row. `GetMessages` returns `"none"` for 0 and the record's blank lines for 1.
- Missing record: `"MSG NOT IN DATABASE!"` → no row, logged once per source.
- 74-char fields without a terminator: `ClipNull` stops at `Chr(0)`; lines of exactly 74 characters are kept whole.
- `%s` holding a number on room lines: not exported (room lines are excluded), but `NormalizeLine` maps `%d` → `{dmg}` and `%s` by position, so if a room line is ever added the table in 2.2 extends.
- `MsgStyle` bit 0: both token tables implemented; other bits ignored (stock reads only bit 0).
- Monster `DeathMsg` equals the hit-message record (stock convention): `Death` reads line 3 of `DeathMsg`; if `DeathMsg` ≤ 1 fall back to attack 0's hit record line 3; identical texts dedupe by `Name|Text`.
- Monster-cast spells with no caster line and "The %s" written into the text: `{caster}` token replaces the `%s`, leaving "The {caster} casts {spell} on you!" — documented.
- Paradigm differences: type-24 Para3/Para4 swap (both exported), poison ticks print `Status` text repeatedly, room spells print the `Target` line with an empty caster — all documented, none branched on.
- `EndCast`/`CastOnEnd%`/`KillSpell` chains: chained spells are ordinary spell records and get their own rows; no chain walking.
- Spells with the same name but different texts (five "fear"s with two DescMsgs): separate rows, same `Name`; the doc tells clients a name can map to several `Status` rows.
- Blank `Name` (unnamed spells such as number 2): row kept, `Name` = `Chr(0)`, logged.
- Excluded/hidden rooms: trap rows are still exported (inclusion is everything); buffer rooms skipped as in `ExportRooms`.
- "Export Only → Messages" with an existing MDB: rows are appended after clearing the table (`tabMessages` delete-all first, as `ExportLairs` does), never merged.

---

## 5. Consumer documentation (to append to `docs/mme_export_notes.txt`)

```
-= Messages (NMR 1.8.4+) =-
One row per game line a client may need to recognise. Not a dump of the realm's message file:
message numbers, caster-only lines, room lines and flavor text are not exported.

Columns: ID (row id), Name (spell or monster name; empty for Hit/HitOther/Trap rows), Kind, Text,
EndText, Effects, Duration (spell duration, 0 otherwise).

Kind:
  Status      Text = line printed to the affected player when the effect starts and shown by STAT;
              EndText = line printed when it wears off. (GreaterMUD also reprints Text on poison ticks.)
  Target      what the target sees when the spell hits them (damage/heal amount in {dmg})
  Self        what the caster sees for their own damage/heal spell or item
  Fail        line printed when a confused/afraid player's action fails
  Shield      "you are shocked" line when you hit a damage-shielded target
  ShieldOwner the shield owner's line
  Hit         monster melee hit line that does not follow "{monster} <verb> you for {dmg} damage!"
  HitOther    the same hit as others in the room see it ("{monster} ... {target} for {dmg} damage!")
  Death       monster death text (the monster's Name is set)
  Trap        exit trap trigger line; TrapFail = failed disarm line (both do damage;
              on GreaterMUD the two may be swapped for spell traps)

Tokens: {caster} {spell} {target} {dmg} {monster} {attacker}. Match them as wildcards; {dmg} is an
integer. {target} may be empty for area spells. "{?}" marks a placeholder NMR could not identify.
Lines that the standard client formats already cover are omitted:
  You cast {spell} on {target} for {dmg} damage!  /  {caster} casts {spell} on you for {dmg} damage!
  {monster} <single verb> you for {dmg} damage!
Effects (CSV): Poisoned Confused Blinded Held Fear Mute Sleep Damage DamageOverTime Drain Heal
  HealOverTime HPRegen ManaRegen ManaHeal ManaSteal DamageShield Cure Timed RoomSpell Trap Death.
  Disease = DamageOverTime plus stat drains; "cannot attack" = Sleep or Held.
A spell name can appear in several Status rows when the realm has variants with different text.

-= Items (NMR 1.8.4+) =-
HitVerbs / HitVerbsOther / MissVerbs: "|"-separated verb alternatives used in
"You <HitVerbs> X for N damage!", "<player> <HitVerbsOther> X for N damage!" and misses.
```

(Version label to be whatever the release ends up as; no bump is made by this change.)

---

## 6. Verification

**Scripted (me, after the build):** a 32-bit Jet PowerShell script run via `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -File`:
1. Schema diff: every table/column/index of `data-Paradigm-1.9.1x.mdb` exists in the new MDB with the same type; the only additions are `Messages` (+`pkMessages`) and `Items.HitVerbs/HitVerbsOther/MissVerbs`.
2. Row checks: `bless` → `Status` "You feel lucky!" / "The effects of bless wear off!" / `Timed`; `acid rain` → `Status` "You are covered in acid!" / "The acid covering you dries up." / effects contain `DamageOverTime` and `Timed`; `poison cloud` and `bites` → `Poisoned` with "You feel ill."; `magic missile` → `Self` "You fire a {spell} at {target} for {dmg} damage!" and `Target` "{caster} fires a {spell} at you for {dmg} damage!"; `giant rat` → no `Hit` row (standard verb) but `Death` "The giant rat falls to the ground with a tortured squeak."; `lashworm` → `Hit` "{monster} darts forward and bites you for {dmg} damage!" and `HitOther` "{monster} darts forward and bites {target} for {dmg} damage!"; a type-9 trap row "An arrow shoots out of the wall and strikes you!" with `Damage,Trap`; `longsword` → `HitVerbs` "slash|impale|hack".
3. Negative checks: no `Text` contains `%s`/`%d`; no row matches the exclusion regexes; no `Kind` outside the list; `Name` non-empty for `Status/Target/Self/Fail/Death`.
4. MegaMud spot check: 10 `MESSAGES.md` records (acid rain, bless, confusion, hold person, poison cloud, magic missile, fireball, blind end, arrow trap, giant rat death) each derivable from a row.

**Manual (you):** build in the IDE; fresh export; update-existing run against a pre-change MDB (expect the refusal dialog, MDB untouched); open the new MDB in the current MMUD Explorer (no missing-field warnings); open `NMR-Log_MMEMessages.txt` and send me its line count and a sample.

---

## 7. Decisions taken with the author

- `Hit` rows drop the monster name and dedupe by text.
- Hit line 2 is exported as `HitOther` when unique.
- The single-verb exclusion rule stays as proposed.
- Table name `Messages`; columns reduced to `ID, Name, Kind, Text, EndText, Effects, Duration` (no source, record number or attack slot).
