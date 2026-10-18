# Task: Add game-message export to the MME export (new Messages table + links), backward compatible

## Orchestration

You (Claude Fable 5.1) are the planner and overseer. Enter plan mode, study the code and facts listed below, verify the facts you rely on against the sources cited, and produce a detailed implementation plan for my approval. After I approve, delegate the coding to Claude Opus 5.5 subagents (Agent tool with `model: "opus"`), one per independent work package, each with precise, self-contained instructions from your plan. Review every diff yourself line by line against the plan before reporting. Do not write the implementation code yourself; do write the plan, the subagent briefs, the review, and the verification results.

This is VB6. **Nobody but me can compile or run it.** Subagents therefore cannot test; the plan must make each change reviewable by reading, must keep the codebase conventions in `CLAUDE.md`, and must end with a manual verification checklist for me (build in the VB6 IDE, run the export, inspect the MDB). Where a verification can be scripted without VB6 (querying the produced `.mdb`), write the script for me (see Verification).

## Goal

Extend `frmMME_Export` (the "MME export", consumed by MMUD Explorer and, going forward, by game clients/bots) so that it exports the game **messages** that clients need to react to the game engine: status ailments and their wear-off text, timed buffs, damage/heal lines that are not in the engine's standard hard-coded formats, trap text, each tied to the spell, monster attack, item or trap that produces it, plus weapon verb lists. The export must stay **backward compatible with older MME versions**: existing tables, columns, column types, index names and the `Info` row must be unchanged; new tables and columns are additive.

## Background

- The MME export has two goals: (a) give players realm data for exploration without sysops hand-feeding it; (b) **not** make it easy to reconstruct the sysop's content from the export. Messages are content, so the export must be useful to a client parser while leaking as little as the client doesn't need.
- Until now the dominant client (MegaMud) shipped a hand-mined `MESSAGES.md`. Every realm-custom message had to be discovered by players. Several new clients are in development; we want realm exports to carry what the clients need.
- `README.md` is the changelog (newest on top, there is an "Unreleased" section); `docs/mme_export_notes.txt` documents the in-game cross-reference scan.

## Established facts (verified during prompt preparation; re-verify anything you build on)

### NMR data model

- The message record (`modFieldmaps_vO.bas:484-497`, identical in `_vN`) is `Number` + three 74-char lines. **No record layout changes are needed for this task**, so the vN/vO two-module rule does not bite; both EXEs share this code.
- What a line *means* depends entirely on which record field points at the message. The authoritative role table (from the stock DLL decompile and Paradigm) is below. NMR's own help text for the message editor says the same thing (`frmHelpMessages.frx`, readable as plain strings).

| Referenced by | Line 1 | Line 2 | Line 3 |
|---|---|---|---|
| Spell `CastMsgB` (the cast message) | caster sees | target sees | room sees |
| Spell ability 115 `DescMsg` | **wear-off text** (to the affected player only; stock passes spell name as `%s`) | "look at" text (stock only; Paradigm ignores) | **`stat` line; also printed to the target when the spell lands** (even on self-cast). Paradigm also prints it on each poison tick and on held move attempts |
| Spell ability 120 `StartMsg` | caster | target | room |
| Spell ability 101 `ConfuseMsg` | confused player, each fumbled action | room | (unused in stock; MegaMud treats as stat line) |
| Spell ability 137 `ShockMsg` (with 72 DamageShield) | shield owner | attacker shocked | room |
| Spell `CastMsgA` | **never read by either engine**; usually 0 or 1 | | |
| Monster `AttackHitMsg(i)` | victim (`%s` monster, `%d` dmg) | room (`%s` monster, `%s` victim, `%s` dmg-as-string) | (stock data stores the monster's death text here; `DeathMsg` usually points at the same record) |
| Monster `AttackDodgeMsg(i)` | victim, armour glance | room, glance | victim, dodge |
| Monster `AttackMissMsg(i)` | room, victim dodged | victim, miss | room, miss |
| Monster `MoveMsg` | arrival text | departure | follow text |
| Monster `DeathMsg` | | | death text |
| Item `HitMsg` / `MissMsg` | attacker verbs `slash\|impale\|hack` | victim verbs | room verbs |
| Item `DistructMsg` | user | room | |
| Room exit type 9 Trap | mover (`%d` damage, stock only) | origin room | destination room (stock) |
| Room exit type 24 Spell trap | as above; **stock uses Para4 for the trigger text, Paradigm uses Para3** | | |
| Exit types 3/5/6/10/12/22/23 | narrative / command words (see table below) | | |

- Spell → message links: `CastMsgB` (field), plus ability slots `AbilityA(x)=101/115/120/137` with `AbilityB(x)` = message number. NMR already treats 101/115/120 as message links (`modMMUDFunctions.bas:899`, `frmSpell.frm:2074`) but **not 137**. Abilities 151 `EndCast`, 164 `CastOnEnd%`, 153 `KillSpell`, 43 `CastsSp` chain to *spells*, which carry their own messages.
- Monster → message links: `AttackHitMsg(4)`, `AttackDodgeMsg(4)`, `AttackMissMsg(4)`, `MoveMsg`, `DeathMsg`. Monster → spell links already exported: `AttHitSpell-x` (spell cast on a successful hit, 100%), `AttAcc-x` holds the spell number when `AttType-x = 2`, `MidSpell-x`, `DeathSpell`, `CreateSpell`.
- Item → message links: `HitMsg`, `MissMsg`, `DistructMsg`. Item → spell links are already in the exported ability columns: `114 %Spell` followed by `43 CastsSp` (hit spell), `43` alone (cast on use), `72`+`137` (damage shield).
- Room → spell: `Roomrec.Spell` (already exported as `Spell`). Room exits (`Para1..Para4`, see `frmRoom.frm:5700-5960` tooltips and Paradigm `Rooms\RoomManager.cs:223-370`): type 9 Trap = Para1 max damage, Para2 state, Para3 trigger msg, Para4 disarm-fail msg; type 24 Spell trap = Para1 spell, Para3/Para4 messages (engine-dependent, above); type 22 Cast = Para1/Para2 spells, Para3 pass msg, Para4 look msg; type 10 Text / 12 Remote action / 5 Action / 3 Item / 6 Hidden / 23 Ability carry narrative or command-word messages.
- Classes, Races, Actions, Shops: no message references. Textblocks: Paradigm supports a `message N` textblock command (`TextBlockPart.cs:214-245`); not seen in stock data.
- Placeholders: engines never rewrite `%s` into "you"; perspective comes only from which line is sent to whom. Argument order for the cast message (stock, `MsgStyle` bit 0 clear): L1 = (spell, target, value), L2 = (caster, spell, value), L3 = (caster, spell, target, dmg-as-string); bit 0 set: L1 = (target, value), L2 = (value), L3 = (target, dmg). Room spells print **only L2** to the affected player with caster = "The room" (stock) or "" (Paradigm), then DescMsg L3. Damage on room lines is `%s`, not `%d`. Duration ticks print nothing except poison ("You feel ill." hard-coded in stock; DescMsg L3 in Paradigm).
- Status effects by ability (full list in `ability.mdb`, table `TABLE`): 19 Poison, 20 CurePoison, 21 ImmuPoison, 71 Confusion (+101), 107 BlindUser, 74 HoldPerson, 75 Paralyze, 81 Freedom, 60 Fear, 76 Mute, 170 Sleep, 1/17 Damage (timed = damage over time), 8 Drain, 18 Heal (timed = heal over time), 123 HPRegen, 145 ManaRgn, 150 HealMana, 174/175/176 mana/HP steal, 72 DamageShield, 73 DispellMagic, 84 RemoveCurse, 122 RemovesSpell. **There is no "disease" ability**; disease is modelled as timed damage plus stat drains. "Can't attack" has no single ability either (sleep/petrify-style spells use confusion 100% or hold).

### MegaMud's consumer format (what a client needs)

- `MESSAGES.md` records are exactly three lines, no separators: `name:FLAGS:action:reply`, start text, ending text (may be empty). Flags (`MEGAMUD.H:1708-1722`): 0001 Blinded, 0002 Confused, 0004 Poisoned, 0008 Losing HPs, 0010 Can't move, 0020 Can't attack, 0040 Diseased, 0080 HP regen, 0200 Mana regen/flux, 1000 Ends combat, 2000 Last action failed; matching options 0100/0400/4000/8000. Action 0-6 = ignore / check room / wait it off / rest HP / rest mana / run / hangup. Start text is a prefix match (template wildcards `{target} {source} {dmg} {1}-{5}`); ending text is literal. A record becomes an active tracked effect only if it has an ending text. Spells/items link to records **by name only**.
- Empirical attribution of the 837 shipped MegaMud records against a stock-ish NMR export (`docs/megamud-messages-attribution.txt`, script alongside): **ending texts are DescMsg line 1 (258 of 264 matched)**; **status start texts are DescMsg line 3 (235)**; damage/heal start texts are cast-message line 1 (137, mostly item-cast and area spells) or line 2 (46, monster/room spells hitting you); 35 "monster entry" records are `MoveMsg` line 1; ~10 are trap trigger lines (line 1); "chase" records are hidden/action exit line 2 (room text); a handful are monster death line 3 and ConfuseMsg line 1. The ~180 unmatched start texts are either non-stock content absent from the sample export or hard-coded engine strings (lives-left warnings, "Your weapon is blessed!", "You struggle for breath!").
- Hard-coded client formats that need no export are listed in `C:\Users\synta\Downloads\message details from megamud.txt` (standard `<attacker> <verb> you for <N> damage!`, `You cast <spell> on <target>...`, wear-off keyword lines, etc.).

### Current MME export mechanics (`frmMME_Export.frm`)

- `CreateTables` (6363-6886) builds the schema with ADOX, then `OpenTables` (7015) opens DAO table recordsets; `Export*` subs fill rows; `LocateRecords` (4974) runs the multi-pass in-game cross-reference (`ScanRooms/Shops/Monsters/Spells/Textblocks/Items/Containers/Greets/RoomCommands`, marking `In Game` and the `Learned From`/`Casted By`/`Obtained From` reference strings); `CalculateScriptValue`/`ExportLairs`; `FillTextblockCommands`; `VerifyOneRecordInDBs` (filler rows for "Export Only" mode); `CloseDB` compacts.
- Update-existing-MDB path: `CreateDatabase` returns 2 → `bUpdateExistingADB`, `CheckVersion` warns on version mismatch, and each `Export*` does Seek/Edit instead of AddNew. **An older MDB will not have the new table/columns**; see decisions.
- Config: `MME-Config.ini` keys via `LoadConfig`/`SaveConfig` (883-1027) with the `RoomCommandItems` toggle as the newest example; "Export Only" checkboxes `chkOnly(0..7)`; "Legit" export (`chkLegit(0)`, `EnableLegit`) and "For GreaterMUD Engine" (`chkLegit(1)`) flags are written to `Info.Legit` (1/2).
- `Info."NMR Version"` is `sAppVersion` built from `App.Major/App.Minor` (`frmMain.frm:517`); MME parses it numerically and gates features at ≥1.7/1.71/1.8/1.82/1.83.
- Messages dat is already opened at startup (`modMain.bas:1369`); `GetMessages(n, line)` in `modMMUDFunctions.bas:1884` reads one record (line 0 = all three comma-joined). `frmDatabaseExport.frm:2450 ExportMessages` is the full-dump precedent (Number, Line 1..3).

### MMUD Explorer compatibility rules (from reading `C:\Users\synta\GitHub\MMUD-Explorer`)

- MME opens the MDB read-only with DAO 3.6, opens the nine required tables (Items, Classes, Races, Spells, Monsters, Shops, Rooms, Info, TBInfo; Lairs when version ≥ 1.83) as table-type recordsets, and reads every column **by name**. It never enumerates tables or columns. Only `rsPing.Fields(0)` is positional, on a single-column query.
- Therefore: adding tables and adding columns (anywhere, but append for tidiness) is safe; renaming/retyping/removing columns, changing the `pk*`/`idxRooms` index names or key columns, adding Nulls to existing columns, leaving `.mdb`/Jet 4 format, or making "NMR Version" non-numeric would break it. Jet's 255-columns-per-table limit matters for `Monsters` (already ~170 columns).
- Reserved table names: the ten above. `Messages` is free.

## Files to study before planning

- `CLAUDE.md`, `README.md` (Unreleased section), `docs/mme_export_notes.txt`
- `frmMME_Export.frm` — whole file; especially `CreateTables`, `OpenTables`, `CloseDB`, `cmdGo_Click` (1444), `ExportSpells` (5405), `ExportMonsters` (5709), `ExportItems` (5309), `ExportRooms` (5838, exit encoding), `LocateRecords` + `Scan*`, `VerifyOneRecordInDBs`, `CalcTotalRecords`, `LoadConfig`/`SaveConfig`, `ExportVersionInfo`, `CheckVersion`
- `modFieldmaps_vO.bas` — `MessageRecType` (484), `SpellRecType` (100-140), `MonsterRecType` (180-235), `ItemRecType` (320-340), `RoomRecType` (395-460); confirm `_vN` matches for these types
- `modMMUDFunctions.bas` — `GetMessages` (1884), `LookupAbility` (880-910), `GetMonsterAttackName` (959, already derives the attack verb from hit-message line 1), `GetRoomExitType` (102)
- `frmDatabaseExport.frm` `ExportMessages` (2450) and `CreateTables` there — precedent for a Messages table
- `frmRoom.frm` 5700-5960 — exit parameter meanings per type
- `frmHelpMessages.frx` — message line roles as documented in-app (plain strings in the binary)
- `ability.mdb` — ability numbers/names (table `TABLE`)
- `docs/megamud-messages-attribution.txt` (+ `.ps1`) — which NMR line/reference each MegaMud record came from
- `C:\Users\synta\Downloads\message details from megamud.txt` — client hard-coded formats
- MegaMud source for the consumer format: `C:\Users\synta\GitHub\MegaMud2\MegaMudSource\MESSAGES.C` (load 1658-1774, match 575-1006, wear-off 1281-1497, template 3747-3895), `MEGAMUD.H:1688-1746`, `COMBAT.C:4063-4280` (damage-line credit), `PARSE.C:3557-3595`
- Engine semantics, Paradigm: `GreaterMUD.Module\Message.cs` (line-role comments 40-158), `Spells\Spell.cs:841-1411` (cast display), `Spells\SpellEffectManager.cs:89-93` (wear-off), `Player.cs:8090-8166` (poison tick), `ActionFigure.cs:1096-1183` (confusion/held), `Mobs\Mob.cs:1824-2052` (monster attack text), `Exits\TrapExitBase.cs`, `TrapExit.cs`, `SpellTrapExit.cs`, `Rooms\RoomManager.cs:223-370`, `Abilities\GMUDAbilities.cs`, `Items\ItemType.cs:98-207`
- Engine semantics, stock DLL decompile: `C:\Users\synta\GitHub\wcc-dll\analysis\NOTES.md` first, then `analysis\decomp\decomp_0043cd0a.c:1275-1488` (`display_spell_success`), `decomp_00424361.c:852-963` (`monster_display_spell_success`), `:6087-6365` (`attack_monster_user`), `decomp_0041c815.c:3448-3484` (`check_confusion`), `:4711-4726` (poison tick), `decomp_0046cc26.c:5726-6390` (room spells), `decomp_00401000.c:12101-12150`, `:12590-12640` (traps), `analysis\types\mmud_types.h:466-481`
- MME consumer: `C:\Users\synta\GitHub\MMUD-Explorer\modMMudDatabase.bas:888-1009` (`OpenTables`), `modReachability.bas:541-561` (`FieldExists` guards), `frmMain.frm:30905-31029` (`LoadInfo` version gating)
- Sample data: `P:\MajorMUD\_Euphoria\NMR-EuphStockishDataExport\NMR-EuphStockishDataExport.mdb` (full-dump export incl. a `Messages` table; query it with 32-bit Jet, see Verification) and `G:\My Drive\mud\dats\data-Paradigm-1.9.1x.mdb` (current MME export schema)

## Scope

### Part 1 — New `Messages` table: derived effect rows, not a message dump

One row per **client-relevant message as a specific source produces it**, not one row per Btrieve message record. The raw three-line records are never exported, and message numbers are not exported (use a surrogate row id). Sources:

- **Spells** (reached from anywhere: player-learnable, monster attack/mid-round/death/create spells, item cast-on-use and `%Spell` hit spells, room spells, spell-trap exits, `EndCast`/`CastOnEnd%` chains). For each spell with a message worth exporting, emit a row with the spell's name and the texts the client needs, from these roles: DescMsg line 3 (start/stat text), DescMsg line 1 (wear-off text), cast message line 2 (what the target sees), cast message line 1 only when the caster is also the one affected (self/area casts and item use), ConfuseMsg line 1, ShockMsg lines 1-2. Caster-only and room-only lines are **not** exported unless the plan shows a client need for a specific role.
- **Monster attacks**: per attack slot with a non-standard hit message, the victim line (line 1) and, if the plan finds party-damage tracking needs it, the room line (line 2) with the damage token. **Death text** (`DeathMsg`/hit-message line 3) is exported. Miss, dodge and move messages are not.
- **Exit traps and spell traps** (exit types 9 and 24): the mover line of the trigger message (and the disarm-fail message) as damage rows. They need no link back to the room; rooms already describe traps in their direction strings. Spell traps additionally resolve to their spell's rows.
- **Room spells**: resolve to the spell's rows; the plan must note that stock prints only cast line 2 with "The room" as caster.

Rules:
- Every text is **normalized**: `%s`/`%d` become named tokens using the per-role argument order in the facts above (`{caster}`, `{spell}`, `{target}`, `{dmg}`, `{monster}`, `{item}`), including the `MsgStyle` bit-0 variants and the "damage arrives as `%s` on room lines" case. The plan lists the exact token set and the substitution table per role; unknown positions must be flagged, never silently dropped.
- **Only what is necessary**: messages whose rendered form is fully covered by the client hard-coded formats (`message details from megamud.txt`, e.g. `You cast {spell} on {target} for {dmg} damage!`, `{monster} {verb} you for {dmg} damage!`) are omitted. The plan must list the exact exclusion patterns for approval, and must err toward exporting when a line is only partially covered.
- **Dedupe**: rows with identical name, texts and effects collapse to one row (e.g. five spells named "fear" sharing one DescMsg). Rows with identical texts but different source names stay separate, because clients link by name.
- **Inclusion is everything referenced**, independent of the `In Game` scan, excluded rooms, Legit mode and GreaterMUD mode. Messages are always exported; there is no enable toggle. A blank record (number 0/1) or a missing record produces no row.
- **Effects** column: a text CSV computed by NMR from the source's abilities, duration and target type. Starting vocabulary (edit as the engine evidence requires, and list ambiguous spells found in the sample export): `Poisoned` (19), `Confused` (71), `Blinded` (107), `Held` (74/75), `Fear` (60), `Mute` (76), `Sleep` (170), `Damage` (1/17, instant) / `DamageOverTime` (1/17 with duration), `Drain` (8), `Heal` (18) / `HealOverTime`, `HPRegen` (123), `ManaRegen` (145), `ManaHeal` (150), `ManaSteal` (174-176), `DamageShield` (72), `Cure` (20/73/81/84/122), `Timed` (duration > 0), `RoomSpell`, `Trap`, `Death`. Values must be stable identifiers a client can switch on.
- Columns to propose (final set is the planner's, justified in the plan): row id, source type, source name, attack index where relevant, start/stat text, wear-off text, target-cast text, self-cast text, effects, duration, plus whatever the token/role documentation needs. Jet types and sizes stated; string columns follow the existing `Chr(0)`-for-empty convention.

### Part 2 — Changes to existing tables (appended columns only)

- **Items**: export the weapon verb lists from `HitMsg` and `MissMsg` (`slash|impale|hack` forms) so clients can parse custom weapons; the plan decides whether attacker, victim and room forms are all needed and names the columns.
- **Spells / Monsters / Rooms**: no new columns are required, because the `Messages` rows carry their source. If the plan finds that a client needs a direct link column (e.g. a spell → row id), justify it and keep it to one column. Mind Jet's 255-column limit on `Monsters`.
- Treat ability 137 as a message link everywhere NMR treats 101/115/120 as one (`LookupAbility`, spell filter) — small, in-scope cleanup.

### Part 3 — Export plumbing

- `CreateTables`/`OpenTables`/`CloseDB`/`CalcTotalRecords`/`VerifyOneRecordInDBs`/progress labels/`cmdGo_Click` ordering. Messages depend on nothing from `LocateRecords`, so the planner picks the cheapest order.
- **Update-existing-MDB path**: if the target MDB lacks the `Messages` table (or the new Items columns), refuse the update with a clear message telling the user to create a new export; do not alter old schemas in place.
- **Config / UI**: no enable toggle. Optionally add a "Messages" entry to the "Export Only" checkboxes (`chkOnly`), saved like the others; the plan decides whether it is worth the UI work.
- Legit and GreaterMUD modes export messages unchanged. Paradigm-vs-stock differences (spell-trap Para3/Para4, room-spell caster text, poison tick text) are documented, not branched on, unless the plan shows a cheap `Info.Legit = 2` branch is clearly correct.
- No version bump. README "Unreleased" entry. `docs/mme_export_notes.txt` gets the consumer-facing schema description: tables, columns, token set, per-role argument order, effects vocabulary, exclusion patterns.

### Out of scope

- Any change to MMUD Explorer itself (separate repo); any Btrieve record layout change; the full database export (`frmDatabaseExport`); textblock parsing beyond what the in-game scan already does.

## Plan requirements

The plan you present for approval must include:

1. The final schema as a table: every new table/column, Jet type, size, nullability, index; and a statement of which existing MME contracts (table names, index names, column names/types, Info row) are untouched.
2. For every message role exported: which of the three lines are exported, under what column name, and why (client need), including the argument order of placeholders per role so the consumer doc can state it.
3. The exact exclusion patterns (lines covered by client hard-coded formats), the dedupe rule, and the normalization token table per role, each with an example from the sample export showing raw line → exported text.
4. The effects mapping with the engine evidence for each entry and a list of ambiguous spells found in the sample export.
5. A file-by-file change list with the procedures touched, in execution order, and the subagent work packages with their boundaries (no two subagents editing the same procedure).
6. Edge cases: message number 0/1 (blank record), missing records (`MSG NOT IN DATABASE!`), 74-char lines with trailing `Chr(0)`, `%s`-as-number on room lines, Paradigm-vs-stock differences (spell-trap Para3/Para4, room-spell caster text), `MsgStyle` bit 0, monsters whose `DeathMsg` equals the hit message record.
7. The consumer-facing documentation text (what a client developer reads to parse the new tables).
8. The manual verification checklist for me, and the scripted MDB checks (below).

## Verification (after implementation; you run the scripted parts, I run the IDE parts)

1. I build `Nightmare Redux.vbp` in the VB6 IDE and run a fresh MME export against my dats, then an "update existing MDB" run against an export made by the previous version, which must refuse cleanly. Report back exact dialogs/errors.
2. You verify the produced `.mdb` with a 32-bit Jet script (64-bit PowerShell cannot load Jet): run `C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -File <script>` using `ADODB.Connection` with `Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<mdb>`, `OpenSchema(20)` for tables, and `SELECT` queries for row counts and spot checks: the "bless" row (start "You feel lucky!", wear-off "The effects of bless wear off!", effects include `Timed`); "acid rain" (start "You are covered in acid!", wear-off "The acid covering you dries up.", `DamageOverTime`); "poison cloud"/"bites" (`Poisoned`, "You feel ill."); a non-standard monster hit line and a death text; a type-9 trap row; a weapon's verb columns; and that no row holds a raw `%s`/`%d` or a standard-format line. Then a schema diff against `G:\My Drive\mud\dats\data-Paradigm-1.9.1x.mdb` proving every pre-existing table/column/index is unchanged and the only differences are the new `Messages` table and the appended Items columns.
3. Regression: open the new MDB in the current MMUD Explorer build (I do this) and confirm it loads with no "missing fields" warnings.
4. Spot-check against MegaMud: pick 10 records from `MESSAGES.md` with different flag types and show that each one's start and ending text is derivable from the exported rows.

## Constraints

- Follow `CLAUDE.md`: `Option Explicit`, `On Error GoTo error:` + `HandleError`, CRLF line endings untouched, no reformatting of unrelated code, `bDisableWriting` irrelevant here (export is read-only on dats) but respect `bStopExport`/`DoEvents` progress conventions.
- Additive schema only; never reorder or rename existing columns; keep all `pk*`/`idxRooms` names.
- Do not commit; leave changes in the working tree on this branch.
- Both EXE variants (`_vN`/`_vO`) must keep building; no fieldmap edits unless a verified field is missing.
- Never export raw message records, message numbers, caster-only or room-only lines, or flavor text (monster movement, hidden-exit/chase text, item destruct, exit command words).
