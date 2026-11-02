# Plan: item uses in the MME export (decode into `Items`)

Branch `261101-item-messages`. Status: **sections 1-6 implemented, built and verified (commits 2b2dcc4,
31d488d). Revision 2 (spell-keyed rows, at the end) implemented, built and verified; it replaces 3.2 and
changes how 3.1 names rows.** The code is in `modMMEMessages.bas` (`ItemSpellLists`, `AddItemCaster`,
`ExpandItemChains`, `EndCastChain`, `AddCastRows`) and `frmMME_Export.frm` (`PrepareUseCodes`, `ItemUseCode`,
`NewUseCode`, `UseCodeWords`, `UseCodeFree`, and the `ExportItems` / `CreateTables` / `HasCurrentSchema` /
filler-row changes).

## 1. Goal and constraints

A client must be able to tell, for every item, what using it does, what hitting with it can do, how to
invoke it, and which `Messages` rows belong to it. That must work without:

- renaming spell rows (a class still casts `jade amulet` with its game code `jade`),
- made-up records or numbers (they cannot stay stable across incremental updates),
- exposing anything the export does not already contain. `Items` already exports all 20 `Abil-n` /
  `AbilVal-n` columns, so "CastsSp (43) -> spell number" and "%Spell (114) -> chance" are already public.
  This plan only decodes them and adds a generated command code.

The item is the record a client treats as "the thing I use". It has a real, stable key (`Items.Number`),
its name already labels its `Messages` rows, and `Spells` stays exactly as the game has it. A client that
wants MegaMud's shape builds its own "spell" per usable item, keyed by item number.

## 2. Schema: three columns appended to `Items` (after `MissVerbs`)

| Column | Type | Content |
|---|---|---|
| `UseCode` | Text 5 | `#` + up to 4 lowercase letters (digits only as a last resort); one code per distinct item name, so items with exactly the same name share it; `Chr(0)` when the item has no use spell (proc-only items get none) |
| `UseSpells` | Text 120 | spell numbers the item casts when used, `\|`-separated, in ability-slot order, duplicates removed; `Chr(0)` when none |
| `HitSpells` | Text 120 | spells cast on hit as `<spell>:<chance>` pairs, `\|`-separated, in slot order; `Chr(0)` when none |

Examples from the GreaterMUD sample: `shimmering longsword` = `#slon`, `114`, `170:40`;
`tankard of mead` = `#tmea`, `119|317`, empty.

### 2.1 Decoding rule (one shared routine)

Walk `AbilityA(0..19)` in order:
- `114` (%Spell): remember its value as the pending chance.
- `43` (CastsSp) with a value > 0: if a chance is pending, the spell goes to `HitSpells` with that chance
  and the pending chance is cleared ("the spell cast is specified by the NEXT CastsSp", ability.mdb);
  otherwise it goes to `UseSpells`.
- A `114` with no later `43` is ignored. A second `114` before a `43` replaces the pending chance.

The routine lives in `modMMEMessages.bas` as `Public Sub ItemSpellLists(sUse, sHit)` (reads the current
`Itemrec`), so `ExportItems` and the Messages item scan use the same rule.

### 2.2 `UseCode` generation

Only items with at least one use spell get a code (315 in the GreaterMUD sample; its 179 proc-only items
get none).

- Words: lowercase, drop a possessive `'s`, letters only; drop `a`, `an`, `the`, `of` unless nothing else
  is left.
- First choice:
  - One word: its first 4 letters (`incense` -> `ince`).
  - Several words: the first letter of each word; if that gives fewer than 4 letters, replace the last
    initial with the whole last word; cut to 4 (`jade amulet` -> `jamu`, `minor healing potion` ->
    `mhpo`, `tankard of mead` -> `tmea`, `embroidered black gauntlets` -> `ebga`).
  - No letters at all: `item`.
- Items with exactly the same name (case-insensitive) share one code: in game, "use void sphere" cannot
  tell them apart either.
- When the first choice belongs to a differently named item, try letter alternatives in this order, all
  keeping the first letter and the word order:
  1. the first 2, then 3 letters of the first word, the other initials, then the last word, cut to 4
     (`black potion` -> `blpo` when `blue potion` has `bpot`; `granite orb` -> `gror` after `glass orb`
     -> `gorb`);
  2. any 4 letters of the name in order, starting with its first letter.
- Digits only when every alternative is taken: `<first choice cut to 4 - len(n)><n>`, n = 2, 3, ...
- Items are processed in item-number order, so the lower-numbered item keeps the natural code.
- Sample (GreaterMUD): 315 usable items, 5 share a code with a same-named item (3 void spheres, a
  silverwood staff, a wooden box), 24 take a letter alternative (families such as coloured potions,
  chests, parchment deeds, keyrings), 0 need a digit.
- `#` keeps item codes apart from game spell codes: no spell short code in the sample starts with `#`.
  NMR does not check spell codes, and the docs say so.

### 2.3 Stability across updates

- A new export: same names and item order give the same codes.
- Updating an existing MDB: before the item loop, read every existing `Items` row's `Number`, `Name` and
  `UseCode`. An item keeps its earlier code if it still has a use spell and its name is unchanged. Kept
  codes are reserved for their name before any new code is generated, so a new lower-numbered item with
  a different name cannot take an existing item's code (a new item with the same name shares it). The code of a renamed item, or one that lost its use spell, is not reused in that
  run (it is free again in the next one).
- Docs tell clients to key item uses by `Items.Number`; `UseCode` is the command hint, not the key.

### 2.4 Schema gate

`HasCurrentSchema` also requires `Items.UseCode`. Updating an MDB made by an earlier 1.8.4 build is then
refused with the existing "create a new export" message (same as Titles/Abilities). 1.8.4 is unreleased,
so this is one schema bump.

## 3. Messages changes

### 3.1 EndCast chains (always)

After the Messages item scan, follow `EndCast` (spell ability 151) from every spell an item casts (use or
hit), transitively, with a visited set and a depth limit of 10. A chained spell is treated as cast by the
same items, so its caster line becomes a `Cast` row under each item's name (and its friendly-target
`Self` rows too, by the existing rule). The spell is read by number with `BGETEQUAL`, as
`frmMME_Export.frm:3810` does. Sample: 2 chains (`spear` -> `spear slam knockdown`, `elemental
earthquake` -> `elemental lightning`). Assumption, from the ability description: the chained spell's line
1 goes to the same caster.

`CastOnEnd%` (164) is a percentage that applies to the next 151, so following 151 already covers it.

### 3.2 Item markers on `Cast` rows (optional)

`Cast` rows written under an item's name get `Use` and/or `OnHit` appended to `Effects`, according to how
that item casts the spell (inherited through EndCast chains). A row under a player spell's name gets no
marker. One item that both uses and procs a spell with the same line gets both. Only `Cast` rows are
marked; status and incoming rows stay as they are. Old readers ignore unknown effects; MegaMud maps them
to 0.

Implementation: the item scan records the mode per spell and item name (`vbTextCompare`, so a spell-name
row whose name equals the item's name also matches). `SpellCastRow` and `SpellStartRow` call a new
`AddCastRows` that adds the marker per name, instead of `AddRowEach`.

### 3.3 Not changed

- Victim-side lines of hostile item spells stay nameless.
- Item destruct messages stay out.
- Spells textblocks cast when an item's spell runs a textblock (ability 148) stay nameless self-cast rows.

## 4. File-by-file

| File | Change |
|---|---|
| `modMMEMessages.bas` | `ItemSpellLists`; item scan records use/hit per spell and item; EndCast expansion; (3.2) mode dictionary and `AddCastRows` |
| `frmMME_Export.frm` | `CreateTables`: 3 columns; Export Only filler row: `Chr(0)` for all 3; `ExportItems`: decode, kept-code read in update mode, `UseCode` generation helpers; `HasCurrentSchema` checks `UseCode`, and its message mentions the item columns |
| `docs/MME_MESSAGES_HANDOFF.md` | 2.1 detection; 3.4 use vs on-hit and EndCast chains; 3.5/3.7/4.6 markers; 4.4 names; section 5 becomes "Item columns": verbs, use/hit spells, codes, and how a client models an item use; gotchas (key by `Items.Number`) |
| `docs/MEGAMUD_MME_MESSAGES_HANDOFF.md` | Section 9: building MegaMud item-use spell records from `UseCode`/`UseSpells`, keyed by item number with a mapping like 8.5; markers -> 0 in the section 7 table; O8 updated |
| `docs/mme_export_notes.txt` | Items section: 3 columns and rules; Messages: markers and EndCast |
| `docs/verify-mme-messages.ps1` | schema diff allows the 3 columns; decode `Abil-n` in PowerShell and compare with `UseSpells`/`HitSpells` for every item; `UseCode` format `^#[a-z0-9]{1,4}$`, present only with `UseSpells`, the same code for the same name and different codes for different names; spot check `jade amulet` has a `#ja..` code; (3.2) every `Use`/`OnHit` Cast row is named after an item with `UseSpells`/`HitSpells`; INFO count of usable items without a `Use` row |
| `README.md`, `docs/changelog.txt` | one 1.8.4 entry |
| `nightmare_redux.exe` | not rebuilt by me (needs the IDE) |

All `.bas`/`.frm` edits keep CRLF and follow the existing `On Error GoTo error:` / `HandleError` pattern.

## 5. Verification

- I cannot compile here (non-elevated VB6 `/make` fails on this machine). You build in the IDE.
- Then a stock v1.11p export, and `docs/verify-mme-messages.ps1 -New <export> -Old <1.8.3 export>`. I can
  run the script and read the results.
- Expected: every usable item has a code, codes unique, decode matches `Abil-n`, stock row counts change
  only by the EndCast-chain rows (and the marker text on item Cast rows).
- Update-mode check: export, rename one item and add one in the dat, update the same MDB, and confirm
  untouched items kept their codes.

Result (2026-11-01, built in the IDE; `data-v1.11p.mdb` and `data-Paradigm-1.9.1d.mdb`):
- Stock v1.11p: every check passes. 1,417 Messages rows (Cast 518 -> 523, Self and Other unchanged);
  209 usable items, 206 codes, no digits; 127 items with on-hit spells; 165 Cast rows with `Use`, 122 with
  `OnHit`. The five new rows are EndCast chains: Dusk and the nexus spear (spear -> spear slam
  knockdown) and the staff of the elements (elemental earthquake -> elemental lightning ...).
- Paradigm 1.9.1d: one failure, the stock bane dart spot check. The pre-change 1.9.1c export lacks that
  row too, so it is realm data, not this change. Against 1.9.1c, ignoring markers, the only difference is
  the same five chain rows; nothing was removed. 310 usable item names, no digits.
- Usable items without a `Use` line (58 stock, 123 Paradigm) have a use spell that prints no cast line
  (mostly chests, boxes and pouches); no marker is missing.
- Not run: the update-mode check.

## 6. Risks

- Clients that expect `Items` to end at `MissVerbs` by position: none known (MegaMud builds its column list
  by name; MME already handled the verb columns).
- The 114 -> next 43 reading is from the ability description, not from engine source.
- Codes change when an item is renamed. The docs say to key by item number.

---

# Revision 2: spell-keyed rows (approved, implemented and verified 2026-11-01)

## R2.1 Why

With `Items.UseSpells` / `HitSpells` a client can get from an item to its spells, so `Messages` no longer has
to repeat a spell's lines once per item. Each spell's lines are written once, under the spell's name, with
the spell's number. Measured on the 2026-11-01 exports by mapping each item row to its spell: stock v1.11p
about 1,417 -> 1,324 rows, Paradigm 1.9.1d about 2,094 -> 1,936 (7-8%). Rows of items that cast several
spells could not all be placed, so the real saving may be a little larger.

Decisions (2026-11-01): go ahead; add a `Spell` column for an exact join; clients follow EndCast through
`Spells` themselves; remove the `Use` / `OnHit` markers.

## R2.2 Schema

`Messages` gains `Spell` (Long), appended after `Effects`:

- **Set** on rows that belong to one spell: Cast rows (cast and start lines) and named Self rows (status,
  incoming and shield rows of a friendly player or item spell, 3.4 of the generic handoff).
- **0** on every other row: nameless Self rows (hostile spells' victim lines, fail lines, self-cast,
  traps, monster hits), Other rows (monster hits on others, death text, chase, monster entrances) and
  item damage-shield rows (item ability 137 is not a spell; those rows keep the item's name).
- So a number appears only where the row's name already identifies the spell. It adds the means to tell
  same-named spells apart (35 of 224 item spells on stock share a name), nothing about realm-wide lines.

`Items` is unchanged (`UseCode`, `UseSpells`, `HitSpells` as in section 2). `HasCurrentSchema` also requires
`Messages.Spell`, so an export made by an earlier 1.8.4 build (including the two of 2026-11-01) cannot be
updated in place.

## R2.3 Rows

| Row | Before | After |
|---|---|---|
| Cast (cast line, start line) | one per name: the spell (if players can cast it) and each item that casts it | one per spell: Name = the spell's name, Spell = its number |
| Status / incoming / shield of a friendly player or item spell | one per name, as above | one per spell, Name and Spell as above |
| Same, hostile spell | nameless | unchanged (Spell 0) |
| Item damage shield (item ability 137) | one per item | unchanged (Name = item, Spell 0) |
| Fail, self-cast, trap, monster, death, static | | unchanged (Spell 0) |

- **Which spells get named rows** is unchanged: player-castable, item-cast (use or on hit), or chained
  through EndCast from an item spell (3.1). Only the names change.
- **Merge key** for a row with a spell: Kind + Spell + Msg + EndMsg. Two spells with the same name and the
  same line stay two rows, one per number. Rows without a spell keep today's keys.
- **Markers** `Use` / `OnHit` are removed (3.2 withdrawn): how an item casts a spell is per item, in
  `UseSpells` / `HitSpells`.
- **Chains**: `UseSpells` / `HitSpells` stay the plain decode. A chained spell's rows exist under its own
  name and number; clients find them by following ability 151 in the `Spells` table's `Abil-n` columns,
  transitively. The docs give the rule.

What a client does for an item: rows with `Spell` in `UseSpells` (plus their EndCast chain) are the lines of
using it; rows with `Spell` in `HitSpells` (plus chains) are its on-hit lines; Self rows named after the item
with `DamageShield` are its damage shield.

## R2.4 Code

`modMMEMessages.bas`:
- Row store: `mRowSpell()`; `AddRow` takes the spell number (default 0) and keys rows with a spell by it;
  `FinalizeRows` writes `Spell`.
- `ProcessSpell`: one row name instead of a name list: the spell's name for Cast rows; for Self rows the
  spell's name when the target is friendly and players or items cast it, else "". The spell number goes
  with a non-empty name only. `BuildNameList` and `AddCastRows` go; `SpellStatusRow`, `SpellCastRow`,
  `SpellStartRow`, `SpellIncomingRow` take the single name; `SpellShieldRows` keeps its name list (items)
  and takes the spell number (0 for items).
- Item scan: `dictItemCast` becomes a set of spell numbers; `AddItemCaster`'s mode bookkeeping and
  `mdictItemMode` go; `ExpandItemChains` only adds chained spells to the set. `ItemSpellLists` stays.

`frmMME_Export.frm`: `CreateTables` appends `Spell` (adInteger) to `Messages`; `HasCurrentSchema` checks it;
the refusal message mentions the newest columns.

## R2.5 Docs and checks

- Generic handoff: 2.1/2.2 (`Spell`), 3.4/3.5/3.8 (one row per spell), 3.7/4.6 (markers gone), 4.1/4.4/4.9
  (names, joins, examples), 5.4 (modelling an item through `Spell`, with the EndCast rule), 6 (join on
  `Spell`, not on the name), 7 (new counts after the re-export).
- MegaMud doc: 9.1 says how the import fans rows out per item (rows whose `Spell` is in the item's
  `UseSpells` / `HitSpells` and their chains, copied under the item's name, which its name-based credit and
  `FindItem` need); marker row removed from section 7; O8 updated.
- Notes, README, changelog: entries rewritten (no markers; one row per spell; `Spell` column).
- `verify-mme-messages.ps1`: expected columns include `Spell`; the duplicate check keys on Spell too; marker
  checks replaced by: every Cast row has a `Spell` that exists in `Spells` and whose name is the row's
  `Name` (or the row is text-named because the spell's name is blank); rows with `Spell` 0 are never Cast;
  no Cast row is named after an item; the schema diff accepts `Spell` as new in `Messages`; spot checks:
  "You feel a little better." is one Cast row named `food` (stock spell 316), and the jade amulet's use
  spell has a Cast row with that `Spell`. INFO: items whose use spells have no Cast row.

## R2.6 Verification

- You rebuild and re-export stock v1.11p and Paradigm 1.9.1d; I run the verify script on both.
- Nothing lost: every (Kind, Msg, EndMsg) of the 2026-11-01 exports must still be in the new ones, apart
  from the marker change. I compare the old and new exports.
- The counts in the docs are updated from the new exports.

## R2.7 Risks

- Clients lose the item name on item lines and must join through `Items`. Name-matching clients (MegaMud)
  need the per-item fan-out in their import before item lines work again.
- Item lines show the game's spell names, some of them internal-looking (`weapon major bless`, `deed
  give`). They were already in the `Spells` table.
- The two 2026-11-01 exports cannot be updated in place by the new build.

## R2.8 Implementation notes (2026-11-01)

- `AddRow` takes `Optional nSpell`; rows with a spell are keyed `Kind#<spell>|Msg|EndMsg`. `AddRowEach`
  (shield rows) passes the number through. `AddCastRows`, `BuildNameList`, `AddItemCaster` and
  `mdictItemMode` are gone; `dictItemCast` is a set (value True).
- Cast rows always get the number, also when the spell's name is blank (the row is then text-named). Self
  rows get name and number only for a friendly player or item spell with a non-blank name.
- The verify script was run on a copy of the 2026-11-01 stock export converted to this shape by a script
  (1,310 rows; all checks pass) and on a copy with five planted faults (all caught). The real counts come
  from the re-export.

## R2.9 Result (2026-11-01)

Re-exported with the Revision 2 build; `verify-mme-messages.ps1` passes on both.

| Export | Before (per item) | After | Cast / Self / Other | Distinct lines lost |
|---|---|---|---|---|
| stock v1.11p | 1,417 | 1,311 | 415 / 477 / 419 | 0 (1,210 before and after) |
| Paradigm 1.9.1d | 2,094 | 1,930 | 511 / 669 / 750 | 0 against 1.9.1c; the 2 new lines are the EndCast chain rows |

- 12 stock rows repeat a line under a second spell of the same name (resist cold, barkskin, drain life...),
  as R2.3 intends; the simulation had merged those, hence its 1,310.
- The Paradigm spot check "dark force has no Cast row" failed because two weapons cast dark force spells
  (124, 5162) on hit, whose Cast rows were named after the weapons before. The check now allows Cast rows
  of item-cast dark force spells. The bane dart failure and the 2 `{?}` rows on Paradigm predate this
  branch (1.9.1c has them too).
- Both exports write `NMR-Log_MMEMessages.txt` into the same folder, so the later run replaced the other's
  log (the one left is stock's).

## R2.10 Branch audit (2026-11-01)

- `EndCastChain` looked spells up with a 4-byte Long key and a key length of 255 (`KEY_BUF_LEN`); every other
  spell lookup in NMR passes an Integer and `Len(...)`, so Btrieve is told the buffer's real size. Fixed the
  same way.
- The Messages item scan skipped blank-named items, while `Items.UseSpells` / `HitSpells` list every item's
  spells, so such an item's spells would not count as item-cast. The scan no longer needs a name (no
  stock or Paradigm item is affected).
- Verify script: same-named items now pool their use spells for the INFO count; new check that rows with a
  `Spell` belong to learnable, item-taught (ability 42) or item-cast spells only (Paradigm spell 24 is
  taught by a scroll but not marked Learnable, which NMR counts as player-castable, as documented).
- Docs: the notes' standard-formats intro and generic 4.4 still described per-item cast lines; corrected.
- Not run yet: updating an existing export (rename one item, update the MDB, check the other codes stay).
