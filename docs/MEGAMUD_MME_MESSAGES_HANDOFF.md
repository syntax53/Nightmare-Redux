# MME Messages and weapon verbs: MegaMud import notes (from Nightmare Redux)

Status: current as of 2026-10-19 for the MegaMud facts (NMR branch `261018-messages-fix1`, NMR 1.8.4;
MegaMud 3.x branch `@syntax53/feature/260926-sqlite-migration` as of commit 5e0b703e). Split on
2026-11-01: the client-neutral description of the table moved to `docs/MME_MESSAGES_HANDOFF.md`. Item uses
(section 9.1) added on NMR branch `261101-item-messages`.

This is the MegaMud companion to `docs/MME_MESSAGES_HANDOFF.md`. That document covers the table itself:
how NMR builds each row, the columns, kinds, names, tokens and effects, and how to match lines. **Read it
first.** This one covers only what is particular to MegaMud: how its matcher uses a message record, how a
row maps onto `MSGREC`, which `ME_*` / `MA_*` bits each NMR effect corresponds to (checked against the
stock `Messages.md`), and how to marry imported rows with the messages a user already has.

Read both before adding Messages (or item hit/miss verbs) to MegaMud's MME import (`MMEIMPORT.C`,
`MMEDB.C`, `MMEREC.C`).

Related: `MME_IMPORT_HANDOFF.md`, `MME_IMPORT_PLAN.md`, `GAMEDATA_HANDOFF.md` §3a (Messages.db),
`TEXT_TO_DB_HANDOFF.md` on the MegaMud side; `docs/MME_MESSAGES_HANDOFF.md` (the generic handoff),
`docs/mme_export_notes.txt` (the short client notes) and `docs/verify-mme-messages.ps1` on the NMR side.

Line numbers below for MegaMud source are from the 3.x branch. "Generic n.n" refers to a section of
`docs/MME_MESSAGES_HANDOFF.md`.

---

## 1. Where MegaMud stands

NMR's `Messages` table does the same job as MegaMud's Messages database, and its columns map onto `MSGREC`
almost one to one. NMR also adds six columns to `Items`: weapon verbs, and the spells an item casts when
used and on hit, with a command code for using it. MegaMud's import does not read any of them today (`MME_IMPORT_HANDOFF.md` lines 34-37: "Paths, Players, Messages, Rooms and Macros are not
imported"), and nothing in its plan schedules it.

---

## 2. Detecting it

- `Info.[NMR Version]` is `v1.8.4` or newer (`MmeSnapNmrVersion` gives 10804).
- **Probe for the table before selecting**: `MmeDescribe` on a missing table fails the whole read
  (`MME_IMPORT_HANDOFF.md`: there is no table-existence probe; a bad query reports SQLSTATE `42S02`).
  A version of 10804 or more implies the table exists (generic 2.1). Probe anyway.
- The extra `Items` columns are harmless to the current Items reader (it builds its column list from what
  exists).
- Columns, the `Chr(0)` empty value and the reserved-word brackets: generic 2.2 and 2.3.

---

## 3. Against the stock MegaMud data

A stock v1.11p export has 1,417 Messages rows (generic 7). Matched against the stock MegaMud
`Messages.md` (612 records), 502 records have an NMR counterpart. The rest are room sounds, MegaMud-only
records (lives warnings, `Cancel FUgram`), standard engine forms MegaMud parses elsewhere, text from older
game versions, and monster text NMR leaves out on purpose (death text of monsters worth under 101
experience). Keep all of those: they are not NMR's to replace.

NMR's one-row-per-name shape (generic 3.8) is the same as the stock `Messages.md`: `bless`, `chant`,
`glass orb` and `shimmering longsword` are separate records there too. NMR spell and item names (generic
4.4) are the same strings MegaMud's Spells and Items records hold after an MME import.

---

## 4. How MegaMud uses a message record (verified in source)

`MSGREC` (`MEGAREC.H:300`): `szName[31]`, `szMsg1[81]`, `szMsg2[81]`, `szReply[81]`, `wEffects`, `iAction`;
`schema_messages.sql` bounds them to 30 / 80 / 80 / 80 **bytes** (CP437).

- **First text, no `{`**: `ParseMessages` (`MESSAGES.C:564`) tests the line with `StrBegin` (the line must
  start with `szMsg1`), or `StrPos` (anywhere) when `ME_ANYPOS` is set. It is a prefix test: a stored
  "You are covered in acid" matches "You are covered in acid!".
- **First text with `{`** (`MS_PARSE1`, set at load): `MatchMessage` (`MESSAGES.C:3918`). The literal before
  the first token may sit anywhere in the line; each token is a lazy wildcard up to the next literal;
  `{source}`, `{target}`, `{1}`..`{5}` are captured, `{dmg}` is read with `atoi`; any other `{word}` is an
  uncaptured wildcard. Nothing after the pattern has to match, and literals compare case-sensitively.
- **Second text** (`ParseWornOff`, `MESSAGES.C:1302`): plain `StrBegin` / `StrPos`, **no wildcards**. A
  non-empty `szMsg2` makes the record go `MS_ACTIVE`, and the status bits it set stay on until that line.
  `MA_WAITITOFF` does nothing without a `szMsg2` (`MESSAGES.C:865`).
- **The name is functional**, not a label:
  - A damage line without `{source}` is credited to the pending combat spell only when the record's name
    equals that spell's name: `_stricmp` against the pending name (`MESSAGES.C:754`), then
    `strcmp (pCombatSpell->szName, pmm->szMsgMatch)` per cast slot (case-sensitive, `MESSAGES.C:771-811`).
    For weapon procs, `FindItem (pmm, pmm->szMsgMatch)` must be the current weapon (`MESSAGES.C:818`).
  - A non-damaging success ("A cool wind blows throughout the room!") counts as a successful combat cast
    when the name equals the pending spell's name (`_stricmp`, `MESSAGES.C:1030-1040`).
  - Bless timers: the record name is compared with each configured bless spell's name (`_stricmp`,
    `MESSAGES.C:964`). `AbilMonNoteSpellMsg (pmm, pMsg->szName)` feeds the ability monitors (`:975`).
- **Several records can fire on one line**: the loop normally goes on after a match (`continue`), and
  `szMsgMatch` ends up holding the last matching name. It stops early when a damage line is credited to the
  pending combat spell or the current weapon (`break`), and after a chase reply or a non-damaging combat
  success (`return`). So order matters only in those cases: the first record that credits the line wins.
- `szReply` is sent when the record matches; a chase record (`ME_CHASEMSG`) is only tested while chasing,
  its `{target}` must be the chased player or leader, and its reply is the follow command.

---

## 5. Mapping a `Messages` row onto `MSGREC`

| Field | Source / rule | Owner |
|---|---|---|
| `szName` | `Name`, cut to 30 bytes. Spell and item names fit (29 max). Text-derived names are longer on 303 stock rows; cut them the same way every time, because the cross-layer merge key is name + `szMsg1` (`GAMEDATA_HANDOFF.md`). | MME |
| `szMsg1` | `Msg` with tokens converted (section 6), cut to 80 bytes. Two stock rows are 81 and 82: cut at 80 and back off to before a partial `{...`; prefix matching keeps the cut text working. | MME |
| `szMsg2` | `EndMsg` on `Self` rows; empty otherwise. Never contains a token (verified on stock), which matters because `ParseWornOff` has no wildcards. Max 69 on stock. | MME |
| `szReply` | Never from NMR. Keep the stored value. | user |
| `wEffects` | Section 7, plus the stored `ME_DISABLED`, `ME_ANYPOS`, `ME_CHECKANY` bits. | mixed |
| `iAction` | `MA_CHECKROOM` for CheckRoom (and Death, if adopted); otherwise keep the stored value, `MA_IGNORE` for a new record. `MA_WAITITOFF`, `MA_RESTHPS`, `MA_RESTMANA`, `MA_RUN`, `MA_HANGUP` are user policy: never set them from NMR. | mixed |

Kind decides the shape: Cast and Other rows never have a `szMsg2`; only Self rows with an `EndMsg` do.

---

## 6. Tokens

What each token means: generic 3.3 and 4.3.

| NMR | MegaMud | Note |
|---|---|---|
| `{source}` | `{source}` | |
| `{target}` | `{target}` | |
| `{dmg}` | `{dmg}` | Read with `atoi`, which also covers the monster-hit-on-another line where the engine prints the amount through a string placeholder. |
| `{attacker}` | `{target}` | Only on damage-shield rows: the creature your shield or spikes hit (10 stock rows). MegaMud has no `{attacker}`; as-is it would still match, uncaptured, but `{target}` is the damage's target. |
| `{?}` | leave, or `{1}` | A placeholder NMR could not identify (0 on stock). Matches as a wildcard either way. |

**The "you" forms.** NMR drops a Self line that is the "you" form of a Cast or Other line (generic 3.8);
clients are told to accept "you"/"your" for `{target}`. In MegaMud a `{target}` wildcard does match "you",
but "{target}'s legs" does not match "your legs" (the literal "'s" is compared). 16 stock rows contain
"{target}'s". If MegaMud should react when it is the one affected, add a second record with "{target}'s"
replaced by "your".

---

## 7. NMR effects → `ME_*` / `MA_*`

The vocabulary and where each effect comes from: generic 3.7 and 4.6. "Stock" below counts how the stock
`Messages.md` flags the records that match an NMR row with that effect (612 records cross-checked; 502
matched).

**Rules first:**
- **Cast rows: `wEffects` 0, always.** Their effects describe the target. Stock: 110 of 110 matched Cast
  rows are 0. MegaMud credits them by name (section 4).
- **Other rows: no status bits.** They are about someone else. Only `ME_CHASEMSG` and the actions below.
- **Status bits (`ME_STATUS` in `MESSAGES.C`) only on Self rows with an `EndMsg`.** Without a `szMsg2`
  nothing turns the bit off again until a reset or `stat`.

| NMR effect (on a Self row with EndMsg unless noted) | MegaMud | Stock evidence |
|---|---|---|
| Poisoned | `ME_POISONED` | 3 of 3 |
| Blinded | `ME_BLINDED` | 7 of 11 |
| Held | `ME_NOMOVE` | 16 of 19 (5 also `MA_WAITITOFF`: user policy) |
| Sleep | `ME_NOMOVE` + `ME_NOATTACK` | stock `sleep` is 0x0030 |
| Confused | `ME_CONFUSED` | 8 of 23. The rest are stun-like: `stunned` is 0x1030 (`ME_NOMOVE` + `ME_NOATTACK` + `ME_COMBATOFF`). NMR cannot tell a stun from a confuse (both use the same game ability), so treat a stored record's bits as better than the default. |
| Fear | owner decision | No consistent precedent: `fear` and `roar` are 0x0010 (`ME_NOMOVE`), `terror` 0x0022 (`ME_CONFUSED` + `ME_NOATTACK`), `banshee's wail` 0x0002. |
| DamageOverTime | `ME_LOSINGHPS` | 20 of 41 (poisons also carry Poisoned → `ME_POISONED`). `ME_DISEASED` is not used by any stock record. |
| HPRegen | `ME_REGENING` | 6 of 10. The REGEN tracker (`REGEN_HANDOFF.md`) watches this; `bRegenerating` is a status display now. |
| HealOverTime | none, or `ME_REGENING` when HPRegen is also present | 3 of 14 |
| ManaRegen | `ME_FLUXING` | 6 of 13 |
| Fail (Self, no EndMsg) | `ME_FAILED` | `Afraid` and `fumble` are 0x2002 |
| Trap (Self, no EndMsg) | `ME_FAILED` | 11 of 11. MegaMud's trap records are disarm failures, and `ME_FAILED` ends a disarm attempt (`bDisarmed`). NMR's trap rows also include the lines printed when a trap fires as you walk through; `ME_FAILED` on those also clears the current action. |
| Chase (Other) | `ME_CHASEMSG` | 78 of 78. NMR does not export the follow command (`szReply`, e.g. "go alley"); see 8.4. |
| CheckRoom (Other) | `iAction` = `MA_CHECKROOM`, `wEffects` 0 | 31 of 32 |
| Death (Other) | suggest `iAction` = `MA_CHECKROOM` | stock `obsidian statue death` / `Statue death` are 0000 with action 1 |
| Mute, Drain, DamageShield, Damage, Heal, ManaHeal, ManaSteal, Cure, RoomSpell, Monster | 0 | Mute 4 of 4, Drain 10 of 10, DamageShield 7 of 7, one-off Damage 115 of 128 |
| Use, OnHit (Cast rows under an item's name) | 0 | New in NMR, no stock counterpart. They say how the item cast the spell; MegaMud ties the line to the item by its name (section 4). |

`ME_CHECKANY` (conversation lines) and `ME_ANYPOS` have no NMR counterpart: NMR text is the whole line from
its start, so a new record needs neither.

---

## 8. Marrying imported rows with existing messages

Messages do not fit the current import design: the five imported types are keyed by `Number` and updated
record by record, while Messages are keyed by row order, merged across layers on name + `szMsg1`, and saved
by rewriting the whole layer (`GAMEDATA_HANDOFF.md` §3a). NMR's `ID` cannot help: it is renumbered by every
export. So the import needs its own rules. Suggested ones:

### 8.1 Layer

Write only `Default\Messages.db`, like the other types. `Chars\All`, `BBS\<bbs>` and `Chars\<char>` keep
overriding in the merged view.

### 8.2 Match key

1. **Canonical text**, for both sides: tokens (`{source}`, `{target}`, `{dmg}`, `{attacker}`, `{1}`..`{5}`,
   `{?}`) replaced by one marker, case folded, runs of spaces collapsed, trailing `.` / `!` and spaces
   removed. Stock MegaMud text often drops the final punctuation ("You feel lucky") and is sometimes a
   fragment ("flays you viciously").
2. **Text match**: equal canonical `szMsg1`, or the stored canonical `szMsg1` is a prefix of NMR's (a
   substring when the stored record has `ME_ANYPOS`). When both sides have a second text, the canonical
   `szMsg2` must match the same way: one start line can end in different ways.
3. **Name**: for a row named after a spell or item (generic 4.4), the stored name must also match,
   case-insensitively. For a realm-wide row, ignore the name: stock names are labels too ("monster entry",
   "trap disarm", "chase (alley)").

### 8.3 What to take

- **A record the import created earlier**: refresh every MME-owned field.
- **A record the import did not create** (shipped stock or user-made, matched by 8.2): keep its text,
  `szReply`, `iAction` and user bits. A stored prefix that still matches is often deliberately broad.
  Taking NMR's exact text gives nothing MegaMud needs. Two exceptions are worth considering:
  - Set the name to NMR's spell or item name when the stored record is the only match and its name differs
    only by case. The combat-credit compare is case-sensitive.
  - Fill in missing status bits from section 7, but only on records whose effects are 0.
- **No match**: insert a new record. Append it after the existing Default records. Order matters only
  where the scan stops early (section 4): the first record that credits a damage line to a spell wins.
  Within one layer, duplicates of name + `szMsg1` collapse on load, with the last content winning and the
  first position kept.

### 8.4 Rows to treat specially

- **Chase rows** are useless without a follow command, which NMR does not have. Import a Chase row only
  when it marries a stored chase record, and keep that record's reply. Alternatively, insert it with
  `ME_DISABLED` so the user can add the command. MegaMud only sends a non-empty reply, so an empty one is
  inert either way.
- **Chase and CheckRoom rows are realm-independent**: NMR writes the same 107 lines in every export, in the
  stock game's wording. Most already exist in the stock Default data.
- **Monster-only status rows** are nameless (realm-wide text). The same text may also exist under a player
  spell's name, e.g. "You feel resistant to cold!" as `resist cold`, while stock MegaMud also has it under
  `ice shield`. Both are legitimate.

### 8.5 Knowing what the import owns (owner decision)

The `messages` table has no column to mark a record as imported, and `seq` changes on every save. Without
a marker, a re-import cannot tell its own stale rows from stock or user rows, and the orphan-delete step the
other types have cannot exist. The options:

- **A side table in `Default\Messages.db`** (for example `mme_owned(name TEXT, msg1 TEXT)`) that each
  import rewrites. Check first that the loader reads only `messages`, and that the whole-layer save deletes
  only from `messages`. This is recommended.
- **The unused `wEffects` bit 0x0800** as an "imported" marker. It is cheaper, but it has to be masked out
  of every test that compares the whole word, e.g. `wEffects == 0 && pMsg->wEffects != 0` in the
  `MA_WAITITOFF` case, the list icons and the filters.
- **Accept no ownership.** Then the import only inserts what doesn't marry and never deletes.

With ownership, offer the deletion of owned records that NMR no longer exports, the same way the other
types do: one box, default No.

### 8.6 Plumbing gaps on the MegaMud side

These come from the import's own handoff:

- A sixth type needs `MMEIMP_TYPES` raised, `s_aImpType` extended, and `MMEREC.C:293`'s hard-coded
  `"Races,Classes,Items,Monsters,Spells"` (the "Only ... were updated" test) adjusted.
- Messages are saved by whole-layer rewrite. That makes them a load / modify / save of the full Default
  list, not per-record `CDB_Update`.
- `CheckMessageParserConflicts (pmm)` already runs on a messages reload (`MMEIMPORT.C:4012`). Run it after
  an import too: imported records can shadow stored ones.

---

## 9. Item columns (`Items`)

What the columns hold and how NMR fills them: generic 5.

### 9.1 Item uses

MegaMud keeps an item use as a spell record named after the item. NMR does not write such records: a made-up
spell number could not stay stable across exports, and the game spell keeps its own name and code for the
classes and monsters that cast it. The `Items` columns give the import what it needs to build them itself:

- **Which items**: every item with a `UseCode`. Items that only cast on hit have none and need no record.
- **Name and command**: `Items.Name` (29 characters, fits `szName`) and `UseCode` (`#` plus up to 4
  letters, 5 characters in all). Items with exactly the same name share a code.
- **Stats**: the `Spells` row of the first number in `UseSpells` (rarely more than one; food and drink
  often have two).
- **Key**: `Items.Number`, not the code (a code changes when the item is renamed). MegaMud assigns its own
  record number and keeps an item-number map, like the ownership options in 8.5, so a re-import updates
  the same record and the user's edits to it survive.
- **Messages**: the item's Cast rows already carry the item's name, so the name-based credit in section 4
  works once the item record has the same name. `Use` in `Effects` marks the lines that confirm a use.
- **On-hit spells** (`HitSpells`) need no record: their Cast rows are named after the weapon, which
  `FindItem (pmm, pmm->szMsgMatch)` already checks against the current weapon (section 4).

Whether MegaMud's spell list should hold these records, and how they show in its UI, is an owner decision
(O8).

### 9.2 Weapon verbs

MegaMud has nowhere to put them: `ITEMREC` (268 bytes) and `schema_items.sql` have no verb field, and
`ParseHitting` recognises your own hits by their shape, with monster verbs in `ATTACKS.LST`. Storing them
means widening the record again (schema version 3, `MdConvert`'s frozen layouts). That is only worth doing
if a feature needs it, such as telling which hand or weapon landed a blow after a swap. Owner decision;
skipping them loses nothing MegaMud uses today.

---

## 10. Gotchas for the MegaMud import

The client-neutral ones (`Chr(0)` as empty, `ID` not stable, realm-wide names, GreaterMUD exports) are in
generic 6.

- **Do not set status bits from Cast or Other rows** (section 7). A Cast row's "Blinded" means the target
  is blinded.
- **The loader skips rows with no name or no first text** (`GAMEDATA_HANDOFF.md` §3a). NMR never writes an
  empty `Name`, but `Msg` can be empty (a status message with only a wear-off line, generic 2.3) or become
  empty after conversion. Do not insert those.
- **Byte limits are CP437 bytes**, not UTF-16 characters. Jet hands the text back as Unicode. The stock
  export has no character outside ASCII; on a realm that has some, convert to the client's bytes before
  checking lengths.
- **Two NMR rows can marry one stored record.** A stock record that is a fragment is a prefix of several
  NMR lines. Decide whether the stored record absorbs them (then insert nothing) or stays and the rest are
  inserted. Inserting is harmless: all of them fire.
- **NMR leaves out the engine's standard forms** (generic 3.6), except a player spell's own cast line.
  MegaMud parses those in code, so nothing changes there.
- **Engine strings that are not in the data**: "You have N lives left." (`check_kill_user`), "You feel your
  life-force ebbing." (`load_player`) and "You look around stupidly and do nothing!" (`check_confusion`).
  NMR never exports them. Keep MegaMud's own records for them.

---

## 11. Open decisions for the owner

| # | Question | Recommendation |
|---|---|---|
| O1 | Import Messages at all, or keep them hand-maintained? | Import, into Default, with the 8.2 match rules. |
| O2 | How does the import know which Default records it owns (8.5)? | A side table in `Default\Messages.db`. |
| O3 | On a married stock or user record, change nothing, or also adopt NMR's case-corrected name and fill in empty status bits? | Adopt the name only when it differs just by case; fill bits only when they are 0. |
| O4 | Fear mapping (no stock consensus). | `ME_NOATTACK`, or leave 0 and let the user set it. |
| O5 | Death rows: `MA_CHECKROOM`, or ignore? | `MA_CHECKROOM` (stock precedent). |
| O6 | Chase rows with no stored counterpart: skip, or insert disabled? | Skip. |
| O7 | "your" variants for the 16 "{target}'s" rows (6). | Add them. |
| O8 | Item uses and weapon verbs (9). | Build item-use records from `UseCode` / `UseSpells`, keyed by item number. Skip the verbs until a feature needs them. |

---

## 12. Checking an implementation

- **On stock data**: `TestData/Stock111p/Messages.db` holds the stock records. A stock v1.11p NMR export
  has 1,417 Messages rows, and about 500 of the 612 stock records should marry. NMR's audit of that pair is
  in the NMR repository (`docs/megamud-stock-audit.txt`, a local file, with per-record message numbers).
- **Idempotence**: importing the same MDB twice must leave `Default\Messages.db` byte-for-byte unchanged,
  the same rule the other types follow.
- **Spot checks** (all stock v1.11p; more context in generic 4.9):

| Name | Kind / effects | Record |
|---|---|---|
| `resist cold` | Self | "You feel resistant to cold!" / "The effects of resist cold wear off!" |
| `black wind` | Cast | "A cool wind blows throughout the room!" |
| `annointed hands` | Cast | "You lay hands on {target}." |
| `agony` | Cast | "You cast agony on {target} for {dmg} damage!" |
| `jade amulet` | Cast, `Use` | "You rub your jade amulet!" (the item `jade amulet` has a `UseCode`) |
| `spiked shield` | Self | "The shield spike stabs {attacker} for {dmg} damage!" (becomes `{target}`) |
| (realm-wide) | Chase | "{target} slips into the dark alley." |
| (realm-wide) | CheckRoom | "A moss zombie rises from the ground!" |

- NMR's own check of an export is `docs/verify-mme-messages.ps1` (32-bit PowerShell, Jet).
