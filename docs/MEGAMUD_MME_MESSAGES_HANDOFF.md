# MME Messages and weapon verbs: MegaMud import handoff (from Nightmare Redux)

Status: current as of 2026-10-19. NMR branch `261018-messages-fix1` (NMR 1.8.4); MegaMud 3.x branch
`@syntax53/feature/260926-sqlite-migration` as of commit 5e0b703e.

Read this before adding Messages (or item hit/miss verbs) to MegaMud's MME import (`MMEIMPORT.C`,
`MMEDB.C`, `MMEREC.C`).

Related: `MME_IMPORT_HANDOFF.md`, `MME_IMPORT_PLAN.md`, `GAMEDATA_HANDOFF.md` §3a (Messages.db),
`TEXT_TO_DB_HANDOFF.md` on the MegaMud side; `docs/mme_export_notes.txt` (the general client notes for the
Messages table) and `docs/verify-mme-messages.ps1` on the NMR side.

Line numbers below for MegaMud source are from the 3.x branch.

---

## 1. The feature in one paragraph

NMR 1.8.4 adds a `Messages` table to the MME export: the game lines a client needs to recognise (status
effects with their start and wear-off text, the cast lines of every player spell and spell-casting item,
non-standard damage lines, traps, monster death text, plus fixed chase and monster-entrance lines). It
is built from the realm's own message file, spells, items, monsters, rooms and textblocks, so it follows
a customised realm where a shipped `Messages.md` cannot. It is the same job MegaMud's Messages database
does, and the columns map onto `MSGREC` almost one to one. NMR also adds three weapon-verb columns to
`Items`. MegaMud's import does not read either today (`MME_IMPORT_HANDOFF.md` lines 34-37: "Paths,
Players, Messages, Rooms and Macros are not imported"), and nothing in its plan schedules it. This
document says what the data means, how it maps onto `MSGREC`, which `ME_*` / `MA_*` bits each NMR effect
corresponds to (checked against the stock `Messages.md`), and how to marry imported rows with the
messages a user already has.

---

## 2. What the export contains

### 2.1 Detecting it

- `Info.[NMR Version]` is `v1.8.4` or newer (`MmeSnapNmrVersion` gives 10804).
- The table is `Messages`. **Probe for it before selecting**: `MmeDescribe` on a missing table fails the
  whole read (`MME_IMPORT_HANDOFF.md`: there is no table-existence probe; a bad query reports SQLSTATE
  `42S02`). Exports made by NMR 1.8.3 and older do not have it, and an 1.8.4 NMR refuses to update such an
  old MDB in place, so a version of 10804 or more implies the table exists. Probe anyway.
- `Items` gains `HitVerbs`, `HitVerbsOther`, `MissVerbs`, appended after `References`. Extra columns are
  harmless to the current Items reader (it builds its column list from what exists).

### 2.2 Columns (`Messages`)

| Column | Jet type | Size | Meaning |
|---|---|---|---|
| `ID` | Long | | Row number 1..n in NMR's write order. **Not a stable key**: it is renumbered by every export. |
| `Name` | Text | 100 | Spell or item name, monster name (death rows), or a name made from the line's own text. Never empty. |
| `Kind` | Text | 12 | `Cast`, `Self` or `Other` (2.3). |
| `Msg` | Text | 100 | The line, literal except for tokens (section 5). |
| `EndMsg` | Text | 100 | Wear-off line. Only on `Self` rows; its presence means the effect lasts until that line. |
| `Effects` | Text | 255 | Comma list (section 6). |

All columns are Text, not Memo, so the reader's normal binding works. Bracket `[Name]` and `[Msg]`.
Like every MME table, an empty string is written as a single `Chr(0)`: treat a value of one NUL as empty.
The index is `pkMessages` on `ID`.

### 2.3 Kinds, from the player's seat

- **Cast**: confirmation that MY spell or item worked. Every player-castable spell (Learnable per NMR's
  cross-referencing) and every item that casts a spell has its cast line here, under the spell or item
  name, whatever its wording ("You cast agony on {target} for {dmg} damage!", "A cool wind blows
  throughout the room!", "You rub your jade amulet!"). When that line is a standard engine form or empty,
  the spell's start line (ability 120, printed only once the spell lands) is exported too ("You lay
  hands on {target}."). Effects on a Cast row describe what the spell does to its **target**, not to you.
- **Self**: something happened to ME. With an `EndMsg` it is a tracked effect: `Msg` when it starts (and
  in `stat`), `EndMsg` when it wears off. Without one it is a one-off line: non-standard incoming damage,
  a failed action, a non-standard monster hit, a trap, my damage shield or item spikes hitting back, or a
  spell a textblock / cast exit / spell trap casts on me. A Self row carries a spell or item name only
  when the player can be the one affected by their own spell or item (self, friendly or party targets).
- **Other**: something about someone else: a monster's non-standard hit on another player, a monster's
  death text (`Name` = the monster), a chase line ("{target} slips into the dark alley."), or a monster
  appearing ("A moss zombie rises from the ground!").

### 2.4 Names

- Each player spell and each casting item gets **its own row under its own name**, even when several
  print the same text. "You feel lucky!" / "The effects of bless wear off!" is exported once as `bless`,
  once as `chant`, once per item that casts bless. This is the same shape as the stock `Messages.md`
  (`bless`, `chant`, `glass orb`, `shimmering longsword` are separate records there too).
- Rows collapse only when Kind, Name, Msg and EndMsg are all equal; their Effects are merged.
- Every other row is realm-wide and named after its own text, shortened ("The {source} claws you with its
  pincers for {dmg} damage!" is named "claws you with its pincers"). Those names are labels, not keys.
- **How to tell the two apart**: a row is named after a spell or item when its `Name` equals a `Spells.Name`
  or an `Items.Name` in the same export (all 518 Cast rows and 150 of 475 Self rows on stock). NMR spell
  and item names are the game's own (up to 29 characters, game case), the same strings MegaMud's Spells and
  Items records hold after an MME import.

### 2.5 Size on stock v1.11p

| | Rows |
|---|---|
| Cast | 518 |
| Self | 475 (272 with an EndMsg) |
| Other | 419 (213 death texts, 78 chase lines, 29 monster entrances) |
| Total | 1,412 |

Matched against the stock MegaMud `Messages.md` (612 records), 502 records have an NMR counterpart. The
rest are room sounds, MegaMud-only records (lives warnings, `Cancel FUgram`), standard engine forms
MegaMud parses elsewhere, text from older game versions, and monster text NMR leaves out on purpose
(death text of monsters worth under 101 experience). Keep all of those: they are not NMR's to replace.

---

## 3. How MegaMud uses a message record (verified in source)

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

## 4. Mapping a `Messages` row onto `MSGREC`

| Field | Source / rule | Owner |
|---|---|---|
| `szName` | `Name`, cut to 30 bytes. Spell and item names fit (29 max). Text-derived names are longer on 303 stock rows; cut them the same way every time, because the cross-layer merge key is name + `szMsg1` (`GAMEDATA_HANDOFF.md`). | MME |
| `szMsg1` | `Msg` with tokens converted (section 5), cut to 80 bytes. Two stock rows are 81 and 82: cut at 80 and back off to before a partial `{...`; prefix matching keeps the cut text working. | MME |
| `szMsg2` | `EndMsg` on `Self` rows; empty otherwise. Never contains a token (verified on stock), which matters because `ParseWornOff` has no wildcards. Max 69 on stock. | MME |
| `szReply` | Never from NMR. Keep the stored value. | user |
| `wEffects` | Section 6, plus the stored `ME_DISABLED`, `ME_ANYPOS`, `ME_CHECKANY` bits. | mixed |
| `iAction` | `MA_CHECKROOM` for CheckRoom (and Death, if adopted); otherwise keep the stored value, `MA_IGNORE` for a new record. `MA_WAITITOFF`, `MA_RESTHPS`, `MA_RESTMANA`, `MA_RUN`, `MA_HANGUP` are user policy: never set them from NMR. | mixed |

Kind decides the shape: Cast and Other rows never have a `szMsg2`; only Self rows with an `EndMsg` do.

---

## 5. Tokens

| NMR | MegaMud | Note |
|---|---|---|
| `{source}` | `{source}` | The caster or attacking monster as the game prints it, including any "The". |
| `{target}` | `{target}` | |
| `{dmg}` | `{dmg}` | An integer. NMR also writes it where the engine prints the amount through a string placeholder. |
| `{attacker}` | `{target}` | Only on damage-shield rows: the creature your shield or spikes hit (10 stock rows). MegaMud has no `{attacker}`; as-is it would still match, uncaptured, but `{target}` is the damage's target. |
| `{?}` | leave, or `{1}` | A placeholder NMR could not identify (0 on stock). Matches as a wildcard either way. |

- Spell names are written literally ("You fire a magic missile at {target} for {dmg} damage!").
- Area, room and party spells print nothing where `{target}` would be, and NMR's text already reflects that
  ("You cast greater healing rain on your party, healing {dmg} damage!").
- **The "you" forms.** NMR drops a Self line that is the "you" form of a Cast or Other line
  ("{target}" read as "you", "{target}'s" as "your"); clients are told to accept "you"/"your" for
  `{target}`. In MegaMud a `{target}` wildcard does match "you", but "{target}'s legs" does not match
  "your legs" (the literal "'s" is compared). 16 stock rows contain "{target}'s". If MegaMud should react
  when it is the one affected, add a second record with "{target}'s" replaced by "your".

---

## 6. NMR effects → `ME_*` / `MA_*`

The vocabulary is `Poisoned Confused Blinded Held Fear Mute Sleep Damage DamageOverTime Drain Heal
HealOverTime HPRegen ManaRegen ManaHeal ManaSteal DamageShield Cure Fail RoomSpell Monster Trap Death Chase
CheckRoom`. "Stock" below counts how the stock `Messages.md` flags the records that match an NMR row with
that effect (612 records cross-checked; 502 matched).

**Rules first:**
- **Cast rows: `wEffects` 0, always.** Their effects describe the target. Stock: 110 of 110 matched Cast
  rows are 0. MegaMud credits them by name (section 3).
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
| Chase (Other) | `ME_CHASEMSG` | 78 of 78. NMR does not export the follow command (`szReply`, e.g. "go alley"); see 7.4. |
| CheckRoom (Other) | `iAction` = `MA_CHECKROOM`, `wEffects` 0 | 31 of 32 |
| Death (Other) | suggest `iAction` = `MA_CHECKROOM` | stock `obsidian statue death` / `Statue death` are 0000 with action 1 |
| Mute, Drain, DamageShield, Damage, Heal, ManaHeal, ManaSteal, Cure, RoomSpell, Monster | 0 | Mute 4 of 4, Drain 10 of 10, DamageShield 7 of 7, one-off Damage 115 of 128 |

`ME_CHECKANY` (conversation lines) and `ME_ANYPOS` have no NMR counterpart: NMR text is the whole line from
its start, so a new record needs neither.

---

## 7. Marrying imported rows with existing messages

Messages do not fit the current import design: the five imported types are keyed by `Number` and updated
record by record, while Messages are keyed by row order, merged across layers on name + `szMsg1`, and saved
by rewriting the whole layer (`GAMEDATA_HANDOFF.md` §3a). NMR's `ID` cannot help: it is renumbered by every
export. So the import needs its own rules. Suggested ones:

### 7.1 Layer

Write only `Default\Messages.db`, like the other types. `Chars\All`, `BBS\<bbs>` and `Chars\<char>` keep
overriding in the merged view.

### 7.2 Match key

1. **Canonical text**, for both sides: tokens (`{source}`, `{target}`, `{dmg}`, `{attacker}`, `{1}`..`{5}`,
   `{?}`) replaced by one marker, case folded, runs of spaces collapsed, trailing `.` / `!` and spaces
   removed. Stock MegaMud text often drops the final punctuation ("You feel lucky") and is sometimes a
   fragment ("flays you viciously").
2. **Text match**: equal canonical `szMsg1`, or the stored canonical `szMsg1` is a prefix of NMR's (a
   substring when the stored record has `ME_ANYPOS`). When both sides have a second text, the canonical
   `szMsg2` must match the same way: one start line can end in different ways.
3. **Name**: for a row named after a spell or item (2.4), the stored name must also match,
   case-insensitively. For a realm-wide row, ignore the name: stock names are labels too ("monster entry",
   "trap disarm", "chase (alley)").

### 7.3 What to take

- **A record the import created earlier**: refresh every MME-owned field.
- **A record the import did not create** (shipped stock or user-made, matched by 7.2): keep its text,
  `szReply`, `iAction` and user bits. A stored prefix that still matches is often deliberately broad.
  Taking NMR's exact text gives nothing MegaMud needs. Two exceptions are worth considering:
  - Set the name to NMR's spell or item name when the stored record is the only match and its name differs
    only by case. The combat-credit compare is case-sensitive.
  - Fill in missing status bits from section 6, but only on records whose effects are 0.
- **No match**: insert a new record. Append it after the existing Default records. Order matters only
  where the scan stops early (section 3): the first record that credits a damage line to a spell wins. Within one layer, duplicates of name +
  `szMsg1` collapse on load, with the last content winning and the first position kept.

### 7.4 Rows to treat specially

- **Chase rows** are useless without a follow command, which NMR does not have. Import a Chase row only
  when it marries a stored chase record, and keep that record's reply. Alternatively, insert it with
  `ME_DISABLED` so the user can add the command. MegaMud only sends a non-empty reply, so an empty one is
  inert either way.
- **Chase and CheckRoom rows are realm-independent**: NMR writes the same 107 lines in every export, in the
  stock game's wording. Most already exist in the stock Default data.
- **Monster-only status rows** are nameless (realm-wide text). The same text may also exist under a player
  spell's name, e.g. "You feel resistant to cold!" as `resist cold`, while stock MegaMud also has it under
  `ice shield`. Both are legitimate.

### 7.5 Knowing what the import owns (owner decision)

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

### 7.6 Plumbing gaps on the MegaMud side

These come from the import's own handoff:

- A sixth type needs `MMEIMP_TYPES` raised, `s_aImpType` extended, and `MMEREC.C:293`'s hard-coded
  `"Races,Classes,Items,Monsters,Spells"` (the "Only ... were updated" test) adjusted.
- Messages are saved by whole-layer rewrite. That makes them a load / modify / save of the full Default
  list, not per-record `CDB_Update`.
- `CheckMessageParserConflicts (pmm)` already runs on a messages reload (`MMEIMPORT.C:4012`). Run it after
  an import too: imported records can shadow stored ones.

---

## 8. Weapon verbs (`Items`)

| Column | Text | Example (`longsword`) |
|---|---|---|
| `HitVerbs` | `"|"`-separated verbs for "You `<verb>` X for N damage!" | `slash|impale|hack` |
| `HitVerbsOther` | for "`<player>` `<verb>` X for N damage!" | `slashes|impales|hacks` |
| `MissVerbs` | for the player's misses | `swing at` |

They are filled only for weapons (`ItemType` = 1, any `WeaponType`: 0 is 1H Blunt); other items have `Chr(0)`. MegaMud has nowhere to put
them: `ITEMREC` (268 bytes) and `schema_items.sql` have no verb field, and `ParseHitting` recognises your
own hits by their shape, with monster verbs in `ATTACKS.LST`. Storing them means widening the record again
(schema version 3, `MdConvert`'s frozen layouts). That is only worth doing if a feature needs it, such as
telling which hand or weapon landed a blow after a swap. Owner decision; skipping them loses nothing
MegaMud uses today.

---

## 9. Gotchas

- **A blank `Chr(0)` is not an empty string** to a naive reader. `Name` is never blank, but `EndMsg` and
  `Effects` usually are.
- **Do not set status bits from Cast or Other rows** (section 6). A Cast row's "Blinded" means the target
  is blinded.
- **The loader skips rows with no name or no first text** (`GAMEDATA_HANDOFF.md` §3a). NMR never writes an
  empty `Name`, but a `Msg` that becomes empty after conversion must not be inserted.
- **Byte limits are CP437 bytes**, not UTF-16 characters. Jet hands the text back as Unicode. The stock
  export has no character outside ASCII; on a realm that has some, convert to the client's bytes before
  checking lengths.
- **Two NMR rows can marry one stored record.** A stock record that is a fragment is a prefix of several
  NMR lines. Decide whether the stored record absorbs them (then insert nothing) or stays and the rest are
  inserted. Inserting is harmless: all of them fire.
- **NMR leaves out the engine's standard forms** ("You cast X on Y", "<X> casts Y on you", "You take N
  damage", "... you for N damage!", misses and so on), except a player spell's own cast line. MegaMud
  parses those in code, so nothing changes there.
- **Engine strings that are not in the data**: "You have N lives left." (`check_kill_user`), "You feel your
  life-force ebbing." (`load_player`) and "You look around stupidly and do nothing!" (`check_confusion`).
  NMR never exports them. Keep MegaMud's own records for them.
- **Realm differences.** On GreaterMUD exports (`Info.Legit` = 2), the placeholder order follows that
  engine's reading of the message-style byte. The exported text is already final, so this changes nothing
  for an importer.

---

## 10. Open decisions for the owner

| # | Question | Recommendation |
|---|---|---|
| O1 | Import Messages at all, or keep them hand-maintained? | Import, into Default, with the 7.2 match rules. |
| O2 | How does the import know which Default records it owns (7.5)? | A side table in `Default\Messages.db`. |
| O3 | On a married stock or user record, change nothing, or also adopt NMR's case-corrected name and fill in empty status bits? | Adopt the name only when it differs just by case; fill bits only when they are 0. |
| O4 | Fear mapping (no stock consensus). | `ME_NOATTACK`, or leave 0 and let the user set it. |
| O5 | Death rows: `MA_CHECKROOM`, or ignore? | `MA_CHECKROOM` (stock precedent). |
| O6 | Chase rows with no stored counterpart: skip, or insert disabled? | Skip. |
| O7 | "your" variants for the 16 "{target}'s" rows (5). | Add them. |
| O8 | Weapon verbs (8). | Skip until a feature needs them. |

---

## 11. Checking an implementation

- **On stock data**: `TestData/Stock111p/Messages.db` holds the stock records. A stock v1.11p NMR export
  has 1,412 Messages rows, and about 500 of the 612 stock records should marry. NMR's audit of that pair is
  in the NMR repository (`docs/megamud-stock-audit.txt`, a local file, with per-record message numbers).
- **Idempotence**: importing the same MDB twice must leave `Default\Messages.db` byte-for-byte unchanged,
  the same rule the other types follow.
- **Spot checks** (all stock v1.11p):

| Name | Kind / effects | Record |
|---|---|---|
| `resist cold` | Self | "You feel resistant to cold!" / "The effects of resist cold wear off!" |
| `black wind` | Cast | "A cool wind blows throughout the room!" |
| `annointed hands` | Cast | "You lay hands on {target}." |
| `agony` | Cast | "You cast agony on {target} for {dmg} damage!" |
| `spiked shield` | Self | "The shield spike stabs {attacker} for {dmg} damage!" (becomes `{target}`) |
| (nameless) | Chase | "{target} slips into the dark alley." |
| (nameless) | CheckRoom | "A moss zombie rises from the ground!" |

- NMR's own check of an export is `docs/verify-mme-messages.ps1` (32-bit PowerShell, Jet).
