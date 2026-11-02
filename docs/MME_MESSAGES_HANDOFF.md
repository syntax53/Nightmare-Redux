# The MME `Messages` table and item columns: client handoff (from Nightmare Redux)

Status: current as of NMR 1.8.4 (2026-11-01). The code is `modMMEMessages.bas` (the Messages pass, and the
use / on-hit decoding of item spells) and `frmMME_Export.frm` (table creation, the `Items` columns, the call into
the pass). Design notes: `PLAN-mme-messages-export.md` and `PLAN-mme-item-uses.md`.

This is for anyone writing a client that reads an NMR MME export and wants to recognise game lines or
use items: a mapper, a bot, a log parser, or MME itself. It says how NMR builds each row from the realm's
data, what the columns mean, and how to match and use the rows. It assumes no particular client.

Client-specific notes: MegaMud, see `docs/MEGAMUD_MME_MESSAGES_HANDOFF.md` (mapping onto its message
records and `ME_*` / `MA_*` flags, and merging with its stored messages). The short reference version of
this document is the Messages, Items and "Standard formats" sections of `docs/mme_export_notes.txt`.

---

## 1. What the table is

`Messages` holds the game lines a client may need to recognise and that it cannot work out on its own:
status effects with their start and wear-off text, the cast line of every spell that players or items
cast, damage lines in non-standard wording, traps, damage-shield lines, monster death text, plus a fixed
set of chase and monster-entrance lines. NMR builds it from the realm's own message file, spells, items,
rooms and monsters, so it follows a customised realm. A row that belongs to one spell carries the spell's
number, so a client can join it to `Spells` and, through the item columns, to the items that cast it.

It is **not** a dump of the message file. Lines the engine prints in a fixed, standard shape ("You cast X
on Y for N damage!", "<monster> <verb> you for N damage!" and so on) are left out because a client parses
those by shape (section 4.8). Lines no player sees (a monster-only spell's caster line, room-perspective
lines) are left out too.

NMR 1.8.4 also adds six columns to `Items` (section 5): weapon verbs, the spells an item casts when used
and on hit, and a command code for using it.

---

## 2. Finding and reading the table

### 2.1 Detecting it

- `Info.[NMR Version]` is `v1.8.4` or newer.
- Check that the table and its `Spell` column exist before you query them. Exports made by NMR 1.8.3 and
  older do not have the table, and exports made by 1.8.4 test builds before 2026-11-01 lack `Spell`. NMR
  refuses to update either in place ("Create a new export instead"), so a 1.8.4 version string implies
  both are there, but probing costs nothing.
- `Items` gains `HitVerbs`, `HitVerbsOther`, `MissVerbs`, `UseCode`, `UseSpells` and `HitSpells`,
  appended after `References`. No other column of any existing table changed.

### 2.2 Columns

| Column | Jet type | Size | Meaning |
|---|---|---|---|
| `ID` | Long | | Row number 1..n in NMR's write order. Primary key `pkMessages`. **Not a stable key**: every export renumbers it. |
| `Name` | Text | 100 | A spell name, an item name (item damage-shield rows), a monster name (death rows), or a label made from the line's own text (3.8). Never empty. |
| `Kind` | Text | 12 | `Cast`, `Self` or `Other` (4.1). |
| `Msg` | Text | 100 | The line, literal except for tokens (3.3). |
| `EndMsg` | Text | 100 | Wear-off line. Only on `Self` rows. When present, the effect lasts until this line. |
| `Effects` | Text | 255 | Comma-separated list (3.7, 4.6). |
| `Spell` | Long | | `Spells.Number` of the spell the row belongs to: set on every `Cast` row and on the named `Self` rows of a player or item spell (3.4); 0 on every other row. |

### 2.3 Reading rules

- All columns except `ID` and `Spell` are Text, not Memo.
- `Name` is a Jet reserved word: write `[Name]` in SQL, and bracket `[Msg]` too.
- As in every MME table, an empty value is written as a single `Chr(0)`, not as an empty string or Null.
  Treat a value of exactly one NUL as empty. `EndMsg` and `Effects` are usually empty. `Msg` can be
  empty only on a `Self` row whose status message has a wear-off line but no status line.
- `Kind` is always one of the three values. The `Effects` vocabulary is fixed (3.7). Ignore effects you
  do not know, in case a later NMR adds some.
- Jet returns Unicode. The stock data has no character outside ASCII. Colour codes and other control
  characters are already removed.

---

## 3. How NMR builds the table

### 3.1 Inputs and pass order

The Messages pass runs after NMR's cross-referencing, because it uses the `Learnable` flags (3.4). It
reads the dat files in this order, collects rows in memory, merges duplicates and writes the table once:

1. **Items**: which spells items teach (ability 42), which they cast when used or on hit (abilities 43
   and 114, section 5.2), and which items have a damage-shield message (ability 137). No rows yet.
2. **Item spell chains**: spells reached through EndCast (spell ability 151) from a spell an item casts
   count as item-cast too (3.4).
3. **Rooms** (every room except "Buffer room"s): trap rows, the spells rooms cast (`Spell` field), and the
   spells cast exits and spell traps cast on the player.
4. **Spells**: status, fail, cast, start, self-cast, incoming and shield rows.
5. **Item shields**: one row per name of an item that carries a damage-shield message.
6. **Monsters**: melee hit rows and death text.
7. **Static rows**: 78 chase lines and 29 monster-entrance lines, the same in every export.
8. **Finalize**: drop the "you" duplicates, give nameless rows a name, write the rows with `ID` 1..n.

A row keeps the position where it first appeared, so `ID` order follows the passes. That is a side
effect, not something to rely on.

### 3.2 Which message lines are used

A MajorMUD message record has three lines. Which line a player sees depends on the field that points at
the record:

| Message (field) | Line | Printed to | Becomes |
|---|---|---|---|
| Spell cast message (`CastMsgB`) | 1 | the caster | `Cast` row (player and item spells), or a nameless `Self` row (spells a textblock, cast exit or spell trap casts with the player as caster) |
| | 2 | the target | `Self` row (the "incoming" line) |
| | 3 | the room | never |
| Spell status message (ability 115, DescMsg) | 3 | the target when the spell lands, and by `stat` | `Self` row, `Msg` |
| | 1 | the target when the spell wears off | the same row's `EndMsg` |
| | 2 | look text | never |
| Spell start message (ability 120, StartMsg) | 1 | the caster, once the spell has landed | `Cast` row, only as a fallback (3.5) |
| Spell confuse message (ability 101, ConfuseMsg) | 1 | the affected player when an action fails | `Self` row, `Fail` |
| Spell damage-shield message (ability 137) | 1 | the shield's owner | `Self` row |
| | 2 | the creature that hit the shield | `Self` row |
| Item damage-shield message (item ability 137) | 1 | the wearer | `Self` row, one per item |
| Monster melee hit message (`AttackHitMsg`, attack type 1) | 1 | the victim | `Self` row |
| | 2 | the room | `Other` row |
| Monster death message (`DeathMsg`; if it is unset and attack slot 0 is a melee attack, that attack's hit message) | 3 | the room | `Other` row, `Death` |
| Trap exit messages (exit types 9 and 24, Para3 and Para4: the trigger line and the failed-disarm line; the stock and GreaterMUD engines swap them on type 24, so both are exported) | 1 | the player walking through or disarming | `Self` row, `Trap` |

"none" lines and blank lines are skipped. A message number missing from the message file is skipped and
logged (3.10).

### 3.3 From raw line to `Msg`

1. **Clean**: ANSI colour sequences and every other control character are removed; trailing spaces are
   trimmed.
2. **Substitute placeholders**. The engine fills `%s` and `%d` at print time. NMR replaces them with
   tokens, or with the literal text the engine would print:
   - `%d` always becomes `{dmg}`.
   - Each `%s` takes the next entry from the line's argument list:

     | Line | Arguments, normal form | Short form |
     |---|---|---|
     | Cast line 1 | spell name, target, `{dmg}` | target, `{dmg}` |
     | Start line 1 | target, target | same |
     | Cast line 2 | `{source}`, spell name, `{dmg}` | `{dmg}` (stock engine only; see below) |
     | Status line 3 | none | |
     | Status line 1 (wear-off) | spell name | |
     | Confuse line 1 | none | |
     | Shield line 1 | `{attacker}`, `{dmg}` | |
     | Shield line 2 | `{dmg}` | |
     | Monster hit line 1 | `{source}`, `{dmg}` | |
     | Monster hit line 2 | `{source}`, `{target}`, `{dmg}` | |
     | Death line 3 | `{source}` | |
     | Trap line 1 | none | |

   - "spell name" is written **literally**: "You fire a magic missile at {target} for {dmg} damage!".
   - "target" is `{target}`, except for area target types (3 Divided Area Not You, 9 Divided Attack
     Area, 10 Divided Party Area, 11 Full Area, 12 Full Attack Area, 13 Full Party Area). The engine
     prints nothing there, so NMR substitutes nothing: "You cast greater healing rain on your party,
     healing {dmg} damage!".
   - A `%s` with no argument left becomes `{?}` and is logged. It means the realm's data is
     inconsistent. There are none on stock.
   - The short form is chosen by the spell's message-style byte. Stock exports test bit 0. Exports made
     with "For GreaterMUD Engine" checked (`Info.Legit` = 2) follow that engine: only the exact value 1
     selects the short form, and cast line 2 always gets the normal arguments.
3. **Trim** again, so an empty last substitution leaves no trailing space.

The result is the line exactly as the game prints it, with tokens where the text varies. The monster hit
line 2 receives its damage through a `%s`, so its `{dmg}` is a number printed as text. Match digits there
like anywhere else.

### 3.4 Player spells, item spells and friendly targets

- **Player-castable**: the export's cross-referencing marked the spell `Learnable` (books, trainers,
  textblocks, NPC greets, items), or an item teaches it (ability 42). When no spell in the export is
  marked `Learnable` (an "Export Only" run that re-exported spells without cross-referencing), NMR falls
  back to "has a magery type".
- **Item-cast**: some item casts it when used or on hit (ability 43, CastsSp, with or without a %Spell
  114 before it; section 5.2). A spell chained through EndCast (spell ability 151, followed transitively,
  at most 10 links) from a spell an item casts counts as item-cast too: the item's user is its caster as
  well.
- **One row per spell**: a player or item spell's `Cast` rows carry the spell's own name and number
  (`Spell`). An item that casts the spell gets no rows of its own; its `UseSpells` / `HitSpells` list the
  spell (5.2), and a client finds the item's lines through them (5.4).
- **Friendly target types** are 0 Single, 1 Self, 2 Self or Single, 6 Any and 13 Full Party Area: the
  player can be the one affected by their own spell or item. A player or item spell with one of these
  types also writes its `Self` rows (status, incoming, shield) under its name and number. Every other
  spell's `Self` rows are nameless, with `Spell` 0 (realm-wide). A hostile spell's status text only ever
  reaches its victim, and the victim did not cast it. A spell with a blank name gets neither on its `Self`
  rows.
- **Self-cast spells**: spells the player ends up casting without choosing them, from a textblock `cast`,
  a cast exit (exit type 22, Para1 and Para2) or a spell trap (exit type 24, Para1). The player is the
  caster, so the player sees cast line 1.

### 3.5 Row rules, source by source

Abbreviations: "the spell" = the spell's name, with `Spell` = its number; "self-name" = the same for a
player or item spell with a friendly target type, otherwise nameless with `Spell` 0; "standard" = matches
a filter in 3.6.

**Spells** (every spell record):

| Row | Kind | Name | Msg / EndMsg | Effects | Written when |
|---|---|---|---|---|---|
| Status | Self | self-name | status line 3 / status line 1 | the spell's (3.7), plus `Fail` when the confuse line equals the status line | the spell has a status message and line 3 or line 1 is not blank |
| Fail | Self | nameless | confuse line 1 | `Fail` | the spell has a confuse message, and its line 1 is not the same text as the status line (then it went into the status row) |
| Cast | Cast | the spell | cast line 1 | the spell's | the spell is player-castable or item-cast. **Every form**, standard or not: a client can always find the line of a spell it cast or an item it used (5.4). |
| Start | Cast | the spell | start line 1 | as Cast | as Cast, and the cast line is empty or standard, and the start line does not begin "You cast / invoke / sing " |
| Self-cast | Self | nameless | cast line 1 | the spell's without `RoomSpell`, plus `Trap` for spell traps | the spell is self-cast (3.4) but neither player-castable nor item-cast; its target type is not 4 Monster or 7 Item; it has some effect besides `RoomSpell`; the line is not standard |
| Incoming | Self | self-name | cast line 2 | the spell's | the line contains `{dmg}` and is not standard; **or** the spell has no status message but has a status ability (19, 71, 107, 74, 75, 60, 76, 170); **or** the line has no `{dmg}` and the status message lacks its status line or its wear-off line |
| Shield | Self | self-name | shield line 1, and shield line 2 | `Damage,DamageShield` | the spell has a damage-shield message; each line on its own, unless it is standard |

The incoming rule in plain words: when the status message already shows the effect starting and ending,
the incoming line adds nothing and is left out. When it does not, the incoming line is the client's only
sign of the effect.

**Items**: an item with a damage-shield message (item ability 137) gets a `Self` row for line 1, one per
item name, `Damage,DamageShield`, `Spell` 0 (the shield is an item ability, not a spell), unless the line
is standard. Line 2 reaches only the wearer's attacker, never the wearer, so it is not exported. These are
the only rows named after an item.

**Rooms**: each trap exit's message line 1 is a nameless `Self` row with `Damage,Trap`. A spell in a
room's `Spell` field gets `RoomSpell` in its effects. Cast exits and spell traps feed the self-cast rows.

**Monsters**:

| Row | Kind | Name | Msg | Effects | Written when |
|---|---|---|---|---|---|
| Hit on me | Self | nameless | hit line 1 | `Damage,Monster` | a melee attack (type 1, slots 0-4) whose line 1 is not standard |
| Hit on another | Other | nameless | hit line 2 | `Damage,Monster` | the same attack's line 1 was exported, and line 2 does not contain `{target} for {dmg} damage` |
| Death | Other | the monster | death line 3 | `Death` | the monster regenerates (`RegenTime` > 0) or is game-limited (`GameLimit` > 0), **and** gives at least 101 experience (`Experience` times `ExpMulti` on v1.11j and newer data) |

Spell attacks (type 2) are covered by the spell rows.

**Static rows**: 78 chase lines ("{target} slips into the dark alley.", `Other`, `Chase`) and 29 monster
entrances ("A moss zombie rises from the ground!", `Other`, `CheckRoom`), nameless, in the stock game's
wording, in every export. A chase line is what you see when someone leaves through a command exit. NMR
does not export the command that follows them.

### 3.6 Standard forms that are filtered out

These patterns run on the substituted text. A line that matches is left out where 3.5 says "standard".
The cast row of a player or item spell is never left out; there a match only decides whether the start
line is added.

| # | Pattern (regular expression on the substituted text) | What it is |
|---|---|---|
| P1 | `^You cast .+ (on\|at\|upon) .*\{dmg\} (damage\|healing\|hit points)!?$` | my standard cast line with an amount |
| P2 | `^\{source\} casts .+ (on\|at\|upon) you.*\{dmg\} damage!?$` | standard incoming cast line |
| P3 | `^You are .* for \{dmg\} damage!?$`, `^You take \{dmg\} damage`, or `you for \{dmg\} damage` anywhere | standard incoming damage |
| P4 | `^\{target\} .*\{dmg\}.* damage` or `^\{attacker\} .*\{dmg\}.* damage` | third-party damage or heal observation |
| P5 | `^(The \|An? )?\{source\} [a-z-]+ you for \{dmg\} damage!$` | standard monster hit on me (one verb word) |
| P6 | `\{target\} for \{dmg\} damage` anywhere | monster hit on someone else in the standard shape, any verb phrase |
| P7 | `^You (cast\|invoke\|sing) ` | my standard cast line without an amount |

Where they apply: cast line 1 uses P1, P3, P4 (with `{dmg}`) or P7 (without) to decide on the start line;
start line P7; self-cast line P1, P3, P4, P7; incoming line P2, P3, P4; shield lines P3, P4; monster hit
line 1 P3, P5; monster hit line 2 P6. Status, fail, trap and death lines are never filtered.

### 3.7 Effects

A spell's effects come from its ability slots. Duration only selects the `...OverTime` variants:

| Effect | Spell ability (name in NMR's ability list) |
|---|---|
| `Poisoned` | 19 Poison |
| `Confused` | 71 Confusion |
| `Blinded` | 107 BlindUser |
| `Held` | 74 HoldPerson, 75 Paralyze |
| `Fear` | 60 Fear |
| `Mute` | 76 Mute |
| `Sleep` | 170 Sleep |
| `Damage` / `DamageOverTime` | 1 Damage, 17 Damage(-MR); `OverTime` when the spell's duration > 0 |
| `Drain` | 8 Drain, 175 StealHPtoMP |
| `Heal` / `HealOverTime` | 18 Heal; `OverTime` when the duration > 0 |
| `HPRegen` | 123 HPRegen |
| `ManaRegen` | 145 ManaRgn |
| `ManaHeal` | 150 HealMana |
| `ManaSteal` | 174 StealMana, 176 StealMPtoHP |
| `DamageShield` | 72 DamageShield |
| `Cure` | 20 Cure Poison, 73 DispellMagic, 81 Freedom, 84 Rcrs (remove curse) |
| `RoomSpell` | the spell is in some room's `Spell` field |

Fixed effects on other rows: `Fail` (confuse rows), `Damage,Trap` (traps), `Damage,DamageShield`
(shield rows), `Damage,Monster` (monster hits), `Death`, `Chase`, `CheckRoom`. Self-cast rows from spell
traps add `Trap`.

A spell with none of these abilities (bless, armour spells, stat buffs) has empty `Effects`. Its `EndMsg`
still marks it as timed, and the `Spells` table's ability columns give the details. When rows merge
(3.8), their effects are combined, so a row can carry the union of several sources.

### 3.8 Merging duplicates, and names

- **Merge key**: rows with a spell (`Spell` > 0) merge when Kind, Spell, Msg and EndMsg are all equal.
  Item damage-shield rows merge when Kind, Name, Msg and EndMsg are equal. All other rows (nameless `Self`
  rows, every `Other` row) merge when Kind, Msg and EndMsg are equal. The comparison is exact
  (case-sensitive) and happens after substitution. A merged row keeps its first position and gets the
  union of the effects, in order of first appearance.
- So each spell gets **its own rows under its own number**, even when several spells print the same
  text, and an item adds no rows for its spells. "You feel lucky!" / "The effects of bless wear off!" is
  exported once as `bless` and once as `chant`; the items that cast bless list it in `UseSpells` instead.
  Two spells with the same name and the same line stay two rows, one per number.
- A nameless line with the same text as a named row stays a separate row. Example: a monster-only spell
  that shares a player spell's status text.
- **Death text** printed by differently named monsters (compared case-insensitively) loses its monster
  name and becomes a realm-wide row.
- **"You" duplicates**: after all passes, a nameless `Self` row without an `EndMsg` is dropped when its
  `Msg` equals a `Cast` or `Other` row's `Msg` with `{target}'s` read as "your" and `{target}` read as
  "you". Its effects go into that row. A client is expected to accept "you" / "your" for `{target}`
  (4.3).
- **Names for nameless rows**: the row is named after its `Msg` (or its `EndMsg` when `Msg` is empty),
  shortened. NMR drops a leading `The {source}'s `, `The {source} `, `{source}'s ` or `{source} `, then one
  trailing ` for {dmg} damage!`, ` {dmg} damage!` or ` {dmg} hit points!`, then every `{target}'s ` and
  `{target} `, and trims what is left. If nothing is left, the whole line is used. "The {source} claws you
  with its pincers for {dmg} damage!" is named "claws you with its pincers". These names are labels for
  display. They change when the text changes, so they are not keys. A `Cast` row of a spell with a blank
  name is named the same way; its `Spell` still says which spell it is.

### 3.9 What is never exported

- Cast line 3 (room), status line 2 (look text), confuse lines 2-3, shield line 3.
- Cast line 1 of a spell no player or item casts, unless a textblock, cast exit or spell trap casts it
  on the player (3.5).
- Lines in a standard form (3.6), except the cast lines of player and item spells.
- An incoming line without an amount, when the spell's status message has both a status line and a
  wear-off line.
- Monster misses, dodges and movement; item destruction messages; exit command words; room descriptions
  and room sounds; textblock message text.
- Death text of monsters below the experience gate, or that neither regenerate nor are game-limited.
- Engine strings that are not in the data: "You have N lives left.", "You feel your life-force ebbing."
  and "You look around stupidly and do nothing!".
- Message numbers. No column links a row back to the message file.

### 3.10 Warnings

The export writes `NMR-Log_MMEMessages.txt` next to the MDB only when something needed attention: a `{?}`
placeholder, a message number missing from the message file, a spell with a blank name, a value cut to its
column size, or a dat file that could not be read. The completion message gives the count.

---

## 4. Using it in a client

### 4.1 Kinds, from the player's seat

- **Cast**: confirmation that **my** spell or item worked. `Name` and `Spell` are the spell, also when an
  item cast it (5.4 finds an item's lines). The line may carry the amount ("You cast agony on {target}
  for {dmg} damage!"), or none ("A cool wind blows throughout the room!", "You rub your jade amulet!"). A
  spell whose cast line is empty or standard may also have a start line ("You lay hands on {target}.").
  Both are printed only once the spell has landed. **The effects on a Cast row describe what the spell
  does to its target, not to me.**
- **Self**: something happened **to me**.
  - With an `EndMsg` it is a tracked effect: `Msg` when it starts (and in `stat`), `EndMsg` when it
    wears off.
  - Without one it is a one-off line: incoming damage in a non-standard wording, a failed action
    (`Fail`), a non-standard monster hit (`Monster`), a trap (`Trap`), my damage shield or my item's
    spikes hitting back (`DamageShield`), or a spell a textblock, cast exit or spell trap cast on me.
  - A Self row carries a spell's name and number only when the player can be affected by their own
    spell or item (3.4). An item damage-shield row carries the item's name.
- **Other**: something about someone else: a monster's non-standard hit on another player, a monster's
  death text (`Name` = the monster, unless several monsters share the text), a chase line, or a monster
  appearing or changing form.

### 4.2 Matching a game line

`Msg` is the whole line from its first character, as the game prints it. A workable matcher:

1. Strip ANSI sequences from the incoming line and trim trailing spaces, as NMR did with the data.
2. Turn each `Msg` (and `EndMsg`) into a pattern: escape the literal text and replace the tokens
   (`{target}'s` before `{target}`):

   | Token | Pattern | Capture |
   |---|---|---|
   | `{dmg}` | `(\d+)` | the amount |
   | `{source}` | `(.+?)` | caster or attacking monster, including any "The" |
   | `{target}'s` | `(?:(.+?)'s\|your)` | target, or me |
   | `{target}` | `(you\|.+?)` | target, or me |
   | `{attacker}` | `(.+?)` | the creature that hit the damage shield |
   | `{?}` | `.*?` | nothing |

3. Anchor at the start. Anchoring at the end is safe too, except on a row the export cut at 100
   characters (logged). A prefix match works for both.
4. Compare case-sensitively. The rows come from the same data the game prints.

Many rows contain no token at all. A hash lookup on the exact text handles those before you try the
patterns.

### 4.3 Tokens

- `{source}` is printed by the game with its own article ("The giant rat", "Ragnar"), so do not add
  one.
- `{target}` may be "you" (and `{target}'s` may be "your"). When it is, the line is about me: apply the
  row's effects to myself. NMR drops a `Self` row that would only duplicate such a line (3.8).
- Area, room and party spells have no `{target}`. The text already reads as the game prints it.
- `{attacker}` appears only on damage-shield rows (10 on stock).
- The spell name is never a token. It is part of the literal text.
- `EndMsg` contains no tokens on stock. The engine passes the spell name there, which NMR writes
  literally.

### 4.4 Names

- A row belongs to a specific spell when its `Spell` is not 0. Join `Spell` to `Spells.Number` for the
  spell's abilities, target type or `Learnable` flag. Its `Name` is then the spell's name, in the game's
  spelling and case (up to 29 characters), but **join on `Spell`, not on the name**: several spells can
  share a name (10 names among the spells with rows on stock v1.11p, 29 on Paradigm 1.9.1d). On stock,
  every Cast row and the named Self rows have a `Spell` (counts in section 7).
- An item damage-shield row (`Spell` 0, `DamageShield`) is named after the item; it matches `Items.Name`.
- Every other row is realm-wide. Its `Name` is a monster name (death rows) or a label made from its text
  (3.8). Display it if you like, but do not key on it.
- `Name` is not unique. One spell can have a Cast row, a start row, a status row, an incoming row and
  shield rows, all with its number.
- To credit a line to "the spell I just cast", look for a Cast row matching the line whose `Spell` is
  that spell. For "the item I just used", look among the rows of the item's spells (5.4).

### 4.5 Timed effects

A `Self` row with an `EndMsg`:

- When `Msg` appears, the effect is on. The game also prints `Msg` in `stat`, so `stat` output can
  re-sync your state.
- When `EndMsg` appears, it is off. One status line can end in different ways on different rows, so
  match the `EndMsg` too, not just the `Msg`.
- The `Effects` say what the effect does (4.6). Empty `Effects` means a buff or debuff with no effect a
  client acts on (bless, armour); it still starts and ends.
- `Messages` has no duration. The `Spells` table's `Dur` / `DurInc` / `DurIncLVLs` give the base and its
  level scaling, but the caster's stats change it further, so rely on the wear-off line.

### 4.6 Reading `Effects`

What an effect means depends on the row's Kind: **Self** = it applies to me; **Cast** = it applies to the
target of my spell; **Other** = it applies to someone else.

| Effect | Meaning | Usually on |
|---|---|---|
| `Poisoned` | poisoned | Self with EndMsg; Cast |
| `Confused` | confused **or stunned**: the game uses one ability for both, so the text is the only clue | Self with EndMsg; Cast |
| `Blinded` | blind | Self with EndMsg; Cast |
| `Held` | cannot move | Self with EndMsg; Cast |
| `Sleep` | asleep: cannot move or attack | Self with EndMsg; Cast |
| `Fear` | afraid: actions may fail, may flee | Self with EndMsg; Cast |
| `Mute` | cannot cast | Self with EndMsg; Cast |
| `Damage` | one-off damage | Self without EndMsg; Cast; Other (monster hits) |
| `DamageOverTime` | losing HP over time; the ticks print nothing | Self with EndMsg; Cast |
| `Drain` | HP drained | any |
| `Heal` / `HealOverTime` | healed now / over time | Cast; Self |
| `HPRegen` / `ManaRegen` | regenerating HP / mana faster | Self with EndMsg |
| `ManaHeal` / `ManaSteal` | mana restored / stolen | any |
| `DamageShield` | a damage shield; on a one-off Self row, my shield or spikes hit `{attacker}` | Self |
| `Cure` | cures poison, dispels, frees, removes a curse | Cast; Self |
| `RoomSpell` | the spell is also cast by a room | any spell row |
| `Fail` | my last action failed | Self without EndMsg, or merged into a status row (stun) |
| `Monster` | a monster's melee hit | Self, Other |
| `Trap` | a trap fired (or a disarm failed) | Self |
| `Death` | a monster died; `Name` says which, unless shared | Other |
| `Chase` | `{target}` left through a command exit; follow if they are in your party | Other |
| `CheckRoom` | a monster appeared or changed form; look at the room again | Other |

Derived conditions: disease is `DamageOverTime` plus stat drains (see the spell's abilities); "cannot
attack" is `Sleep` or `Held`.

### 4.7 When several rows match one line

That is normal:

- The same text under several spells (`bless`, `chant`). For a Cast row, prefer the spell you just cast,
  or a spell of the item you just used (5.4). For a Self row, the effect is the same whoever cast it, so
  apply it once.
- A named row and a nameless row with the same text (a player spell and a monster-only spell sharing a
  status text). Both are right. Apply the effect once.
- A short line can be a prefix of several longer ones if you match by prefix. Prefer the longest match,
  or anchor at the end.

### 4.8 Lines you still parse yourself

The table carries only what deviates from the engine's fixed shapes. A client still needs its own rules
for incoming and outgoing standard damage, misses and dodges, engagement ("moves to attack"), deaths,
kills and experience, standard cast success and failure, and the wear-off keywords. The full list is the
"Standard formats clients must implement themselves" section of `docs/mme_export_notes.txt`, and the
patterns in 3.6 show exactly which shapes NMR assumes you handle.

### 4.9 Examples (stock v1.11p)

The rows are from a stock export; the game lines are made up to show a match.

| Game line | Row | What to do |
|---|---|---|
| You feel resistant to cold! | Self `resist cold`, EndMsg "The effects of resist cold wear off!" | effect on until the EndMsg line |
| You cast agony on goblin for 23 damage! | Cast `agony`, "You cast agony on {target} for {dmg} damage!" | my agony landed on "goblin" for 23 |
| A cool wind blows throughout the room! | Cast `black wind` | my black wind landed |
| You rub your jade amulet! | Cast `jade amulet`, `Spell` 156; the item jade amulet has `UseSpells` `156` | using my jade amulet worked |
| You feel a little better. | Cast `food`, `Spell` 316; the food items list 316 in `UseSpells` | the food I just ate worked |
| You lay hands on Ragnar. | Cast `annointed hands`, "You lay hands on {target}." (start line) | my annointed hands landed on Ragnar |
| The shield spike stabs giant rat for 4 damage! | Self `spiked shield`, "The shield spike stabs {attacker} for {dmg} damage!" | my item's spikes hit the rat for 4 |
| Ragnar slips into the dark alley. | Other, `Chase`, "{target} slips into the dark alley." | Ragnar left through a command exit |
| A moss zombie rises from the ground! | Other, `CheckRoom` | look at the room again |

---

## 5. Item columns (`Items`)

### 5.1 Weapon verbs

| Column | Text | Source | Example (`longsword`) |
|---|---|---|---|
| `HitVerbs` | verbs for "You `<verb>` X for N damage!" | item hit message line 1 | `slash\|impale\|hack` |
| `HitVerbsOther` | verbs for "`<player>` `<verb>` X for N damage!" | item hit message line 2 | `slashes\|impales\|hacks` |
| `MissVerbs` | the player's miss verbs | item miss message line 1 | `swing at` |

- Text 74 each, `"|"`-separated alternatives, control characters stripped.
- Filled only for weapons (`ItemType` = 1), every weapon class including 1H Blunt (`WeaponType` 0).
  Every other item has `Chr(0)`.
- Use: tell which weapon (or which hand, after a swap) landed a blow, or recognise another player's
  weapon from their hit verbs. Several weapons can share the same verbs.

### 5.2 Spells cast when used and on hit

| Column | Size | Content | Example (`serpent staff`, stock v1.11p) |
|---|---|---|---|
| `UseSpells` | 120 | spell numbers the item casts when used, `"\|"`-separated | `81` |
| `HitSpells` | 120 | spells it casts on hit, as `<spell number>:<chance %>` pairs, `"\|"`-separated | `409:25` |

Both are decoded from the item's ability slots, which the export already has in `Abil-n` / `AbilVal-n`.
Walking the slots in order, a `CastsSp` (43) is cast when the item is used, unless a `%Spell` (114) comes
before it. Then it is the spell cast on hit, and the `%Spell` value is its chance per hit ("the spell cast
is specified by the NEXT CastsSp", in NMR's ability list). A `%Spell` with no later `CastsSp` is ignored.
Each list is in slot order without repeats; `Chr(0)` when empty. Every item type is decoded, not only
weapons. Look the spell numbers up in `Spells` for target type, damage range, duration and so on, and in
`Messages.Spell` for the lines they print (5.4).

### 5.3 `UseCode`

A command code for using the item, Text 5: `#` plus up to 4 lowercase letters. Only items with at least
one use spell have one; every other item, including one that only casts on hit, has `Chr(0)`.

- **Made from the name.** Words are lowercased, a possessive "'s" is dropped, only letters count, and
  "a", "an", "the" and "of" are skipped. One word gives its first 4 letters (incense → `#ince`). Several
  words give their initials, with the last initial replaced by the whole last word when there are fewer
  than 4, cut to 4 (jade amulet → `#jamu`, tankard of mead → `#tmea`, embroidered black gauntlets →
  `#ebga`).
- **Unique per name.** Items with exactly the same name (compared case-insensitively) share one code: in
  the game, "use void sphere" cannot tell them apart either. When a differently named item already has the
  first choice, NMR tries the first 2, then 3 letters of the first word plus the other initials and the
  last word (black potion → `#blpo` when blue potion has `#bpot`), then any 4 letters of the name in order,
  starting with its first letter. Digits are the last resort (`#gor2`). Items are processed in item-number
  order, so the lower-numbered item keeps the natural code. None needs a digit on stock v1.11p (209
  usable items, 206 codes) or on Paradigm 1.9.1d (310 usable item names).
- **Stable.** The same names give the same codes. When NMR updates an existing export, every code already
  in it stays with its item name, and new codes avoid them, so adding an item never changes another item's
  code. A code changes when its item is renamed.
- **Apart from spell codes.** The `#` keeps item codes apart from the game's spell short codes
  (`Spells.Short`); no game spell code in the stock or sample data starts with `#`. NMR does not check.
- It is a hint for the client's command list, not a key. Key items by `Items.Number`.

### 5.4 Modelling an item in a client

- **Using it**: an item with a `UseCode` is something the player can invoke, the way a spell is cast. A
  client that keeps spells and item uses in one list can add one entry per such item: name = `Items.Name`,
  code = `UseCode`, key = `Items.Number`, stats from the `Spells` rows in `UseSpells`. The `Spells` rows
  stay as they are: a class still casts a spell by its own name and `Short`, even when an item also casts
  it.
- **Its spells, chains included**: the spells in `UseSpells`, plus every spell reached from them through
  EndCast: a `Spells` row whose `Abil-n` is 151 names the next spell in the matching `AbilVal-n`; follow
  it transitively (NMR follows at most 10 links, each spell once). A chained spell is cast by the item's
  user too.
- **Its lines**: the `Messages` rows whose `Spell` is one of those spells. The Cast rows confirm using
  it; the Self rows (a spell with a friendly target type, 3.4) are its effects on me, with an `EndMsg`
  for status effects. A spell that prints no cast line has no Cast row.
- **On-hit spells**: `HitSpells` says which spells and how often; the rows of those spells (and their
  chains) are the lines they print to the wielder. They need no command.
- **Damage shield**: the Self rows with `Name` = the item and `DamageShield` (`Spell` 0).
- The same spell's rows serve every item that casts it and the class that casts it, so a line alone does
  not say which item fired. The client knows what the player just used or wields.
- Example (stock v1.11p): the serpent staff has `UseSpells` `81` and `HitSpells` `409:25`. The rows with
  `Spell` 81 are the lines of using it; the rows with `Spell` 409 (constriction) are what its on-hit spell
  prints, about a quarter of the time it hits.

---

## 6. Gotchas

- **`Chr(0)` is empty.** A naive reader sees a one-character string.
- **`ID` changes every export.** Do not store it.
- **Do not take effects from Cast or Other rows as your own.** A Cast row's `Blinded` means your target
  is blind. The exception is a `{target}` that matched "you" / "your" (4.3).
- **`Confused` can be a stun.** Check the text, or treat it as "may not act".
- **Not every line is here.** Standard forms are left out on purpose (3.6, 4.8). If a line seems to be
  missing, check those patterns first.
- **The same text can repeat under different spells** by design (3.8). Deduplicate by text yourself if
  you only need text.
- **Join on `Spell`, not on `Name`.** Several spells can share a name, and an item's lines are under its
  spells' names, not the item's (5.4).
- **Realm-wide names are labels.** They are made from the text and change with it.
- **Key item uses by `Items.Number`, not by `UseCode`.** A code follows the item's name, and items with the
  same name share it.
- **An item can both be used and cast on hit**, even the same spell. The spell's rows are the same
  either way; `UseSpells` / `HitSpells` say which.
- **Byte limits are your client's.** Columns are 100 characters (`Effects` 255). If your client stores
  bytes in a code page, convert first, then check lengths.
- **GreaterMUD exports** (`Info.Legit` = 2) follow that engine's placeholder order (3.3). The exported
  text is already final, so this changes nothing for a reader.

---

## 7. Stock numbers and checks

A stock v1.11p export (NMR 1.8.4, 2026-11-01) has:

| Kind | Rows |
|---|---|
| Cast | 415 (all with a `Spell`, 402 spells; 4 from spells reached only through EndCast chains of item spells) |
| Self | 477 (274 with an EndMsg; 144 with a `Spell`; 8 item damage-shield rows named after the item) |
| Other | 419 (213 death texts, 78 chase lines, 29 monster entrances) |
| Total | 1,311 |

Writing item lines per item instead gave 1,417 rows for the same data, with the same 1,210 distinct
lines. 12 rows repeat a line under a second spell of the same name (`resist cold`, `barkskin`,
`drain life` and others), one per spell number.

In `Items`, 209 items have a use spell (206 codes, as three names repeat) and 127 have an on-hit spell.
Spells reached only through EndCast chains of item spells have rows too: the "slam" line of spear slam
knockdown (chained from spear, which Dusk and the nexus spear cast when used) and the lines of elemental
lightning, fire and ice (the staff of the elements' on-hit elemental earthquake chains to them).

NMR's own check of an export is `docs/verify-mme-messages.ps1` (32-bit PowerShell, Jet). It checks the
schema, that no raw `%s` / `%d` is left, the Kinds, that rows are unique, the "you" duplicates, names,
the standard-form filters, the item columns against the `Abil-n` columns, the `UseCode` rules, the
`Spell` column (on every Cast row, named after its spell, never after an item), and spot checks against
stock data.
