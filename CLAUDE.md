# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

Nightmare Redux (NMR) is a **Visual Basic 6** desktop editor for the BBS game *MajorMUD*. It reads and writes the game's **Btrieve (Pervasive) v6 `.dat` files** directly, and can import/export game data to/from **Access `.mdb`** databases in several formats. The primary consumer of one export format is a separate tool, **MME (MegaMud Explorer)** — many code paths exist solely to produce MME export tables.

## Building & running (no CLI build/test exists)

This is VB6: there is **no command-line build, linter, or test suite**. Everything happens in the **VB6 IDE (VB98)** on Windows, 32-bit.

- Open `_NMR-ProjectGroup.vbg` — a project group containing two projects:
  - `NMRTaskBar.vbp` → builds `NMRTaskBar.ocx` (an MDI taskbar UserControl). Build this **first**; the main project references the compiled OCX.
  - `Nightmare Redux.vbp` → builds `nightmare_redux.exe`. Startup object is `frmMain`.
- Build via **File → Make `<project>`** in the IDE. Run/debug with F5.
- Requires registered dependencies (see `.vbp` `Reference=`/`Object=` lines): DAO 3.6 (`Dao360.dll`), Scripting Runtime, ADOX (`msadox.dll`), MSCOMCTL, COMDLG32, TABCTL32, `msstdfmt.dll`, the bundled `exlimiter.ocx`, and **`PAln32.dll`** (Pervasive Btrieve Alignment Library — see below).
- The Btrieve engine itself ships in-repo: `WBTRV32.DLL`, `W32MKDE.EXE`, etc. `WBTRV32 - Registry Settings.reg` configures the MicroKernel.
- COM components must be registered (`regsvr32`) and the runtime is 32-bit only.

There are **no unit tests**; verification is manual against real `.dat` files.

## The two-version build (vN vs vO) — critical

NMR ships as **two EXEs** targeting different MajorMUD data formats, and the difference is which fieldmap module is compiled in:

- `modFieldmaps_vO.bas` (`WorksWithN = False`) → **v1.11o and newer** dat files → `nightmare_redux.exe`. **This is the one currently referenced by `Nightmare Redux.vbp`.**
- `modFieldmaps_vN.bas` (`WorksWithN = True`) → **v1.11h through v1.11n** dat files → `nightmare_redux_n.exe`.
- `modFieldmaps.bas` is an older/legacy combined copy, **not in the `.vbp`** — treat as reference unless you know otherwise.

To build the "n" variant you swap the module in the project so the appropriate `WorksWithN` is active. The two modules define **the same public symbols** (record types, fieldmap arrays, `Add*FieldMap` subs) with different byte layouts. Consequences for editing:

- **Any change to a record layout must usually be made in both `_vN` and `_vO`** (and possibly the legacy `modFieldmaps.bas`). Version-specific lines are flagged with the comment marker `'***I_THROUGH_N***` — grep for it. See e.g. the `RoomFldMap` upper-bound that differs by one between versions.
- `WorksWithN` also drives settings: INI keys get an `_n` suffix in the n-build (e.g. `eDatFileVersion_n`, `DatCallLetters_n`, `WGPath_n`) so both EXEs can share one `settings.ini`.

## Btrieve data access layer

Records are **fixed-length C-aligned binary** read via raw Btrieve calls and unpacked into VB UDTs using the PALN32 alignment library.

- `modBtrieve.bas` declares `BTRCALL` (into `wbtrv32`) plus all operation constants (`BGETFIRST`, `BGETNEXT`, `BINSERT`, `bUpdate`, `BDELETE`, …) and `BtrieveErrorCode()` for human-readable status messages.
- For **each game table** (Race, Class, Spell, Monster, Item, Shop, Room, Message, Textblock, User, Action, Bank, Gang, DBStat) the fieldmap module declares a parallel set:
  - `XxxRecType` — the typed UDT. **Field byte offsets and lengths are documented inline** in the `Type` definitions (the `A/B/C` columns = field len / cumulative len / field index). This is the canonical record-format documentation.
  - `Xxxdatabuf` — the raw byte buffer, `XxxPosBlock` — Btrieve position block, `XxxKeyBuffer` — key buffer.
  - `XxxFldMap()` of `PALN32.FieldMap`, populated by `AddXxxFieldMap` (all wired up in `IntFieldMaps`).
- Conversion between raw bytes and the UDT goes through PALN32 via the `XxxRowToStruct` / `XxxStructToRow` helpers in `modMMUDFunctions.bas`. **When you add or resize a field, update the UDT, its buffer-size const, the `Add*FieldMap`, and the offset comments together** or the alignment silently corrupts later fields.
- Dat files are named `w<CallLetters><table>.dat` (e.g. `wccknms2.dat` monsters, `wccmp002.dat` rooms/maps, `wccitem2.dat`, `wccuser2.dat`). Call letters default to `cc`; the server path comes from the `WGPath` setting. These `.dat` files are **not in the repo** (gitignored; sample sets are in `_fresh dats.7z`).

## Application structure

- `frmMain.frm` is the MDI parent (menus, status bar showing call letters / version / "WRITING DISABLED"). `bDisableWriting` globally gates all writes — respect it.
- Each game entity has its own large editor form (`frmRoom`, `frmMonster`, `frmItem`, `frmSpell`, `frmUser`, `frmClass`, `frmRace`, `frmShop`, `frmTextblock`, …). These are the bulk of the codebase and are mostly self-contained UI + load/save against the Btrieve layer.
- `modMMUDFunctions.bas` is the shared domain library: cross-reference lookups (`GetMonsterName`, `GetItemName`, `GetRoomExits`, `GetSpellRange`, `IsRoomLair`, etc.), textblock `Encrypt`/`DecryptTextblock`, and `CreateAccessTables` (builds the MME `.mdb` export). Editors and the import/export forms call into here rather than re-reading dat files ad hoc.
- `modMain.bas` holds global state, Win32 API declares, enums, and UI helpers.
- `modSettings.bas` wraps `settings.ini` via `Get/WritePrivateProfileString` (`ReadINI`/`WriteINI`). `ReadINI` lazily writes a default when a key is missing — note the `_n` suffix handling tied to `WorksWithN`.
- `modUpdateFile.bas` compiles the game's "update file" (`BINSERT` of change records into the update dat) — this is how edits are pushed to a live server install.
- Import/export forms: `frmDatabaseImport`, `frmDatabaseExport`, `frmDatabaseMerge`, `frmMME_Export` (DAO/ADOX against `.mdb`; `ability.mdb` is a bundled read-only lookup of game abilities).
- Long operations route progress through `frmProgressBar` and are generally cancelable; `bUseCPU` toggles `DoEvents`-heavy "use all CPU" mode.
- `NMRTask_*` files belong to the `NMRTaskBar.ocx` subproject, not the main app.

## Conventions in this codebase

- Modules use `Option Explicit` + `Option Base 0`; `modBtrieve`/fieldmaps use `DefInt A-Z`.
- Errors funnel through `HandleError` (`On Error GoTo error:` then `Call HandleError`). Follow this pattern in new procedures.
- `.frx`/`.oca`/`.exe`/`.dll`/`.ocx` are binary (see `.gitattributes`); source files are forced CRLF. Do not reformat or normalize line endings.
- Editing a real MajorMUD install with this tool **voids the game license** — this is stated intent, not a bug; the app deliberately exposes raw field editing.
- `README.md` is a long-form changelog (newest version at top); `docs/` has format notes (`mme_export_notes.txt`, `roomspara.txt`, `rec change details.doc`, etc.) useful when reverse-engineering field meanings.
