# Plan: Insert in the Message and Textblock editors pre-fills a taken number (#12)

Branch `261102-insert-next-number` (off `v184`, v1.8.4 unreleased). Status: **4.1-4.3 implemented (no
changelog credit, EXEs not rebuilt); not yet compiled or tested (section 5).**

## 1. Problem

`cmdInsert_Click` in `frmMessage.frm` and `frmTextblock.frm` pre-fills the InputBox with "current number + 1".
That number almost always exists, so OK fails with a duplicate-key error (status 5). The other editors
(e.g. `frmItem.frm:4581`) pre-fill `0`, which Btrieve turns into the next highest number.

## 2. Verified facts

Checked against source and the key specs in the FCR page of every sample set in `_fresh dats/` (1.11i through
1.11p, 1.11L, 1.11M, paramud). All ten sets have the same keys:

| File | Key 0 spec (FCR 0x124) | Meaning |
|---|---|---|
| `wccitem2.dat` | type `0x0F` | autoincrement (reference: Item Insert with `0` works) |
| `wccmsg2.dat` | length 4, type `0x0F` | **autoincrement**, so inserting number 0 already gets the next highest number |
| `wcctext2.dat` | length 18, type `0x0E` | **unsigned binary**, no autoincrement, so `0` would really insert block 0 |

- Message: the prompt already says "Enter 0 for the next highest number."; only the default is wrong.
- Textblock: Btrieve compares the unsigned-binary key as one little-endian number, so the block-number bytes at
  the end are the most significant and key order is block order. `BGETLAST` returns the highest block.
  The key is 18 bytes (`PartNum` 2 + `LeadIn` 14 + low 2 bytes of `Number`), so this holds while block numbers
  stay below 65536 (stock max is 10003).
- Textblock record layout (`TextblockRecType`, identical in `_vO`, `_vN` and legacy `modFieldmaps.bas`):
  `PartNum` @0, `LeadIn(1 To 14)` @2, `Number` @16, `LinkTo` @20, `Data` @24. In the 1-based
  `TextblockDataBufType.buf`, `Number` is `buf(17)`..`buf(20)`.
- Both forms are shared by the vO and vN builds; no fieldmap module changes are needed.

## 3. Correction to the fix proposed in the issue thread

The thread's textblock code calls `BGETLAST` on `TextblockPosBlock` before the prompts and says moving the
position is harmless because the routine re-reads the current block (`BGETEQUAL`) before inserting. That re-read
only runs **after both prompts**. If the user presses Cancel on either prompt, the routine exits with the
position left on the last block, and the `<<` / `>>` buttons (`cmdLineMove_Click`, `BGETPREVIOUS` / `BGETNEXT`
from the current position) then move relative to block 10003 instead of the block on screen.

Fix: keep the `BGETLAST` lookup, but move the existing `BGETEQUAL` re-seek up so it runs right after the lookup
and before the prompts. Its failure handling (copy `TextblockKey.LeadIn`, blank `Data`) stays after the prompts so
a Cancel never modifies `TextblockRec`. Nothing between the re-seek and that check touches `nStatus`.

## 4. Changes

### 4.1 `frmMessage.frm`, `cmdInsert_Click` (line 566)

Change the InputBox default from `Val(txtNumber.Text) + 1` to `"0"`, exactly as `frmItem.frm:4581` does.
`BINSERT` writes the assigned number back into `Messagedatabuf`, so the existing `DispMessageInfo` shows it.

### 4.2 `frmTextblock.frm`, `cmdInsert_Click` (lines 1939-1982)

Target shape (new or moved lines marked):

```vb
Private Sub cmdInsert_Click()
On Error GoTo error:
Dim nStatus As Integer, NewTextBlockNumber As String, x As Integer
Dim NewTextBlockPart As String
Dim tLast As TextblockDataBufType, nNext As Long                                  'NEW

If bDisableWriting = True Then MsgBox "Writing Currently Disabled -- Check out the File menu.", vbInformation: Exit Sub
If bLoaded = True Then SaveTextBlock

'textblocks have no autoincrement key; key order is block order, so the last record is the highest block    'NEW
nNext = Val(txtNumber.Text) + 1                                                  'NEW
If BTRCALL(BGETLAST, TextblockPosBlock, tLast, TextblockMaxBufSize, ByVal TextblockKeyBuffer, KEY_BUF_LEN, 0) = 0 Then  'NEW
    nNext = (CLng(tLast.buf(17)) Or CLng(tLast.buf(18)) * &H100& Or CLng(tLast.buf(19)) * &H10000) + 1          'NEW
End If                                                                           'NEW

're-read the current block so Cancel leaves the position where it was          'NEW
TextblockKey.Number = Val(txtNumber.Text)                                        'MOVED up
TextblockKey.PartNum = Val(txtPart.Text)                                         'MOVED up
nStatus = BTRCALL(BGETEQUAL, TextblockPosBlock, TextblockDataBuf, TextblockMaxBufSize, TextblockKey, KEY_BUF_LEN, 0)  'MOVED up

NewTextBlockNumber = InputBox("New TextBlock Number:", "Insert", nNext)          'CHANGED default
If NewTextBlockNumber = "" Then Exit Sub

NewTextBlockPart = InputBox("New TextBlock Part Number" & vbCrLf & vbCrLf & "(First part is actually part 0)", "Insert", 0)
If NewTextBlockPart = "" Then Exit Sub

If Not nStatus = 0 Then                                                          'unchanged block, now after the prompts
    For x = 1 To 14
        TextblockRec.LeadIn(x) = TextblockKey.LeadIn(x)
    Next
    TextblockRec.Data = String(Len(TextblockRec.Data), Chr(0))
End If
' ... rest unchanged (set PartNum/Number/LinkTo, StructToRow, BINSERT, DispTextblockInfo)
```

Rules for the edit:

- `tLast` is a separate buffer on purpose. The insert copies the current block's `LeadIn` and `Data` from
  `TextblockRec`, so neither `TextblockRowToStruct` nor a read into `TextblockDataBuf` may run for the lookup.
- `buf(20)` is left out: `CLng(buf(20)) * &H1000000` overflows for values >= 128, and the key only covers the
  low 2 bytes anyway.
- Keep the routine's existing style (no indentation of the top-level body, `On Error GoTo error:` /
  `Call HandleError`). Keep CRLF line endings; do not use Git Bash `sed -i` on `.frm` files (it strips CRs).

### 4.3 Changelog (`README.md` and `docs/changelog.txt`, identical)

Add one line to the v1.8.4 block, after the last `-FIX:` line and before `-THANKS:`:

```
-FIX: Message and textblock editors' Insert now defaults to the next free number instead of current + 1 (#12)
```

`README.md` keeps its two trailing spaces; both files CRLF. Optionally credit the thread's analysis as
`(#12, Thanks Tehshortbus)` (open decision, see 6). Diff the two blocks afterwards.

## 5. Verification

1. Compile check: per the project notes, unattended `VB6.EXE /make` currently fails non-elevated, so build both
   EXEs in the IDE (vO as referenced; then swap in `modFieldmaps_vN.bas` for `nightmare_redux_n.exe`). Afterwards
   `git checkout -- exlimiter.oca` if a failed CLI run deleted it.
2. Manual test against a **copy** of `_fresh dats/MajorMUD v1.11p final` (point `WGPath` at the copy):

| Case | Expected |
|---|---|
| Message editor on 204, Insert | prompt shows `0`; OK inserts the next highest message number (thread reports 8614) and displays it |
| Textblock editor on 700, Insert | prompt shows `10004`; OK, part `0` inserts block 10004 part 0 and displays it |
| Textblock editor on 700, Insert, Cancel (first or second prompt) | `>>` moves to the block after 700, `<<` to the one before, not to 10002 / end-of-file error |
| Textblock Insert, type an existing number (e.g. 700, part 0) | duplicate-key error message, editor still on 700 |
| Repeat the textblock test with the n-build on `MajorMUD v1.11n` | same results |

## 6. Open decisions

- Credit Tehshortbus on the changelog line for the root-cause analysis in the issue thread?
- Rebuild and commit `nightmare_redux.exe` / `nightmare_redux_n.exe` with this fix now, or with the release build?
  The in-app changelog text (`frmHelpChangeLog.frx`, cp1252) is refreshed at the same time; it already lagged
  behind `v184`'s changelog before this branch.

## 7. Notes from the review pass

- A scripted Btrieve check (32-bit PowerShell P/Invoke into `WBTRV32.DLL` on copied dats) could not run here:
  every open returned status 20 (MicroKernel inactive), also with `W32MKDE.EXE` started by hand. Section 5's
  manual test is still the only runtime check.
- `txtNumber` / `txtPart` are locked, so the re-seek key always matches the block on screen.
- Out of scope: the Room editor's Insert and Copy prompts also default to room + 1 (`frmRoom.frm:6194`, `:6551`).
  Rooms have a map/room composite key, so this is a separate question from #12.
