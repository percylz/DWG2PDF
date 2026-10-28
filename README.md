# DWG2PDF — Batch DWG → PDF

Prints **every sheet (layout) of every DWG in a folder as its own PDF**, named after the sheet, with an optional common prefix.

For example, a DWG containing three sheets — `Ground Floor`, `First Floor`, `Roof Drainage` —
becomes three files with the prefix `A-`: `A-Ground Floor.pdf`, `A-First Floor.pdf`, `A-Roof Drainage.pdf`.

English | [简体中文](README.zh-CN.md)

> **Note on languages:** this README, the CLI flags and the log *file names* are in English.
> The GUI labels and the log *messages* inside the app are still in Chinese (an English UI is planned).

---

## 1. Quick start (recommended: the `.exe`)

Double-click **`DWG2PDF.exe`** — your browser opens the control page:

1. Pick the **folder that contains your DWGs** (sub-folders are included by default);
2. Ignore the **output folder** field: once step 1 is done it is filled in automatically as `<that folder>\PDF`,
   and the `PDF` folder is created for you when the run starts. Point it somewhere else with *Browse*,
   or press the **Default** button to go back to `<DWG folder>\PDF`;
3. Choose **one PDF per DWG**:
   - **Merge into a single multi-page PDF** (default) — one page per sheet, named after the DWG;
   - **One PDF per sheet** — named after the sheet;
4. Optionally tick **“LAYOUT sheet first, BOM sheet second when merging”** (**off by default since v2.3.0**, see below);
5. **Paper size** defaults to **A1**: a sheet smaller than A1 (an A3/A4 panel or system diagram, say) is scaled
   **up** to A1, which makes the text bigger and the print sharper; a sheet larger than A1 (an A0 layout) keeps
   its own size and is **never shrunk**. A0 / A2 / A3 / A4 are available, plus *Auto (use the sheet's own size)*;
6. Optionally set a **filename prefix**, e.g. `A-`;
7. Click **Start** and watch the log. When it finishes, open the output folder directly from the page.

> **Page-order rule (available since v2.0.1, off by default since v2.3.0):** when this is ticked and several
> sheets are merged into a single PDF, the app moves the
> **LAYOUT (arrangement) sheet to page 1** and the **BOM (bill of materials) sheet to page 2**; every other
> sheet keeps its original order. It looks at the sheet name (layout name / title-block title) first and then
> at the text inside that page (`LAYOUT` / `LAYOU` / `布置图`, `BOM` / `BILL OF MATERIAL` / `物料清单` / `设备清单`).
> If nothing is recognised, or more than one sheet matches the same type, the original order is kept and the log
> says so — the app never guesses. Tick the box or pass `--page-order` to enable it; without it the PDF simply
> follows the sheet order (the v2.3.0 default).

> **Paper size (v2.3.0):** size up, never down — when the paper is larger than the sheet the sheet is scaled up
> to fill it (bigger text, sharper print); when the sheet is larger than the chosen paper it keeps its own size
> (shrinking would only make it blurrier). The log prints *“plotting on the A1 paper you selected (841x594 mm,
> sheet … scaled up)”* once per paper size.

> The *Browse* button opens the app's own folder picker (large window, large text).
> Left side: common locations and drive letters. Middle: double-click to enter a folder, single click to select.
> You can also paste a path into the box at the top and press Enter. The bottom line tells you how many DWGs the current folder holds.

> The exe is a **single-file, no-install** build: copy it to a USB stick or another PC and double-click.
> **Python is not required** — that machine only needs AutoCAD installed.

> AutoCAD runs in the background during conversion (no visible window). The first launch takes about 30 seconds;
> after the log shows the AutoCAD connection, plotting runs at normal speed.
> Tick *Keep AutoCAD running* if you want the next run to start faster.

### Running from source (to try code changes immediately)

Double-click **`start.bat`**. It behaves exactly like the exe, but uses the local Python environment.

## 2. Command line / batch use

```bat
:: Basic: whole folder -> output directory
python dwg2pdf.py -s "D:\Projects\Architecture" -o "D:\Plot\PDF"

:: With a filename prefix
python dwg2pdf.py -s "D:\Projects\Architecture" -o "D:\Plot\PDF" -p "A-"

:: Resume: skip PDFs that already exist
python dwg2pdf.py -s "D:\Projects\Architecture" -o "D:\Plot\PDF" -p "A-" --skip-existing

:: Dry run: list what would be produced, plot nothing
python dwg2pdf.py -s "D:\Projects\Architecture" -o "D:\Plot\PDF" -p "A-" --dry-run

:: A single DWG works too
python dwg2pdf.py -s "D:\Projects\Architecture\Building-A.dwg" -o "D:\Plot\PDF"
```

### Options

| Option | Description |
|--------|-------------|
| `-s, --src` | Source folder (or a single `.dwg` file). Required. |
| `-o, --out` | Output folder. Required; created if missing. |
| `-p, --prefix` | Filename prefix. Default: empty. |
| `-t, --tpl` | Naming template. Default `{prefix}{layout}`. Placeholders: `{prefix}`, `{layout}` (sheet name), `{dwg}` (source file name), `{n}` (index). |
| `--include-model` | Also export model space (by default only layouts are plotted). |
| `--frames` | Force frame-based splitting in model space (one PDF per A0–A4 frame). |
| `--layouts` | Layouts only, never split by frame. |
| `--no-recursive` | Top level only, ignore sub-folders. |
| `--keep-subdir` | Mirror the source folder structure under the output folder (the GUI ticks this by default; the CLI keeps it opt-in). |
| `--skip-existing` | Skip PDFs that already exist (resume an interrupted run). |
| `--pc3` | Force a plot device. Default `DWG To PDF.pc3`. |
| `--one-pdf` | One multi-page PDF per DWG (the default; may be omitted). |
| `--split` | One PDF per sheet (the behaviour before v2.0). |
| `--paper A0..A4\|auto` | Output paper size, default `A1` (size up, never down); `auto` follows each sheet's own size |
| `--page-order` | Put the LAYOUT sheet on page 1 and the BOM sheet on page 2 when merging (off by default) |
| `--no-page-order` | Keep the original page order (the default; spelled out for scripts) |
| `--dry-run` | List the files that would be produced; plot nothing. |
| `--check` | Only check that AutoCAD is reachable and count the DWGs. |
| `--keep-acad` | Keep the AutoCAD process alive after conversion (faster next start). |
| `--json` | Machine-readable result output (for pipelines). |
| `--log FILE` | Also write the log to a file. |

Exit codes: `0` all succeeded, `2` some files failed, `3` environment/runtime error.

## 3. When one DWG holds several sheets

Two cases; the app decides automatically, and the **one PDF per DWG** switch decides how they are packed:

| Setting | Result | Filename |
|---------|--------|----------|
| **Merge into one PDF** (default) | One DWG → **one multi-page PDF**, one page per sheet, each page keeps its own paper size (mixed A0/A1/A3 is fine); the order is the sheet order (only reordered when *LAYOUT first, BOM second* is ticked) | the **DWG name** (prefix supported) |
| **One PDF per sheet** (old behaviour) | One DWG → N single-page PDFs | the **sheet name** (layout or title-block title) |

**1) Sheets live on layouts** → plotted layout by layout, then merged or kept separate as chosen.

**2) Everything is drawn in model space** (no usable layout) → the app splits by **title-block frame**,
each frame counting as one sheet (one page of the merged PDF):

- Detection: frames are blocks named `A0 / A1 / A2 / A3 / A4` (the largest closed rectangle inside the block is the frame).
  Without such names it falls back to rectangles whose size equals a standard sheet size × a common scale
  (e.g. 118900 × 84100 = A0 at 1:100).
- **Each PDF uses that sheet's own paper size** (A0 / A1 / A2 …), plotted by window with *fit to paper*,
  so a 1:100 drawing stays at 1:100 instead of being squeezed onto one common sheet size.
- **The sheet name comes from the title in that sheet's title block** (e.g. `Power Plan L1`); if unavailable it falls back to
  the drawing number (e.g. `FD-12`), then to `Sheet01`. In merge mode these names appear in the log as `第 n/m 张：…`.
- Frame splitting only kicks in when *all layouts are empty*, so drawings that already have layouts are never changed.
  Force it with `--frames`; force layout-only with `--layouts`.

Measured sample 1 (`<data-centre>---electrical.dwg`): 16 sheets in model space (5×A0, 9×A1, 2×A2) →
in merge mode one 16-page PDF whose pages are 1189×841, 841×594 and 594×420 mm.

Measured sample 2 (`1-IP DDC-AC- ICT2.dwg`, 13 layouts, all A4 landscape):
merge mode → one `1-IP DDC-AC- ICT2.pdf` (13 pages, 5.2 MB); split mode → 13 files `布局1.pdf` … `布局13.pdf`.
Pages follow the layout numbers exactly: **page 1 = 布局1 (`Panel Layout`), page 2 = 布局2 (the BOM), …,
page 13 = 布局13**. (Here the LAYOUT sheet happens to be layout 1 and the BOM layout 2, so the
“LAYOUT first, BOM second” rule needs no move at all.)

### Page order: the layout tab order (v2.1.1)

When merging, the page order is the **tab order of the layouts inside that DWG**:
`布局1` → page 1, `布局2` → page 2, … — exactly what you see in the tabs under the drawing area in AutoCAD.

- It is based on each layout's AutoCAD `TabOrder`, **not** on the alphabetical order of the names.
  That matters: the `Layouts` collection itself is enumerated alphabetically
  (`布局1, 布局10, 布局11, 布局12, 布局13, 布局2 …`), so using it directly puts `布局10` before `布局2` —
  which is what versions before v2.1.1 did.
- Layouts with non-numeric names (`Ground Floor`, `A-101`) also follow the tab order; nothing is guessed.
- If `TabOrder` cannot be read, the app falls back to a natural sort on the numbers in the name
  (`布局2` before `布局10`).
- Split mode plots in the same order, and the *“page n/m”* log lines follow it.

### Page-order boost: LAYOUT page 1, BOM page 2 (v2.0.1, on by default)

On top of the tab order, with this option ticked the app also moves the **LAYOUT (arrangement)** sheet to
page 1 and the **BOM (bill of materials)** sheet to page 2 — useful when those sheets live further down the
tab row. If they are already pages 1–2, nothing moves.

Before merging, every sheet is classified:

| Type | Keywords matched (case-insensitive, spaces/punctuation ignored) | Placed at |
|------|---------------------------------------------------------------|-----------|
| LAYOUT (arrangement) | `LAYOUT`, `LAYOU`, `布置图` | page 1 |
| BOM (bill of materials) | `BOM`, `BILL OF MATERIAL`, `物料清单`, `设备清单` | page 2 |
| Other | — | original order |

- The **sheet name** (layout name / title-block title) is checked first, then the **text on that PDF page**.
- If **more than one sheet matches the same type**, the app cannot tell which one to promote, so it
  **keeps the original order** and logs *“N LAYOUT sheets, M BOM sheets — page order not unique”*.
- A page containing **both** `LAYOUT` and `BOM` (a layout sheet whose title block prints a BOM field) counts as
  *other*, so it is never moved by mistake.
- Only the merged PDF is reordered; *one PDF per sheet* is unaffected.
- Enable: tick the box in the UI, or pass `--page-order` (off by default since v2.3.0).

## 4. Naming rules

**Merged (one PDF per DWG)** — the filename comes from the DWG:

| Case | Result |
|------|--------|
| Empty prefix | `1-IP DDC-AC- ICT2.pdf` |
| Prefix `A-` | `A-1-IP DDC-AC- ICT2.pdf` |
| A PDF with that name already exists | `Drawing (2).pdf` |

**One PDF per sheet** — the filename comes from the sheet:

| Case | Result |
|------|--------|
| Empty prefix | `Ground Floor.pdf` |
| Prefix `A-` | `A-Ground Floor.pdf` |
| Same sheet name in different DWGs | `Ground Floor.pdf` + `Ground Floor - Structure-B.pdf` (source name appended, no overwrite) |
| Want the source name too | `-t "{prefix}{dwg}-{layout}"` → `A-Structure-B-Ground Floor.pdf` |
| *Keep folder structure* ticked (on by default since v2.1.0) | Output mirrors the source tree, so identical names never collide |

Illegal filename characters (`\ / : * ? " < > |`) are replaced with `_`.

## 5. Output fidelity

- It goes through AutoCAD's own plot pipeline (`PlotToFile`), so **the result matches a manual PLOT**:
  paper size, plot scale, plot style table (CTB/STB), line weights, frame blocks and external references all follow the drawing's own settings.
- If the drawing's plot device is already a PDF device (`.pc3`), its page setup is used as-is.
  If it pointed at a printer (or nothing), the app switches to `DWG To PDF.pc3` and restores the original paper size where possible.
- **When the layout content does not sit on the paper** (in CAD the white sheet and the drawing look offset, or the
  sheet is simply empty), the window is re-framed on the real content and centred: printed **1:1 when it fits** the
  sheet, and scaled with *Fit to paper* when the content is larger than the sheet — the sheet always comes out
  complete, never cropped. The log says which case applied. Such a DWG often "looks fine" on another PC only
  because that machine's driver has a paper with the same name but a different size; this tool no longer depends
  on that and always recomputes from the content.
- **A sheet frame much smaller than the paper** (e.g. a 122×88 mm frame on an A4 page — only 41% of the page) is
  scaled up to fill the paper: below 80% of the paper the 1:1 rule is dropped and the frame is fitted to the page.
  Whenever a page has to be scaled, full-bleed media is used, otherwise the 5/17 mm printable margin of ordinary
  paper adds another white ring. At 92% or more the plot stays 1:1, so engineering scales are untouched.
- **Opened read-only**: the source DWG is never modified or saved, and nothing is written into the drawing
  (missing-font substitution happens in memory only).
- When a drawing has no plot style table, `monochrome.ctb` is used — otherwise colour 7 (white/black) plots white and is invisible in the PDF.
- Frame splitting plots *window = frame, paper = that sheet size, fit to paper*, so a 1:100 frame stays 1:100, exactly like window-plotting in CAD.
- By default only layouts are plotted, not model space. Enable *Export model space too* when needed.
  A DWG with model space only produces a single `Model.pdf`.

## 6. Requirements

| Item | Requirement |
|------|-------------|
| OS | Windows (developed on Windows 10/11) |
| CAD | AutoCAD 2013 or newer, **able to start normally and licensed** (developed against AutoCAD 2024). AutoCAD LT and other CAD apps (ZWCAD, GstarCAD…) are not supported — they expose no automation API |
| Python | **Not needed with the exe.** Source mode needs Python 3.9+ with `pywin32` and `pypdf` |

To run it on another PC, just copy `DWG2PDF.exe` (that PC still needs AutoCAD).

To rebuild after changing the source, double-click `build_exe.bat`; the output is `dist\DWG2PDF.exe` —
copy it back over the one in this folder. Rebuilding needs Python plus `pip install pyinstaller`.

## 7. Troubleshooting

| Symptom | Cause / fix |
|---------|-------------|
| Same problem persists after replacing the exe | The **old background process is probably still alive**: the new exe finds an instance already listening on port 8791 and simply points the browser at it, so the old code keeps doing the work. Since v1.9.1 the launcher compares version numbers and automatically shuts down a stale instance before starting; the UI also shows the version (e.g. `v1.9.1`) next to the title. On older builds, click **Quit** in the UI or end `DWG2PDF.exe` in Task Manager, then start the new exe |
| The drawing sits in one corner of the PDF, the rest blank, and edges are cut off | The drawing was set up for a non-PDF plot device (`Microsoft Print to PDF`, a printer driver…). Its paper name (`A4`) does not exist in DWG To PDF, so AutoCAD fell back to Letter portrait, and a fixed 1:100 scale then puts the sheet in a corner. Since v1.9 the app restores the paper by physical size, re-frames and centres the content, and uses full-bleed media for edge-to-edge sheets. On older versions, set the layout's plot device to `DWG To PDF.pc3` in AutoCAD and replot |
| No browser window after double-clicking the exe | Wait 2–5 s (the exe unpacks itself first). If nothing happens, check for `start_error.log` next to the exe, or open <http://127.0.0.1:8791/> manually |
| Antivirus flags it as suspicious | A common false positive for single-file PyInstaller builds — add an exclusion. If you would rather avoid it, run `start.bat` from source; identical functionality |
| Log stuck at "starting AutoCAD" for a long time | The first AutoCAD launch takes 30–60 s. Beyond 2 minutes, open AutoCAD manually once and confirm it starts without licence dialogs |
| "No .dwg files found" | Check the path; files in sub-folders require *Include sub-folders* |
| A DWG with a dozen sheets produced only two or three blank PDFs | Those sheets are in model space but the frames were not recognised. Make sure frame blocks are named `A0`–`A4` (or their size equals a standard sheet × an integer scale), or force `--frames` |
| A split sheet is slightly clipped | The frame window *is* the sheet boundary; anything drawn outside it never belonged to that sheet. Check the frame in CAD |
| A single sheet failed and produced no PDF | The app retries with `DWG To PDF.pc3` automatically (fixed in v1.1). If it still fails, the layout is genuinely empty, or its plot device does not exist here and cannot be substituted |
| Conversion succeeds but large parts of the PDF are blank (table borders present, no text) | The drawing references fonts that are not installed (e.g. `tssdeng.shx`, `hztxt.shx`, PKPM fonts). AutoCAD cannot measure that text and skips it. Since v1.4 missing fonts are temporarily substituted with SimSun + a GB big-font (in memory only); the log reports the substitution |
| Chinese characters show as `?` or boxes | The text style's font has no CJK glyphs — a drawing issue (CAD's own plot preview shows the same). Change the font |
| Part of the PDF is cut off | That layout's plot area is *Window* and the window is loose. The app reproduces it faithfully; reset the page setup to *Layout* (or redefine the window) in CAD and replot |
| "File is already open" | The DWG is open in AutoCAD. Close it and replot |
| "Cannot open / file not found" | File locked, corrupt, or saved by a newer CAD version (open it with the same or a newer version) |
| Every page of one DWG is blank, or only a small patch in the middle | The paper-space content of that DWG does not sit inside the paper rectangle (in CAD the sheet and the drawing are offset), so plotting with the saved page setup yields blank pages. Since v2.1.2 the window is recomputed from the actual content: 1:1 when it fits, *Fit to paper* when it does not — the log says which |
| Every page has a wide white border and the drawing sits in the middle | The **sheet frame of that DWG is only a fraction of the paper** (page setup does not match the sheet, e.g. a 122×88 mm frame on A4). Since v2.1.3 the frame is fitted to the paper with a 1 mm margin |
| An empty layout also produced a PDF | The layout really is blank; a blank page is correct |
| Every PDF popped open in a reader while converting | Fixed in v1.2: the app writes a `.part` temp file and renames it, so no reader is launched |
| `.part` files left in the output folder | Leftovers from a forced interruption — delete them. A re-run of the same sheet cleans them up |
| Merging into one PDF fails | It falls back to one file per sheet — nothing is lost. The log shows `！合并成一个 PDF 失败：…`. Usual cause: `pypdf` missing in source mode (`pip install pypdf`) |
| Conversion is slow, roughly 20 s per page | The sheet's original plot device cannot produce a PDF (a colleague's `M_PDF.pc3`, say, which does not exist on this machine): AutoCAD neither errors nor writes a file, so the app can only wait out the timeout and switch devices. Since v2.4.0 only the built-in PDF devices are trusted and everything else plots straight through `DWG To PDF.pc3` |
| Detection takes tens of seconds on a complex drawing | Drawings with thousands of block references/texts (2395 blocks, 2317 texts here) cost 1–3 ms per property read. Since v2.4.0 names and positions are read first and filtered afterwards: the 2395-block drawing went from 26 s to 13 s |
| Three or four sheets crammed onto one page, with no frame and no dividing lines | Those sheets are whole drawings packed into one block (frame + title block + content) and copied side by side; model space has no frame at all. Since v2.3.0 the app looks for same-name, same-size, non-overlapping blocks that cover ≥85% of their union and splits them one page each (4 side by side → 4 pages). If the blocks differ in name or size it still plots one page |
| The drawing only fills the left half of the page, the right half is blank | The page setup's plot window is larger than the paper (1690×841 mm on a 1191×843 mm A0), so a 1:1 plot covers only half the sheet. Since v2.3.0 the window is checked before plotting and re-framed from the layout content (49.6% → 100%) |
| The printed text is too small to read | Raise **Paper size** to A1 or A0: a sheet smaller than the paper is scaled up (an A4 detail on A1 is about 2.8×), while a larger sheet keeps its own size. Choose *Auto* to stick to each sheet's own paper |
| Page order is not what I expected | Since v2.1.1 the page order is the **layout tab order** (`布局1` → page 1, `布局2` → page 2, …), read from AutoCAD's `TabOrder`, so the alphabetical order of layout names no longer matters. The only exception is the *LAYOUT first, BOM second* option (on by default): those two sheets are pulled to the front when they are not already there |
| There is a BOM / LAYOUT sheet, but the order did not change | ① The option is off (checkbox or `--no-page-order`); ② neither the sheet name nor the page text contains a recognisable keyword; ③ more than one sheet matches the same type — the log says *“page order not unique, keeping sheet order”*; ④ that page contains **both** `LAYOUT` and `BOM`, so it counts as *other*. To force a sheet to the front, rename the layout or sheet title so it contains `LAYOUT` or `BOM` in CAD |
| The reorder promoted a page it should not have | Untick *“LAYOUT sheet first, BOM sheet second when merging”*, or pass `--no-page-order`, then convert again |
| Want to stop | Click *Stop*; the current sheet finishes first, then it stops |

## 8. Files

| File | Purpose |
|------|---------|
| `DWG2PDF.exe` | **Packaged executable — double-click to run, no install, no Python** (~10.6 MB) |
| `start.bat` | Start from source with the local Python environment |
| `dwg2pdf_ui.py` | Local GUI/HTTP service (127.0.0.1, reachable from this machine only) |
| `ui.html` | The UI page (folder picking, prefix, log and progress) |
| `dwg2pdf.py` | Core engine: scan DWGs, plot per layout/frame, CLI options |
| `build_exe.bat` | Rebuild the exe after code changes |
| `app.ico` | Application icon |
| `QUICKSTART.txt` | Portable quick-start notes — read this when you copy the exe to another PC |
| `run.log` | Runtime log generated automatically (page loads, clicks, exceptions); rotates to `run.log.old` above 512 KB |
| `start_error.log` | Only created when start-up fails; contains the detailed traceback |

To exit: click **Exit** in the UI (the exe has no console window, so this button is the way out).
With `start.bat` you can also just close the console window.

## 9. Version history

| Version | Notes |
|---------|-------|
| v2.4.0 | **Speed only** — output is unchanged (page count, paper size and page fill match v2.3.0 exactly). ①**No longer tries the sheet's original plot device first.** `KUL01-BLKA1-DR-E-10200…LAYOUT` points at a colleague's custom `M_PDF.pc3`, a plotter configuration that does not exist on this machine: `PlotToFile` returns without an error *and* without a file, so the app sat out the full 20 s timeout before falling back to `DWG To PDF.pc3` — 80 wasted seconds on a four-page drawing. It now only trusts the built-in PDF devices (`DWG To PDF.pc3`, `AutoCAD PDF`) and goes straight to `DWG To PDF.pc3` for anything else; the file-wait timeout also drops from 20 s to 2.5 s (`PlotToFile` is synchronous — when the device works the file is already there on return). ②**Sheet detection no longer asks AutoCAD about every single block reference.** Model space holds 2395 block references but only 42 distinct names; reading one name is a 2.8 ms cross-process COM call, so the old loop spent 6.6 s on names alone. Names are now read once and grouped, and the whole-sheet-block detector samples two instances per group before computing boxes for the rest. ③**The title block only reads the bottom-right corner.** The old code read all 2317 text entities in the drawing (position + content + height, ~1.3 ms each, 11.7 s); it now reads insertion points first and only fetches content for the few dozen inside the title-block area. Measured: DATAHALL LAYOUT 4 pages 113.5 s → **32.7 s**, DATAHALL MCB 4 pages 35.1 s → **23.1 s**; regression over 8 DWGs (4 Malaysia + 4 office/data-hall) keeps 4/4/6/3 pages and 100.1% / 99.3–99.7% page fill |
| v2.3.0 | ①**Sheets made of tiled “whole-sheet” blocks are split too.** `KUL01-BLKA1-SC-E-70001~3…MCB BOARD DETAILS` packs frame + title block + content into one block (`图签1`) and copies it four times side by side in model space — no outer frame, no dividing lines, no A0–A4 frame block, so the old code treated the whole model space as one sheet and squeezed four sheets onto one page. When neither frame blocks nor frame rectangles are found, the app now looks for blocks with the same name and size that do not overlap and together cover ≥85% of their own union (four 267.5×189.2 mm blocks here) and splits them one page each. ②**A plot window wider than the paper no longer wastes half a page.** `10200-C-0` in `KUL01-BLKA1-DR-E-10200…LAYOUT` has a 1690×841 mm window on a 1191×843 mm sheet, so the drawing only covered the left half (49.6%). The window is now compared with the paper before plotting and re-framed from the layout content when it overflows (49.6% → 100.1%). ③**Paper-space geometry wins over an over-wide viewport** when deciding what the frame is. ④New UI option **Paper size** (A0–A4, default **A1**, plus *Auto*): sheets smaller than the chosen paper are scaled up (bigger text, sharper print), sheets larger keep their own size. ⑤*LAYOUT first, BOM second* is now **off by default** (`--page-order` enables it). ⑥Long drawing numbers such as `KUL01-BLKA1-SC-E-70001-Z` are recognised as codes, with the title taken from the cell above. Measured: MCB sheets 1 page → 4 pages (A1, 99.7% page fill each), DATAHALL LAYOUT page 1 49.6% → 100.1%, the other three pages unchanged at 100% |
| v2.2.0 | ①**Landscape sheets are no longer cut in half.** `DWG To PDF` only exposes portrait media names (A0 is `841×1189`), and the old code compared the sheet against a landscape paper, so a 1:1 plot of a landscape sheet was clipped left/right and left wide white bands top/bottom (all 6 pages of `办公室-平面图20261023.dwg`). The real page size is now derived from **paper orientation + plot rotation**, and the 90° rotation is applied **after** the window is set (the other order makes the drawing lie on its side) — the page comes out landscape and 1:1 fills it. ②**Several sheets side by side inside one big frame are split by their title blocks.** In `办公室-系统图20261023.dwg` three sheets are drawn side by side in model space inside one decorative frame, with no dividing lines, so the old logic treated them as a single 955×299 sheet and squeezed them onto one page. When the frame is very elongated (aspect ≥1.8 or ≤0.55) the app now looks for title-block blocks that repeat with the same name and size (here three `图签1` at 268×189 mm) and splits the frame accordingly → 3 pages. ③If the page setup names a paper far too small (Letter page setup for an A0 drawing) the paper is re-picked from the sheet's own size. ④Standard-size frames are no longer discarded by the “ignore objects larger than 4× the paper” rule (that rule used to cut off the frame and title block). ⑤When no viewport looks like a sheet, the one whose area is closest to the paper is used instead of the union with the overview viewport. Measured on the four Malaysia DWGs: `办公室-平面图` 6 pages went from “white bands + clipped sides” to full-bleed A0 landscape at 100%, `办公室-系统图` 1 page → 3 pages, `数据中心-平面图` 4 pages, `数据中心-系统图` 1 page → 3 pages; regressions DPRM-1 (17 pages), DPRM-2 (8), ICT2 (13) keep the same page count, paper size and fill ratio as the previous version |
| v2.0 | ① **Selectable output layout**: the UI gained *one DWG →* with two choices — **merge into one multi-page PDF** (one page per sheet) or **one PDF per sheet**. In merge mode the file is named after the DWG, page order follows sheet order, and every page keeps its own paper size (mixed A0/A1/A3 stays mixed). CLI: `--one-pdf` (default) / `--split`. Merging uses pypdf; if the merge fails the app falls back to one file per sheet, so no result is lost. ② **Default output folder renamed** `PDF输出` → `PDF` |
| v2.1.3 | **Pages no longer half-empty** (`1-IP DDC-DPRM-2.dwg` was 41%, nine pages of `2-IP DDC-AC- ICT2.dwg` were 39%). Two causes: ①besides the sheet viewport the layout also holds a 1:1 overview viewport, which the old window logic included, shrinking the whole page; ②the sheet frame itself is only 17% of the paper (122×88 mm frame on A4) and was still plotted 1:1, so the drawing sat in the middle. Now the sheet is the viewport whose aspect ratio matches the paper **and** whose area is 0.5–2× the paper (when a page really holds several tiled sheets their union is used — nothing is cropped); a frame below 80% of the paper is fitted to fill it; any scaling is done on full-bleed media with a 1 mm margin. Measured: DPRM-2 41%→96% on all 8 pages, 2-IP ICT2 39%→96% on all 16, DPRM-1 83%→99%; regression on `1-IP DDC-AC- ICT1.dwg` is identical page by page (0.00% ink difference) |
| v2.1.2 | **Blank pages / half-cut sheets fixed.** In `1-IP DDC-DPRM-1.dwg` the paper-space content (the sheet viewport) is nowhere near the A4 paper rectangle of its page setup, so the old window logic (intersection of content and paper) came out empty and the plot fell back to the saved page setup — all 17 pages were blank. Now ①the window is framed on the real content: **1:1 when it fits, "fit to paper" when it does not**, so pages are no longer blank; ②when a layout holds several viewports, the one whose aspect ratio matches the paper (the largest) is treated as the sheet, so an oversized overview viewport no longer squeezes the drawing; ③if a plot still comes out blank it is re-plotted once with the recomputed window. Verified: 17 pages went from all-blank to complete A4-landscape sheets; regression on `1-IP DDC-AC- ICT2.dwg` matches the previous 13 pages within ±0.6% ink |
| v2.1.1 | **Page order now equals the layout number.** AutoCAD enumerates the `Layouts` collection alphabetically (`布局1, 布局10, 布局11, 布局12, 布局13, 布局2 …`); earlier versions used that order directly, which pushed `布局10` onto page 2 and `布局2` down to page 6. Pages are now sorted by each layout's **`TabOrder`** — `布局1` page 1, `布局2` page 2, … matching the tabs in AutoCAD; if `TabOrder` cannot be read the app falls back to a natural sort on the numbers in the name. Side note: the v2.0.1 report “BOM moved from page 6 to page 2” was this very bug (the BOM sheet is `布局2`); that rule now only kicks in when LAYOUT/BOM sheets sit further down the tab row. Verified on `1-IP DDC-AC- ICT2.dwg`: all 13 merged pages match `布局1.pdf` … `布局13.pdf` one by one |
| v2.1.0 | ①**The output folder now follows the DWG folder**: picking the DWG folder auto-fills the output as `<that folder>\PDF`, which is created on start. It keeps following when you switch folders (for a single-DWG input, the DWG's own folder is used). A manually chosen output folder is never overwritten — the new **Default** button restores `<DWG folder>\PDF`. ②**“Mirror the source folder structure” is now ticked by default**, so sheets with the same name in different sub-folders no longer land side by side |
| v2.0.1 | **Page-order rule**: when merging, the **LAYOUT (arrangement) sheet is moved to page 1** and the **BOM sheet to page 2**, all other sheets keep their order. Classification uses the sheet name (layout name / title-block title) and then the text on that PDF page (`LAYOUT`/`LAYOU`/`布置图`, `BOM`/`BILL OF MATERIAL`/`物料清单`/`设备清单`); if a type matches more than once, or a page contains both keywords, the original order is kept and the log explains why. New UI checkbox (on by default) and `--no-page-order` to disable. Verified on `1-IP DDC-AC- ICT2.dwg`: `Panel Layout` stays on page 1, `BOM LIST` moves from page 6 to page 2 |
| v1.9.1 | Fixed *"still broken after replacing the exe"*: the launcher only checked whether something was listening on port 8791 and, if so, opened a browser page pointing at it — so an **old process with old code** kept doing the actual work. The launcher now reads that instance's version: same version is reused, a different version is shut down and replaced. `run.log` starts with `程序启动 v1.9.1 (PID …, port …)` and the UI shows the version next to the title, so you can always tell which build is running. Verified with the packaged exe on `1-IP DDC-AC- ICT2.dwg`: 13 of 13 sheets, A4 landscape, 96 %×96 % coverage |
| v1.9 | Fixed *"the PDF has the drawing squeezed into one corner, the rest blank, and edges cut off"*. Those drawings were set up for a non-PDF plot device (`Microsoft Print to PDF`, a real printer driver…). After switching to `DWG To PDF.pc3`, the old paper name (`A4`) does not exist there, so AutoCAD silently fell back to Letter portrait — and since the 1:100 scale is fixed, the sheet landed in a corner and got clipped. Now: ① paper is restored by **physical size**, not by name; ② the plot window is re-framed from the real paper-space content and **centred**, with the paper orientation following the drawing (a landscape sheet gets landscape paper instead of being forced in with a 90° rotation); ③ drawings that bleed to the paper edge automatically use the **full-bleed** media, otherwise DWG To PDF's 5/17 mm printable margins cut the edge strip (title box, sheet number). Scale is never changed. Measured on `1-IP DDC-AC- ICT2.dwg` (13 layouts): from "one corner only" to all 13 as A4 landscape at 96 %×96 % coverage with complete border and title block |
| v1.8 | Fixed *"frame detection error: Call was rejected by callee"*, which reduced a multi-sheet DWG to a single PDF. ① The polyline scan re-read `ss.Count` on every iteration (one COM call per entity; AutoCAD could reject it while busy, and the call sat outside the `try`, aborting the whole scan) — the count is now read once. Block references, polylines and text reads now retry when AutoCAD is busy and skip only the bad entity. ② Detection retries up to 3 times (2 s / 4 s), and if it still fails the log states clearly that it falls back to layout plotting. Verified: all 13 architectural sheets plotted, no errors |
| v1.7 | Fixed the `Cannot set properties of null (setting 'className')` error after ticking *Keep folder structure*: the status-bar refresh rewrote its inner dot, so the `id` lookup failed afterwards — now only the text node is updated. Also fixed frame detection hanging for minutes on architectural drawings: block definitions are no longer scanned for every block name (anonymous blocks can hold thousands of entities, ~10 s each); names are pre-filtered first, with progress every 100 blocks / 200 polylines (5-minute hang → 13 sheets detected in 23 s) |
| v1.6 | Fixed *"clicking Start does nothing / looks frozen"*: ① the first three log lines were swallowed (wrong page cursor) — the pending-file count now appears immediately; ② background exceptions no longer close the connection silently but surface as an error line and are written to `run.log`; ③ after 1.5 s without a response the log says AutoCAD is starting; ④ a second double-click reuses the running service instead of starting another; ⑤ the page self-checks every 5 s and shows a red banner if the backend is gone. Also fixed cold-start timing: the first sheets could fail on `Open.Layouts` / `<unknown>.Item` while AutoCAD was still initialising — it now waits for readiness and retries (5/5 on a cold start) |
| v1.5 | Replaced the tiny system folder dialog with an in-app picker: larger window, 30 % larger text, shortcuts for common locations and drives, paste-a-path-and-Enter navigation, and a live count of DWGs in the current folder |
| v1.4 | Fixed *"conversion succeeds but the content is blank"*: ① set the window before switching to window-plot mode (otherwise: invalid input); ② recognise non-standard sheet sizes such as A1-extended; ③ substitute locally missing SHX fonts (otherwise whole text blocks are skipped); ④ fall back to a monochrome plot style so white text stays visible |
| v1.3 | Multiple sheets per DWG: split model space by frame into several PDFs, each using its own A0–A4 paper size, filename from the title block (only when layouts are empty; drawings with layouts are untouched) |
| v1.2 | Fixed *"a PDF reader pops up after every sheet"*: AutoCAD opens the plot result by extension; the app now writes a `.part` file and renames it when done — no pop-ups, same speed |
| v1.1 | Fixed *"conversion fails silently"* on some project drawings: the plot device name is now passed **explicitly** (AutoCAD sometimes writes no file without it) and the app retries with `DWG To PDF.pc3` when no file appears. Verified on 3 real A0 project drawings |
| v1.0 | First release: folder scan, per-layout plotting, prefix naming, duplicate handling, resume, GUI and exe packaging |
