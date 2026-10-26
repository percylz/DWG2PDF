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
2. Pick the **output folder** (leave empty → `DWG folder\PDF输出`);
3. Optionally set a **filename prefix**, e.g. `A-`;
4. Click **Start** and watch the log. When it finishes, open the output folder directly from the page.

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
| `--keep-subdir` | Mirror the source folder structure under the output folder. |
| `--skip-existing` | Skip PDFs that already exist (resume an interrupted run). |
| `--pc3` | Force a plot device. Default `DWG To PDF.pc3`. |
| `--dry-run` | List the files that would be produced; plot nothing. |
| `--check` | Only check that AutoCAD is reachable and count the DWGs. |
| `--keep-acad` | Keep the AutoCAD process alive after conversion (faster next start). |
| `--json` | Machine-readable result output (for pipelines). |
| `--log FILE` | Also write the log to a file. |

Exit codes: `0` all succeeded, `2` some files failed, `3` environment/runtime error.

## 3. When one DWG holds several sheets

Two cases; the app decides automatically.

**1) Sheets live on layouts** → one PDF per layout, named after the layout (classic behaviour).

**2) Everything is drawn in model space** (no usable layout) → the app splits by **title-block frame**, one PDF per frame:

- Detection: frames are blocks named `A0 / A1 / A2 / A3 / A4` (the largest closed rectangle inside the block is the frame).
  Without such names it falls back to rectangles whose size equals a standard sheet size × a common scale
  (e.g. 118900 × 84100 = A0 at 1:100).
- **Each PDF uses that sheet's own paper size** (A0 / A1 / A2 …), plotted by window with *fit to paper*,
  so a 1:100 drawing stays at 1:100 instead of being squeezed onto one common sheet size.
- **The filename comes from the title in that sheet's title block** (e.g. `Power Plan L1`); if unavailable it falls back to
  the drawing number (e.g. `FD-12`), then to `Sheet01`.
- Frame splitting only kicks in when *all layouts are empty*, so drawings that already have layouts are never changed.
  Force it with `--frames`; force layout-only with `--layouts`.

Measured sample (`<data-centre>---electrical.dwg`): 16 sheets in model space (5×A0, 9×A1, 2×A2) →
16 PDFs at 1189×841, 841×594 and 594×420 mm, every filename taken from the drawing title.

## 4. Naming rules

| Case | Result |
|------|--------|
| Empty prefix | `Ground Floor.pdf` |
| Prefix `A-` | `A-Ground Floor.pdf` |
| Same sheet name in different DWGs | `Ground Floor.pdf` + `Ground Floor - Structure-B.pdf` (source name appended, no overwrite) |
| Want the source name too | `-t "{prefix}{dwg}-{layout}"` → `A-Structure-B-Ground Floor.pdf` |
| *Keep folder structure* ticked | Output mirrors the source tree, so identical names never collide |

Illegal filename characters (`\ / : * ? " < > |`) are replaced with `_`.

## 5. Output fidelity

- It goes through AutoCAD's own plot pipeline (`PlotToFile`), so **the result matches a manual PLOT**:
  paper size, plot scale, plot style table (CTB/STB), line weights, frame blocks and external references all follow the drawing's own settings.
- If the drawing's plot device is already a PDF device (`.pc3`), its page setup is used as-is.
  If it pointed at a printer (or nothing), the app switches to `DWG To PDF.pc3` and restores the original paper size where possible.
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
| Python | **Not needed with the exe.** Source mode needs Python 3.9+ with `pywin32` |

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
| An empty layout also produced a PDF | The layout really is blank; a blank page is correct |
| Every PDF popped open in a reader while converting | Fixed in v1.2: the app writes a `.part` temp file and renames it, so no reader is launched |
| `.part` files left in the output folder | Leftovers from a forced interruption — delete them. A re-run of the same sheet cleans them up |
| Want to stop | Click *Stop*; the current sheet finishes first, then it stops |

## 8. Files

| File | Purpose |
|------|---------|
| `DWG2PDF.exe` | **Packaged executable — double-click to run, no install, no Python** (~9 MB) |
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
