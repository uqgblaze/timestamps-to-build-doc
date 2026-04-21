# Build Document Generator

Generates a lecture build document from a timestamps table and a VTT transcript.

- **V1** — Produces a Markdown file where each slide title becomes a heading and the matching transcript text is inserted as a paragraph beneath it.
- **V2** — Optionally inserts the same content (title + transcript) into the speaker notes of the corresponding PowerPoint slides.

---

## Prerequisites

### 1. Python 3.10 or later

Download from [python.org](https://www.python.org/downloads/).

During installation, tick **"Add Python to PATH"** — this is required for the `.bat` launcher to work.

To verify your installation, open a terminal and run:

```
python --version
```

### 2. python-pptx

Required for the V2 PowerPoint speaker notes feature. Install via pip:

```
pip install python-pptx
```

To verify:

```
python -c "import pptx; print(pptx.__version__)"
```

> **Note:** `tkinter` (the GUI framework) is included with standard Python on Windows. No separate install is needed.

---

## Files

| File | Description |
|---|---|
| `build_doc_generator.py` | Main application |
| `Build Doc Generator.bat` | Double-click launcher for Windows File Explorer |

---

## How to Run

**Double-click** `Build Doc Generator.bat` in File Explorer.

Alternatively, run directly from a terminal:

```
python build_doc_generator.py
```

---

## Input File Requirements

The tool expects the following files to be present in the selected working directory:

### Timestamps file (`*timestamps*.md`)

A Markdown table with four pipe-delimited columns: `Start`, `-`, `Stop`, `Title`.

```
| Start | - | Stop | Title |
| :---: | :---: | :---: | :--- |
| 0:00:00 | - | 0:05:10 | 1. Joint Ore Reserves Committee (JORC) Code |
| 0:05:10 | - | 0:05:37 | 2. Prospecting |
```

- Time format: `H:MM:SS` or `HH:MM:SS` (milliseconds optional)
- Title numbering (`1.`, `2.`, etc.) is used to match segments to PowerPoint slide numbers in V2
- The file must contain `timestamps` somewhere in its filename (e.g. `lecture-timestamps.md`)

### VTT transcript file (`.vtt`)

A standard WebVTT subtitle/transcript file exported from your video or audio processing tool.

```
WEBVTT

00:00:01.000 --> 00:00:10.000
Hello and welcome to this module...

00:00:10.000 --> 00:00:31.000
So I'd just like you to imagine...
```

### PowerPoint file (`.pptx`) — V2 only

The source presentation whose speaker notes will be written. The tool writes to a **copy** and never modifies the original.

Slide matching is by number: the segment titled `1. Introduction` maps to Slide 1, `2. Background` to Slide 2, and so on. Slides that have no matching segment are left untouched.

---

## Usage

1. Launch the app via the `.bat` file or `python build_doc_generator.py`.
2. Click **Browse…** and select the folder containing your timestamps, VTT, and (optionally) PPTX files.
3. The app auto-detects all matching files and populates the dropdowns. Select the correct file in each dropdown if multiple are present.
4. Adjust the **Document title** if needed (defaults to the VTT filename stem).
5. Confirm or change the **output Markdown path** (defaults to `<stem>-build-doc.md` in the same folder).

### For V2 (PowerPoint notes):

6. Tick **"Also insert notes into the selected PowerPoint"**.
7. Confirm the source PPTX and the output PPTX path (defaults to `<original>-with-notes.pptx`).
8. Click **Generate**.

The log panel at the bottom shows progress, cue counts, and any warnings (e.g. segments whose slide number falls outside the deck's range).

---

## Output

### Markdown build document

A `.md` file structured as:

```markdown
# Document Title

## 1. Slide Title

Full transcript paragraph for this segment...

## 2. Next Slide Title

Full transcript paragraph for this segment...
```

### PowerPoint with speaker notes (V2)

A copy of the source `.pptx` where each matched slide's notes panel contains:

```
Slide Title          ← bold
                     ← blank line
Full transcript paragraph for this segment...
```

The original `.pptx` is never modified.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| Double-clicking the `.bat` opens Notepad | Right-click → Open with → Windows Command Processor |
| `python` is not recognised | Re-run the Python installer and tick "Add Python to PATH" |
| `ModuleNotFoundError: pptx` | Run `pip install python-pptx` in a terminal |
| No files detected after Browse | Check the folder contains a `*timestamps*.md` and a `.vtt` file |
| Slide notes not appearing in PowerPoint | Ensure slide numbers in titles (`1.`, `2.`, …) match the deck's slide order |
| Warning: slide out of range | The segment number exceeds the total slide count — check the timestamps table |
