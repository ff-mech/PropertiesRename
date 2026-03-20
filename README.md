<div align="center">

# PropertiesRename

**SOLIDWORKS VBA Macro — Batch Property & Revision Table Updater**

![Language](https://img.shields.io/badge/Language-VBA-blue?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-SOLIDWORKS-red?style=flat-square)
![Type](https://img.shields.io/badge/Type-Macro-green?style=flat-square)
![Maintained](https://img.shields.io/badge/Maintained-Yes-brightgreen?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-yellow?style=flat-square)

*Scans an entire folder of SOLIDWORKS files and silently updates custom properties and revision tables — no manual file-by-file editing required.*

</div>

---

## Contents

- [Overview](#overview)
- [What It Does](#what-it-does)
- [How to Use](#how-to-use)
- [Dialog Fields](#dialog-fields)
- [Requirements](#requirements)
- [Log File](#log-file)
- [Behavior & Edge Cases](#behavior--edge-cases)
- [License](#license)
- [Latest Release](https://github.com/ff-mech/PropertiesRename/releases/latest)

---

## Overview

PropertiesRename eliminates the tedious task of manually updating `DrawnBy`, `DwgDrawnBy`, and revision table entries across a batch of SOLIDWORKS files. A single dialog collects all inputs upfront, then the macro processes the entire folder automatically — opening each file, applying the correct changes, saving, and closing with no further interaction.

> **Tip:** Works on folders of any size. Progress is tracked in a full log file written to the same folder when the run completes.

---

## What It Does

| File Type | Extension | Action |
|---|---|---|
| **Part** | `.SLDPRT` | Sets the `DrawnBy` custom property |
| **Assembly** | `.SLDASM` | Sets the `DrawnBy` custom property |
| **Drawing** | `.SLDDRW` | Enforces all 15 standard custom properties, updates `DwgDrawnBy`, and inserts **Rev A — INITIAL RELEASE** into the revision table |

> All 15 drawing properties are written back in the exact order required by the Foxfab drawing template. Existing values on properties other than `DwgDrawnBy` are preserved.

---

## How to Use

**Step 1** — Open SOLIDWORKS and run the macro:
> **Tools → Macros → Run** → select `PropertiesRename.swp`

**Step 2** — Fill in the dialog and click **Run**.

**Step 3** — Review the confirmation summary showing the exact file counts and settings, then click **Yes** to begin processing.

**Step 4** — When finished, a results popup displays the outcome. A full log is saved automatically to the target folder.

---

## Dialog Fields

| Field | Description |
|---|---|
| **Folder Path** | Full path to the folder containing the SOLIDWORKS files |
| **DrawnBy Initials** | Initials written to `DrawnBy` on parts and assemblies |
| **DwgDrawnBy Initials** | Initials written to `DwgDrawnBy` on drawings |
| **Skip files starting with `003-`** | When checked, any file whose name begins with `003-` is skipped entirely |

> All three text fields are required. The macro will not proceed if any of them are left blank.

---

## Requirements

- SOLIDWORKS with macro execution enabled
- Network access to the Foxfab revision table template:

```
\\npsvr05\FOXFAB\FOXFAB_DATA\ENGINEERING\SOLIDWORKS\Foxfab Templates\Revision Table v1.1.sldrevtbt
```

> If the template cannot be reached, drawings without an existing revision table will fail and be reported in the log.

---

## Log File

A log named `DrawnBy_Update_Log.txt` is written to the target folder at the end of every run. It is structured as follows:

| Section | Contents |
|---|---|
| **Failures** | Files that could not be opened, processed, or saved |
| **Skipped** | Read-only files and files that already had the correct values |
| **Updated** | Every file that was successfully modified, with a summary of what changed |
| **Summary** | Total counts (updated / skipped / failed) and elapsed time |
| **Debug Log** | A detailed per-file trace of every operation — useful for troubleshooting |

---

## Behavior & Edge Cases

- **Already-open files** — If a file is currently open in SOLIDWORKS, the macro closes it first, then reopens it silently for processing.
- **Read-only files** — Detected before opening and skipped automatically. Reported in the log under *Skipped*.
- **Property preservation** — On drawings, all 14 non-`DwgDrawnBy` properties retain their existing values. Only `DwgDrawnBy` is overwritten with the new initials.
- **Property ordering** — All 15 properties are deleted and re-added in the correct template order even if they already existed, ensuring the drawing template links function correctly.
- **Existing revision rows** — If a drawing already has a revision table, all existing revision rows are cleared before Rev A is added.
- **Midnight rollover** — Elapsed time calculation handles runs that cross midnight correctly.


---

## License

Released under the [MIT License](LICENSE).
