# D‑Forge — SCORM Preparation & Publishing GUI

**D‑Forge** is a single-file PowerShell WPF application that automates common SCORM course preparation tasks (patching, compliance checks, score calculation/verification) and zips courses for LMS ingestion. It provides a small GUI so non‑technical users can run checks and publish packages without opening a console.

---

## Table of contents

* [Purpose](#purpose)
* [Features](#features)
* [Prerequisites](#prerequisites)
* [Installation](#installation)
* [How it works (high level)](#how-it-works-high-level)
* [GUI & Buttons (explainers)](#gui--buttons-explainers)
* [Typical workflow / examples](#typical-workflow--examples)
* [Troubleshooting & common errors](#troubleshooting--common-errors)
* [Development notes](#development-notes)

---

## Purpose

This script was built to speed up and standardize preparation of HTML5 courses (published from authoring tools like Articulate Storyline, Captivate, or custom players) for SCORM-compliant LMS delivery. The GUI removes the need for users to run multiple manual checks and repetitive steps (inserting SCORM fetch code, ensuring scormAPI inclusion, verifying scoring/completion settings, and packaging the course).

## Features

* WPF GUI (RichTextBox log with colored messages)
* Hide console on start for a standalone feel
* **Zip Only** — verify scoring & compliance then zip
* **Single Publish** — patch, debug (open in Chrome), verify, and zip a single course folder
* **Bulk Publish** — process all subfolders of a parent folder, debug, bulk verify, and zip
* **Check Completion (Single/Bulk)** — detect completion triggers and explicit status strings in `data.js`
* **Check Score (Single/Bulk)** — parse `html5/data/js/data.js` and sum all `maxpoints` values
* Folder compliance checks: presence of `index_lms.html`, `index.html` (or `story.html`), `scormAPI.min.js`, and script injection checks
* Safe modifications: inserts script tag when missing, copies `scormAPI.min.js` from the script's directory if available

## Prerequisites

* Windows with PowerShell (Recommended: PowerShell 7+ for best experience, though this script is compatible with Windows PowerShell in most cases)
* .NET/WPF support (present on standard Windows desktop)
* Optional: Google Chrome (for the debug workflow). Update `$ChromePath` at the top of the script if Chrome is installed in a different location.

## Installation

1. Save `D-Forge-SCORM-GUI-Fixed.ps1` (or your preferred name) to a folder, for example `D:\work\Script_tools\`.
2. Place `scormAPI.min.js` beside the script if you want the script to copy it automatically to course folders when missing. Otherwise, the script will require it in each folder.
3. (Optional) Right-click the `.ps1` file, `Properties` → `Unblock` if Windows blocks downloaded scripts.
4. Run the script by double-clicking (if you have PowerShell execution configured to allow it) or from a PowerShell session: `powershell -ExecutionPolicy Bypass -File "D:\work\Script_tools\D-Forge-SCORM-GUI-Fixed.ps1"`.

> The script hides the console window at startup so only the GUI is visible. When the window closes, the PowerShell process exits.

## How it works (high level)

1. The GUI collects a path (either a course folder or a parent folder with many course folders).
2. Depending on the chosen operation, the script runs one or more checks and actions:

   * Validate `data.js` scoring configuration (presence of `scorings` with `type:"action"`). This is critical for LMS "completed" status.
   * Parse `data.js` to calculate the total possible score by summing `maxpoints` fields.
   * Ensure required files exist and that `index_lms.html` contains an injected hook for SCORM data collection. If the script finds a missing `<script>` tag referencing `scormAPI.min.js`, it inserts it.
   * Copy `scormAPI.min.js` from the script folder into target folders (if available) when missing.
   * Optionally open `index_lms.html` files in Chrome so the user can check the console for messages produced by the injected script.
   * Compress the contents of the course folder into `ZippedFiles\<foldername>.zip` (zips contents, not parent folder).

## GUI & Buttons (explainers)

* **Target Folder Path**: Enter a single course folder (for Zip Only or Single Publish) or a parent folder containing subfolders (for Bulk operations).
* **Browse**: Opens a folder picker.

Primary action buttons (7 total):

1. **Zip Only** — Performs scoring compliance check and folder compliance check. Prompts to verify calculated score. If verified, zips the course into `ZippedFiles` next to the course.
2. **Single Publish** — Runs `Process-ScormFolder` (scoring check, copy `scormAPI.min.js`, patch `index_lms.html` safely, rename `story.html` to `index.html` if needed). Optionally opens Chrome for debugging, then verifies score and zips.
3. **Bulk Publish** — Processes every subfolder of the provided parent folder using the same `Process-ScormFolder` logic, optionally opens all modified `index_lms.html` files in Chrome for manual debug, performs a single bulk score verification dialog listing the calculated totals for all processed folders, and zips approved folders.
4. **Check Completion (Single)** — Parses `data.js` in the single folder to report whether a `type:"action"` scoring trigger (completed trigger) exists and whether any explicit completion/status strings are present.
5. **Check Score (Single)** — Calculates and displays the total sum of `maxpoints` values for the selected single folder.
6. **Check Completion (Bulk)** — Runs the completion check across all first-level subfolders and shows a report.
7. **Check Scores (Bulk)** — Calculates scores for all subfolders and shows a report.

* **Exit** — Closes the GUI and exits the script/process.

## Core functions & behavior (mapping to code)

* `Append-ColoredText` — Writes color-coded messages to the RichTextBox log.
* `Get-ScormTotalScore` — Reads `html5/data/js/data.js` and aggregates all `"maxpoints": <number>` entries.
* `Get-ScormCompletionInfo` — Looks for `scorings` array entries with `"type":"action"` and common completion/status fields.
* `Test-ScormScoringCompliance` — Verifies the presence of completion trigger in `data.js` (critical check).
* `Test-ScormFolderCompliance` — Quick checks for `scormAPI.min.js`, `index_lms.html`, and the injected script presence.
* `Process-ScormFolder` — High-level orchestration for single-folder fixes and patches.
* `Compress-FolderContent` — Creates `ZippedFiles\<foldername>.zip` containing the folder contents.

## Typical workflow / examples

**Publish a single folder**

1. Enter the course folder in the path box (or click Browse).
2. Click **2. Single Publish**.
3. Follow prompts to open Chrome and verify console messages.
4. Approve the calculated score when prompted to create the zip.

**Publish many courses at once**

1. Enter the parent folder that contains many course subfolders.
2. Click **3. Bulk Publish**.
3. Optionally open all modified files in Chrome for debugging.
4. Review the single bulk verification dialog and approve zipping.

**Quick checks**

* If you only need to verify scoring or completion, use the Check Completion / Check Score buttons for quick readouts without making changes.

## Troubleshooting & common errors

* **Parser errors pointing at bracketed text (e.g. `[2] Checking...`)**: This happens when a logging function used earlier was undefined or the script contains malformed tokens. This repository provides a fixed single-file named `D-Forge-SCORM-GUI-Fixed.ps1` that replaces inconsistent logging calls and uses safe string/escaping approaches.
* **`data.js` not found**: Ensure that the course follows the folder structure `html5/data/js/data.js`. Some publishers use alternate structure — adapt the script or move/point to the correct path.
* **`scormAPI.min.js` missing**: Put `scormAPI.min.js` beside the script, or ensure it exists in every course folder. The script will attempt to copy it from its own directory if found.
* **Chrome not found / not opening**: Update the `$ChromePath` variable to fit your installation.
* **Encoding problems / strange characters**: Save the script using UTF-8 (without BOM is safest) and ensure `Get-Content` reads files as UTF8, which the script attempts to do.

## Development notes

* The script is intentionally cautious about auto-editing `index_lms.html`. It inserts the script tag if missing and logs missing SCORM fetch code; a more aggressive insertion routine exists in earlier iterations but was simplified to reduce risk of breaking generated player code.
* The score parsing uses a simple regex for integers; if `maxpoints` uses floats or different formats, update the regex accordingly.
* The script is developed as a single `.ps1` file to make deployment easy for end users (no dot-sourcing or modules required).

