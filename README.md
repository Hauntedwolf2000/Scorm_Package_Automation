# D-Forge SCORM Automation Tool - Technical Documentation

## Overview

D-Forge is a GUI-based PowerShell automation tool designed to streamline the preparation, validation, and packaging of SCORM (Sharable Content Object Reference Model) e-learning content. The tool automates repetitive tasks involved in making Articulate Storyline exports LMS-ready, with built-in compliance checking and quality assurance features.

## What It Does

### Core Functions

1. **SCORM File Preparation**
   - Injects required SCORM API tracking code into HTML files
   - Copies and integrates the `scormAPI.min.js` library
   - Renames `story.html` to `index.html` for LMS compatibility
   - Validates folder structure and file compliance

2. **Quality Assurance**
   - **Completion Tracking**: Verifies the presence of completion triggers (`type: "action"`) in course data
   - **Score Validation**: Calculates total possible scores by summing `maxpoints` values from course questions
   - Performs pre-zip compliance checks to catch issues early

3. **Packaging**
   - Creates LMS-ready ZIP files using Windows Shell compression
   - Organizes output files in a `ZippedFiles` subdirectory
   - Batch processing capability for multiple courses

4. **Debugging Support**
   - Opens courses in Chrome for SCORM API console verification
   - Provides real-time logging with color-coded messages
   - Interactive verification dialogs for score confirmation

### Operation Modes

| Button | Mode | Description |
|--------|------|-------------|
| **1. Zip Only** | Quick Package | Validates compliance and zips without modifications |
| **2. Bulk Publish** | Batch Process | Processes multiple courses in one folder, with bulk debugging and zipping |
| **3. Single Publish** | Individual Fix | Prepares one course with full fix → debug → verify → zip workflow |
| **4. Check Completion (Single)** | QA Tool | Verifies completion trigger exists in one course |
| **5. Check Score (Single)** | QA Tool | Calculates and displays total score for one course |
| **6. Check Completion (Bulk)** | QA Tool | Checks completion triggers across all subfolders |
| **7. Check Scores (Bulk)** | QA Tool | Generates score report for all subfolders |

## How It Works

### Technical Architecture

**Language**: PowerShell 5.1+ with WPF (Windows Presentation Foundation)  
**GUI Framework**: XAML-based interface  
**Compression**: Windows Shell.Application COM object  
**Console Hiding**: Win32 API calls via C# interop

### Development Approach

The tool follows a **defensive programming** pattern with extensive validation:

```
Input Validation → Compliance Check → Modification → Verification → Output
```

#### Key Technical Decisions

1. **Console Window Hiding**
   ```powershell
   Add-Type -TypeDefinition $consoleHelperCSharp
   [ConsoleHelpers]::ShowWindow([ConsoleHelpers]::GetConsoleWindow(), 0)
   ```
   Uses Win32 API to hide the PowerShell console, presenting only the GUI.

2. **Shell.Application Compression**
   ```powershell
   $shell = New-Object -ComObject Shell.Application
   $zipFile.CopyHere($srcFolder.Items(), 20) # Flags: 4 (no UI) + 16 (yes to all)
   ```
   Chosen over `Compress-Archive` for reliability with nested folder structures.

3. **HTML Injection Strategy**
   - Locates anchor line: `DS.connection.startAssetGroup = startAssetGroup;`
   - Inserts SCORM event listeners immediately after
   - Adds `<script>` tag reference before `</body>`

4. **Regex-Based Validation**
   ```powershell
   $scoringActionPattern = '"scorings":\s*\[.*?\"type\":\s*\"action\".*?\]'
   ```
   Searches `data.js` for completion trigger patterns.

### Processing Workflow

#### Single Publish Flow
```
1. Test-ScormScoringCompliance (abort if no completion trigger)
   ↓
2. Copy scormAPI.min.js (if missing)
   ↓
3. Inject SCORM tracking code into index_lms.html
   ↓
4. Add <script> tag for scormAPI.min.js
   ↓
5. Rename story.html → index.html
   ↓
6. Open in Chrome for debugging
   ↓
7. Calculate and verify total score (user confirms)
   ↓
8. Compress folder contents to ZIP
```

#### Bulk Publish Flow
```
1. Iterate through all subfolders
   ↓
2. Process each folder (same as Single Publish steps 1-5)
   ↓
3. Optional: Open ALL modified files in Chrome tabs
   ↓
4. Generate bulk score verification report
   ↓
5. User confirms scores → ZIP all folders
```

### Code Injection Details

**Injected JavaScript** (29 lines):
```javascript
window.addEventListener("load", function () {
  setTimeout(function () {
    window.API.on('LMSInitialize', () => { sendScormData(); });
    window.API.on('LMSFinish', () => { sendScormData(); });
    // ... additional event listeners
  }, 1);
}, false);

function sendScormData() {
  try {
    console.log(JSON.stringify(window.API.cmi));
    window.postMessage(JSON.stringify(window.API.cmi), '*');
    window.parent.postMessage(JSON.stringify(window.API.cmi), '*');
  } catch (_) { }
}
```

**Purpose**: Intercepts SCORM API calls and logs CMI (Computer Managed Instruction) data to console for debugging.

## Installation

### Prerequisites
- Windows 10/11
- PowerShell 5.1+ (pre-installed on modern Windows)
- Google Chrome (required for debugging)
- `scormAPI.min.js` file in the same directory as the script

### Installation Steps

**Option 1: EXE Installer (Recommended)**
1. Run the `D-Forge-Setup.exe` file
2. The installer creates a desktop shortcut
3. Double-click the shortcut to launch D-Forge

**Option 2: Manual Setup**
1. Create a folder (e.g., `C:\D-Forge`)
2. Place `main.ps1` and `scormAPI.min.js` in the folder
3. Create a shortcut with target:
   ```
   powershell.exe -ExecutionPolicy Bypass -File "C:\D-Forge\main.ps1"
   ```
4. Set "Run as administrator" in shortcut properties (recommended)

### First-Time Configuration

**Chrome Path Verification**:
The script expects Chrome at:
```powershell
$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
```

If Chrome is installed elsewhere, edit line 35 of `main.ps1` to match your installation path.

## How to Use

### Basic Workflow

1. **Launch D-Forge**
   - Run from desktop shortcut or execute `main.ps1` directly

2. **Enter Target Path**
   - **Single operations**: Path to one course folder (e.g., `C:\Courses\Module1`)
   - **Bulk operations**: Path to parent folder containing multiple course folders

3. **Select Operation**
   - Use buttons 1-3 for processing and packaging
   - Use buttons 4-7 for quality assurance checks

4. **Monitor Progress**
   - Watch the Process Log for real-time status updates
   - ✅ Green = success
   - ❌ Red = error
   - ℹ️ Blue = informational

### Usage Scenarios

#### Scenario 1: Quick Packaging (Already Compliant)
```
1. Enter path: C:\Courses\Module1
2. Click "Zip Only"
3. Verify score when prompted
4. Find ZIP in C:\Courses\Module1\ZippedFiles\
```

#### Scenario 2: Fix Non-Compliant Course
```
1. Enter path: C:\Courses\Module2
2. Click "Single Publish"
3. Review Chrome console output (close tab when done)
4. Confirm score calculation
5. Receive packaged ZIP file
```

#### Scenario 3: Process Multiple Courses
```
1. Enter path: C:\Courses\ (parent folder)
2. Click "Bulk Publish"
3. All subfolders processed automatically
4. Choose whether to debug all in Chrome
5. Review bulk score report
6. Confirm to ZIP all
```

#### Scenario 4: Pre-Flight QA Check
```
1. Enter path: C:\Courses\ (parent folder)
2. Click "Check Completion (Bulk)" → verify all have triggers
3. Click "Check Scores (Bulk)" → review score report
4. Fix any issues in Articulate Storyline
5. Re-export and run Bulk Publish
```

### Expected Folder Structure

**Input (Articulate Storyline Export)**:
```
CourseFolder/
├── story.html
├── index_lms.html
└── html5/
    └── data/
        └── js/
            └── data.js
```

**Output (After Processing)**:
```
CourseFolder/
├── index.html (renamed from story.html)
├── index_lms.html (modified)
├── scormAPI.min.js (copied)
└── html5/
    └── data/
        └── js/
            └── data.js

ParentFolder/
└── ZippedFiles/
    └── CourseFolder.zip
```

## Limitations

### Technical Constraints

1. **Windows-Only**
   - Uses Windows Shell COM objects and Win32 APIs
   - Not portable to macOS or Linux

2. **Chrome Dependency**
   - Debugging feature requires Chrome browser
   - Hardcoded path may need manual adjustment

3. **Articulate Storyline Specific**
   - Assumes standard Articulate HTML5 export structure
   - May not work with other authoring tools (Captivate, Lectora, etc.)

4. **No Memory Between Sessions**
   - Does not track previously processed files
   - No batch history or undo functionality

5. **Limited Error Recovery**
   - Single point of failure: if `data.js` is malformed, script aborts
   - No automatic backup before modifications

### Functional Limitations

1. **Completion Trigger Detection**
   - Only checks for `"type":"action"` pattern
   - Cannot validate if trigger is correctly configured in Storyline

2. **Score Calculation**
   - Sums all `maxpoints` values found
   - Cannot differentiate between question banks or conditional scoring

3. **ZIP Compression**
   - Uses synchronous Shell.Application (slower for large files)
   - No progress bar during compression

4. **Browser Automation**
   - Opens tabs manually; user must close them
   - No automated console log capture

5. **Validation Gaps**
   - Does not verify SCORM manifest (`imsmanifest.xml`)
   - Cannot test actual LMS compatibility
   - No validation of SCORM 1.2 vs. 2004 standards

### Known Issues

- **Performance**: Bulk operations with 20+ courses can take 5-10 minutes
- **UI Freezing**: During compression, GUI may appear unresponsive (expected behavior)
- **Path Length**: Windows MAX_PATH (260 characters) limitation applies

## Best Practices

1. **Always run "Check Completion (Single/Bulk)" before packaging** to catch missing triggers early
2. **Verify scores match your Storyline quiz** before final ZIP confirmation
3. **Test one course with "Single Publish"** before running bulk operations
4. **Keep backups** of original Storyline exports before processing
5. **Use consistent folder naming** (no special characters) for cleaner logs

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Chrome not found" error | Edit line 35 to correct Chrome path |
| Script won't run | Right-click shortcut → "Run as administrator" |
| Completion check fails | Re-export from Storyline with a "Complete Course" trigger |
| Score seems wrong | Check if some slides use question banks (may be counted multiple times) |
| ZIP operation hangs | Wait 2-3 minutes; close other Shell windows if open |

## Version History

**Current Version**: 1.0 (January 2026)
- Initial release with 7-button interface
- Added bulk completion/score checking
- Console window hiding implementation

---

**Developed for**: SCORM content publishers using Articulate Storyline
