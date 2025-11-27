# D-Forge-GUI.ps1
# GUI-based automation script for publishing and preparing SCORM files.
# Adds check-completion and check-score features (single + bulk).
# Hides the console window when run.

# -------------------------------------------------------------------------
# Hide the PowerShell console window so only the GUI remains
# -------------------------------------------------------------------------
$consoleHelperCSharp = @"
using System;
using System.Runtime.InteropServices;

public static class ConsoleHelpers {
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();

    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

# Add the C# type (use -TypeDefinition for full type bodies)
Add-Type -TypeDefinition $consoleHelperCSharp -Language CSharp

# 0 = Hide window, 1 = Show
$hwnd = [ConsoleHelpers]::GetConsoleWindow()
if ($hwnd -ne [IntPtr]::Zero) {
    [ConsoleHelpers]::ShowWindow($hwnd, 0) | Out-Null
}


# -------------------------------------------------------------------------
# --- Configuration & Utility Functions ---
# -------------------------------------------------------------------------
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = Get-Location }

$sourceFilePath = Join-Path -Path $scriptDir -ChildPath "scormAPI.min.js"

# Path to Chrome -- please adjust if required
$ChromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"

# Global controls
$global:RichTextBoxLog = $null
$global:TextBoxPath = $null
$global:Window = $null

function Append-ColoredText {
    param([string]$Message, [string]$Color)
    if ($global:RichTextBoxLog) {
        $run = New-Object System.Windows.Documents.Run
        $run.Text = $Message + "`n"
        try {
            $run.Foreground = [System.Windows.Media.Brushes]::$Color
        } catch {
            $run.Foreground = [System.Windows.Media.Brushes]::Black
        }
        $currentParagraph = $global:RichTextBoxLog.Document.Blocks.LastBlock
        if (-not $currentParagraph) {
            $currentParagraph = New-Object System.Windows.Documents.Paragraph
            $global:RichTextBoxLog.Document.Blocks.Add($currentParagraph)
        }
        $currentParagraph.Inlines.Add($run)
        $global:RichTextBoxLog.ScrollToEnd()
    } else {
        Write-Host $Message
    }
}
function Show-Error { param([string]$Message) ; Append-ColoredText -Message "❗ ERROR: $Message" -Color "Red" }
function Write-Step { param([string]$Message) ; Append-ColoredText -Message "✅ $Message" -Color "Green" }
function Write-Log { param([string]$Message, [string]$Color = "Black") ; Append-ColoredText -Message $Message -Color $Color }

# -------------------------------------------------------------------------
# --- Score & Completion Helpers ---
# -------------------------------------------------------------------------
function Get-ScormTotalScore {
    param([string]$targetFolder)
    $dataPath = Join-Path -Path $targetFolder -ChildPath "html5\data\js"
    $dataJSPath = Join-Path -Path $dataPath -ChildPath "data.js"

    if (-not (Test-Path $dataJSPath)) {
        Show-Error "Cannot calculate score. data.js not found at $dataJSPath"
        return $null
    }
    try {
        $content = Get-Content -Path $dataJSPath -Raw -Encoding UTF8 -ErrorAction Stop
        $matches = [Regex]::Matches($content, '"maxpoints":\s*(\d+)')
        $totalScore = 0
        foreach ($match in $matches) {
            if ($match.Groups[1].Success) {
                $totalScore += [int]$match.Groups[1].Value
            }
        }
        return $totalScore
    } catch {
        Show-Error "Error reading scores from data.js: $($_.Exception.Message)"
        return $null
    }
}

function Test-ScormScoringCompliance {
    param([string]$targetFolder)
    Write-Log "`nChecking internal scoring compliance..." -Color "DarkGray"
    $dataPath = Join-Path -Path $targetFolder -ChildPath "html5\data\js"
    $dataJSPath = Join-Path -Path $dataPath -ChildPath "data.js"
    if (-Not (Test-Path $dataJSPath -PathType Leaf)) {
        Show-Error "data.js not found at '$dataJSPath'. The course must contain the folder structure html5\data\js\data.js."
        return $false
    }

    try {
        $dataJSContent = Get-Content -Path $dataJSPath -Raw -Encoding UTF8 -ErrorAction Stop
    } catch {
        Show-Error "Error reading $dataJSPath. Message: $($_.Exception.Message)"
        return $false
    }

    $scoringActionPattern = '"scorings":\s*\[.*?\"type\":\s*\"action\".*?\]'
    if ($dataJSContent -match $scoringActionPattern) {
        Write-Step "Scoring compliance check passed: Found 'type':'action' (Completed Trigger) in data.js."
        return $true
    } else {
        Show-Error "Scoring compliance check FAILED. The required completed status trigger ('type':'action') is missing."
        return $false
    }
}

# -------------------------------------------------------------------------
# --- Reused: Folder compliance, Compress, Process-ScormFolder (trimmed a bit but same logic)
# -------------------------------------------------------------------------
function Compress-FolderContent {
    param([string]$PathToZip)
    try {
        if (-Not (Test-Path $PathToZip -PathType Container)) {
            Show-Error "Folder to zip not found: $PathToZip"
            return $false
        }
        $parentDir = Split-Path -Path $PathToZip -Parent
        $folderName = Split-Path -Path $PathToZip -Leaf
        $zipDestinationFolder = Join-Path -Path $parentDir -ChildPath "ZippedFiles"
        if (-Not (Test-Path $zipDestinationFolder)) { New-Item -ItemType Directory -Path $zipDestinationFolder | Out-Null }
        $zipFileName = "$($folderName).zip"
        $zipFilePath = Join-Path -Path $zipDestinationFolder -ChildPath $zipFileName
        $itemsToZip = Get-ChildItem -Path $PathToZip -Force
        if ($itemsToZip.Count -eq 0) {
            Show-Error "No content found inside '$folderName' to zip."
            return $false
        }
        Compress-Archive -Path $itemsToZip.FullName -DestinationPath $zipFilePath -Force
        Write-Step "Contents inside '$folderName' zipped successfully: $zipFilePath"
        return $true
    } catch {
        Show-Error "Zipping failed for $PathToZip. $($_.Exception.Message)"
        return $false
    }
}

function Test-ScormFolderCompliance {
    param([string]$targetFolder)
    Write-Log "`nChecking folder compliance for $targetFolder..."
    $isCompliant = $true
    $scormAPIFilePath = Join-Path -Path $targetFolder -ChildPath 'scormAPI.min.js'
    if (-Not (Test-Path $scormAPIFilePath)) {
        Write-Log " - ❌ scormAPI.min.js is missing." -Color "Red"
        $isCompliant = $false
    }
    $indexFilePath = Join-Path -Path $targetFolder -ChildPath 'index.html'
    if (-Not (Test-Path $indexFilePath)) {
        Write-Log " - ❌ index.html (renamed story.html) is missing." -Color "Red"
        $isCompliant = $false
    }
    $htmlFilePath = Join-Path -Path $targetFolder -ChildPath 'index_lms.html'
    if (Test-Path $htmlFilePath) {
        try {
            $scormCodePatternEscaped = [Regex]::Escape("window.API.on('LMSInitialize'")
            if (-Not (Get-Content $htmlFilePath | Select-String -Pattern $scormCodePatternEscaped -Quiet)) {
                Write-Log " - ❌ SCORM API code not found in index_lms.html." -Color "Red"
                $isCompliant = $false
            }
            if (-Not (Get-Content $htmlFilePath | Select-String -Pattern "scormAPI.min.js" -Quiet)) {
                Write-Log " - ❌ <script> tag reference to scormAPI.min.js is missing." -Color "Red"
                $isCompliant = $false
            }
        } catch {
            Show-Error "Error reading $($htmlFilePath): $($_.Exception.Message)"
            $isCompliant = $false
        }
    } else {
        Write-Log " - ❌ index_lms.html file not found." -Color "Red"
        $isCompliant = $false
    }

    if ($isCompliant) {
        Write-Log "Folder is **compliant**." -Color "Blue"
    } else {
        Write-Log "Folder is **NOT compliant**. Fixes needed." -Color "Red"
    }
    return $isCompliant
}

function Process-ScormFolder {
    param([string]$targetFolder)
    Write-Log "`n--- Processing Folder: $targetFolder ---" -Color "Blue"

    if (-not (Test-ScormScoringCompliance -targetFolder $targetFolder)) {
        Show-Error "SCORING COMPLIANCE FAILED. Aborting SCORM preparation for $targetFolder."
        return $false
    }

    $htmlFilePath = Join-Path -Path $targetFolder -ChildPath 'index_lms.html'
    $scormAPIFilePath = Join-Path -Path $targetFolder -ChildPath 'scormAPI.min.js'
    $storyFilePath = Join-Path -Path $targetFolder -ChildPath 'story.html'
    $indexFilePath = Join-Path -Path $targetFolder -ChildPath 'index.html'
    $htmlModified = $false

    if (-Not (Test-Path $htmlFilePath)) {
        Show-Error "index_lms.html file not found at $htmlFilePath. Skipping."
        return $false
    }

    # Step 1: Copy scormAPI.min.js if missing
    if (-Not (Test-Path $scormAPIFilePath)) {
        if (Test-Path $sourceFilePath) {
            Copy-Item -Path $sourceFilePath -Destination $targetFolder -Force -ErrorAction Stop
            Write-Step "Copied scormAPI.min.js."
        } else {
            Show-Error "Source file scormAPI.min.js not found at $sourceFilePath. Cannot complete script."
            return $false
        }
    } else {
        Write-Log "scormAPI.min.js already exists. Skipping copy."
    }

    # Step 2 & 3: Modify / Insert code and script tag if missing (same logic as before)
    try {
        $lines = Get-Content $htmlFilePath -ErrorAction Stop
        $scormCodePattern = [Regex]::Escape("window.API.on('LMSInitialize'")
        $scormCodeExists = $lines | Select-String -Pattern $scormCodePattern -Quiet
        $scriptTagPattern = "scormAPI.min.js"
        $scriptTagExists = $lines | Select-String -Pattern $scriptTagPattern -Quiet

        if (-Not $scormCodeExists -or -Not $scriptTagExists) {
            Write-Log "HTML patching needed. Applying missing patches..."
            if (-Not $scormCodeExists) {
                $targetLinePattern = "^[ \t]*DS\.connection\.startAssetGroup = startAssetGroup;[ \t]*$"
                $startAssetGroupLineIndex = -1
                for ($i = 0; $i -lt $lines.Length; $i++) {
                    if ($lines[$i].Trim() -match "DS.connection.startAssetGroup = startAssetGroup;") { $startAssetGroupLineIndex = $i; break }
                }
                if ($startAssetGroupLineIndex -ne -1) {
                    $insertionPointIndex = $startAssetGroupLineIndex + 1
                    if (($startAssetGroupLineIndex + 2) -lt $lines.Length) {
                        $newLines = $lines[0..$startAssetGroupLineIndex]
                        $endIndex = $lines.Length - 1
                        if (($startAssetGroupLineIndex + 2) -le $endIndex) {
                            $newLines += $lines[($startAssetGroupLineIndex + 2)..$endIndex]
                        }
                        $lines = $newLines
                        $insertionPointIndex = $startAssetGroupLineIndex + 1
                    }

                    $scormCodeToInsert = @"
    window.addEventListener("load", function () {
        setTimeout(function () {
          window.API.on('LMSInitialize', () => { sendScormData(); });
          window.API.on('LMSFinish', () => { sendScormData(); });
          window.API.on('LMSGetValue', (val) => { sendScormData(); });
          window.API.on('LMSSetValue', (ele, val) => { sendScormData(); });
          window.API.on('LMSCommit', (val) => { sendScormData(); });
          window.API.on('LMSGetLastError', () => { sendScormData(); });
          window.API.on('LMSGetErrorString', (val) => { sendScormData(); });
          window.API.on('LMSGetDiagnostic', (val) => { sendScormData(); });
        }, 1);
      }, false);

    })();

function sendScormData() {
        try {
          console.log(JSON.stringify(window.API.cmi))
          window.postMessage(JSON.stringify(window.API.cmi), '*')
          window.parent.postMessage(JSON.stringify(window.API.cmi), '*')
          LMSSetValue.postMessage(JSON.stringify(window.API.cmi))
        } catch (_) { }
      }
"@
                    $linesToInsert = $scormCodeToInsert.Split([Environment]::NewLine)
                    $newContent = @($lines[0..$startAssetGroupLineIndex] ; $linesToInsert)
                    $endOfArray = $lines.Length - 1
                    if ($insertionPointIndex -le $endOfArray) { $newContent += $lines[($insertionPointIndex)..$endOfArray] }
                    $lines = $newContent
                    Write-Step "Inserted the Scorm_fetch_Code."
                    $htmlModified = $true
                } else {
                    Show-Error "'DS.connection.startAssetGroup = startAssetGroup;' not found in $htmlFilePath. Cannot insert code."
                    return $false
                }
            } else {
                Write-Log "SCORM fetch code already present. Skipping injection."
            }

            if (-Not $scriptTagExists) {
                $bodyTagIndex = -1
                for ($i = $lines.Length - 1; $i -ge 0; $i--) {
                    if ($lines[$i] -match "</body>") { $bodyTagIndex = $i; break }
                }
                if ($bodyTagIndex -ne -1) {
                    $scriptTagToInsert = "<script src='scormAPI.min.js' type='text/javascript'></script>"
                    $newContent = @($lines[0..($bodyTagIndex - 1)] ; $scriptTagToInsert)
                    $endOfArray = $lines.Length - 1
                    if ($bodyTagIndex -le $endOfArray) { $newContent += $lines[$bodyTagIndex..$endOfArray] }
                    $lines = $newContent
                    Write-Step "Added <script> tag."
                    $htmlModified = $true
                } else {
                    Show-Error "Could not find </body> tag in $htmlFilePath. Cannot insert script tag."
                }
            } else {
                Write-Log "<script> tag already present. Skipping insertion."
            }

            if ($htmlModified) {
                $lines | Set-Content $htmlFilePath -ErrorAction Stop
                Write-Step "Successfully modified $htmlFilePath."
            }
        } else {
            Write-Log "HTML file already patched (Code and Script Tag). Skipping modification."
        }
    } catch {
        Show-Error "Error modifying $($htmlFilePath). $($_.Exception.Message)"
        return $false
    }

    if (Test-Path $storyFilePath) {
        if (-Not (Test-Path $indexFilePath)) {
            Rename-Item -Path $storyFilePath -NewName "index.html" -Force -ErrorAction Stop
            Write-Step "Renamed story.html to index.html."
        } else {
            Write-Log "story.html exists but index.html already exists. Skipping rename."
        }
    } elseif (Test-Path $indexFilePath) {
        Write-Log "index.html already exists. Skipping rename."
    } else {
        Write-Log "Neither story.html nor index.html found. Skipping rename."
    }

    return $true
}

# -------------------------------------------------------------------------
# --- GUI Helpers ---
# -------------------------------------------------------------------------
function Get-TargetFolderFromUI {
    if ($null -ne $global:TextBoxPath) {
        $path = $global:TextBoxPath.Text.Trim()
        if (-not [string]::IsNullOrEmpty($path) -and (Test-Path $path -PathType Container)) {
            return $path
        }
    }
    Show-Error "Path input is invalid or empty. Please enter a valid folder path in the box above."
    return $null
}

function Prompt-YesNoDialog {
    param([string]$Message, [string]$Title = "Confirmation")
    Add-Type -AssemblyName PresentationCore, PresentationFramework
    $Result = [System.Windows.MessageBox]::Show($Message, $Title, [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    return $Result -eq [System.Windows.MessageBoxResult]::Yes
}

function Open-ChromeDebug {
    param([string]$htmlFilePath)
    if (Test-Path $ChromePath) {
        Write-Log "Opening $htmlFilePath in Chrome..."
        Start-Process -FilePath $ChromePath -ArgumentList "`"$htmlFilePath`""
        if (Prompt-YesNoDialog "Please check the Chrome developer console for SCORM log output, then close the tab.`n`nHave you closed the Chrome tab and are ready to continue?" "Debug Complete?") {
            return $true
        } else {
            Show-Error "Aborted by user."
            return $false
        }
    } else {
        Show-Error "Chrome not found at $ChromePath. Skipping debug."
        return $true
    }
}

# -------------------------------------------------------------------------
# --- Event Handlers: existing operations (Zip Only, Bulk Publish, Single Publish)
# -------------------------------------------------------------------------
function Handle-ZipOnly {
    $folderPath = Get-TargetFolderFromUI
    if (-not $folderPath) { return }
    Write-Log "`n--- Starting Zip Only Operation ---" -Color "Blue"
    if (-not (Test-ScormScoringCompliance -targetFolder $folderPath)) {
        Show-Error "Folder is scoring-non-compliant. Cannot proceed to zip."
        Write-Log "--- Zip Only Operation Finished ---" -Color "Blue"
        return
    }
    $isCompliant = Test-ScormFolderCompliance -targetFolder $folderPath
    if ($isCompliant) {
        if (Verify-ScormScore -targetFolder $folderPath) {
            Compress-FolderContent -PathToZip $folderPath
        }
    } else {
        Show-Error "Folder is NOT compliant. You must use the Single Publish option to fix it."
        if (Prompt-YesNoDialog "The folder is NOT compliant. Do you want to run Single Publish to fix and zip it now?" "Fix Folder?") {
            if (Process-ScormFolder -targetFolder $folderPath) {
                $htmlFilePath = Join-Path -Path $folderPath -ChildPath 'index_lms.html'
                if (Open-ChromeDebug -htmlFilePath $htmlFilePath) {
                    if (Verify-ScormScore -targetFolder $folderPath) {
                        Compress-FolderContent -PathToZip $folderPath
                    }
                }
            }
        }
    }
    Write-Log "--- Zip Only Operation Finished ---" -Color "Blue"
}

function Handle-BulkPublish {
    $parentFolder = Get-TargetFolderFromUI
    if (-not $parentFolder) { return }
    Write-Log "`n--- Starting Bulk Publish Operation on $parentFolder ---" -Color "Blue"
    $subFolders = Get-ChildItem -Path $parentFolder -Directory -ErrorAction SilentlyContinue
    if (-Not $subFolders) {
        Show-Error "No subfolders found in $parentFolder. Exiting Bulk Publish mode."
        return
    }
    Write-Log "Found $($subFolders.Count) subfolders. Starting processing..."
    $successfullyProcessedFolders = @()
    foreach ($folder in $subFolders) {
        $folderPath = $folder.FullName
        $publishSuccess = Process-ScormFolder -targetFolder $folderPath
        if ($publishSuccess) { $successfullyProcessedFolders += $folderPath }
    }

    if ($successfullyProcessedFolders.Count -gt 0) {
        Write-Log "`n--- Debugging Phase ---" -Color "Blue"
        if (Prompt-YesNoDialog "Do you want to open ALL $($successfullyProcessedFolders.Count) modified files in Chrome for debugging?" "Bulk Debug?") {
            if (-Not (Test-Path $ChromePath)) {
                Show-Error "Chrome executable not found at $ChromePath. Skipping debug launch."
            } else {
                Write-Log "Opening all modified index_lms.html files in Chrome. PLEASE CLOSE ALL TABS AFTER CHECKING."
                foreach ($folderPath in $successfullyProcessedFolders) {
                    $htmlFilePath = Join-Path -Path $folderPath -ChildPath 'index_lms.html'
                    Start-Process -FilePath $ChromePath -ArgumentList "`"$htmlFilePath`"" -ErrorAction SilentlyContinue
                    Write-Log "Opened: $folderPath"
                }
                if (-not (Prompt-YesNoDialog "`nHave you closed ALL the Chrome tabs? Press 'Yes' to continue to ZIP all files" "Debug Complete?")) {
                    Show-Error "Bulk Zipping aborted by user."
                    Write-Log "--- Bulk Publish Operation Finished (Aborted) ---" -Color "Blue"
                    return
                }
            }
        }

        # Bulk verification step
        Write-Log "`n--- Bulk Verification & Zipping Phase ---" -Color "Blue"
        Write-Log "Calculating scores for all folders..." -Color "DarkGray"
        $verificationReport = "Please verify the following Total Scores (Sum of 'maxpoints'):`n`n"
        $foldersReadyToZip = @()
        foreach ($folderPath in $successfullyProcessedFolders) {
            $folderName = Split-Path -Path $folderPath -Leaf
            $score = Get-ScormTotalScore -targetFolder $folderPath
            if ($null -ne $score) {
                $verificationReport += "• $folderName : $score`n"
                $foldersReadyToZip += $folderPath
            } else {
                $verificationReport += "• $folderName : ERROR (Could not read score)`n"
            }
        }
        $verificationReport += "`nDo you want to proceed with Zipping these folders?`n(Click Yes to Zip All, No to Cancel)"
        if (Prompt-YesNoDialog $verificationReport "Bulk Score Verification") {
            foreach ($folderPath in $foldersReadyToZip) {
                Compress-FolderContent -PathToZip $folderPath
            }
        } else {
            Show-Error "Bulk Score Verification rejected by user. Zipping aborted."
        }
    } else {
        Show-Error "No folders were successfully processed to zip."
    }
    Write-Log "--- Bulk Publish Operation Finished ---" -Color "Blue"
}

function Handle-SinglePublish {
    $folderPath = Get-TargetFolderFromUI
    if (-not $folderPath) { return }
    Write-Log "`n--- Starting Single Publish Operation on $folderPath ---" -Color "Blue"
    if (Process-ScormFolder -targetFolder $folderPath) {
        Write-Log "`n--- Debugging & Zipping Phase ---" -Color "Blue"
        $htmlFilePath = Join-Path -Path $folderPath -ChildPath 'index_lms.html'
        if (Open-ChromeDebug -htmlFilePath $htmlFilePath) {
            if (Verify-ScormScore -targetFolder $folderPath) {
                Compress-FolderContent -PathToZip $folderPath
            }
        }
    } else {
        Show-Error "Single publish failed for $folderPath. Check log messages."
    }
    Write-Log "--- Single Publish Operation Finished ---" -Color "Blue"
}

# -------------------------------------------------------------------------
# --- NEW: Check functions (Single & Bulk) and UI handlers for them
# -------------------------------------------------------------------------
function Handle-CheckCompletionSingle {
    $folderPath = Get-TargetFolderFromUI
    if (-not $folderPath) { return }
    Write-Log "`n--- Check Completion (Single) ---" -Color "Blue"
    $passed = Test-ScormScoringCompliance -targetFolder $folderPath
    $folderName = Split-Path -Path $folderPath -Leaf
    if ($passed) {
        Write-Step "Completion trigger exists for $folderName."
        Add-Type -AssemblyName PresentationCore, PresentationFramework
        [System.Windows.MessageBox]::Show("Folder: $folderName`n`nCompletion trigger (type:'action') FOUND.`n`nOK.", "Completion Check: $folderName", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    } else {
        Show-Error "Completion trigger MISSING for $folderName."
        Add-Type -AssemblyName PresentationCore, PresentationFramework
        [System.Windows.MessageBox]::Show("Folder: $folderName`n`nCompletion trigger (type:'action') MISSING.`nPlease re-publish the course with a Completed trigger.", "Completion Check: $folderName", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null
    }
}

function Handle-CheckScoreSingle {
    $folderPath = Get-TargetFolderFromUI
    if (-not $folderPath) { return }
    Write-Log "`n--- Check Score (Single) ---" -Color "Blue"
    $score = Get-ScormTotalScore -targetFolder $folderPath
    $folderName = Split-Path -Path $folderPath -Leaf
    if ($null -eq $score) {
        Show-Error "Could not read score for $folderName."
        Add-Type -AssemblyName PresentationCore, PresentationFramework
        [System.Windows.MessageBox]::Show("Folder: $folderName`n`nERROR: Could not read 'maxpoints' from data.js.", "Score Check: $folderName", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    } else {
        Write-Step "Calculated total score for $folderName = $score"
        Add-Type -AssemblyName PresentationCore, PresentationFramework
        [System.Windows.MessageBox]::Show("Folder: $folderName`n`nCalculated Total Score (Sum of 'maxpoints'): $score", "Score Check: $folderName", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    }
}

function Handle-CheckCompletionBulk {
    $parentFolder = Get-TargetFolderFromUI
    if (-not $parentFolder) { return }
    Write-Log "`n--- Check Completion (Bulk) ---" -Color "Blue"
    $subFolders = Get-ChildItem -Path $parentFolder -Directory -ErrorAction SilentlyContinue
    if (-Not $subFolders) {
        Show-Error "No subfolders found in $parentFolder."
        return
    }

    $report = "Completion Check Results:`n`n"
    foreach ($folder in $subFolders) {
        $folderPath = $folder.FullName
        $folderName = $folder.Name
        $passed = Test-ScormScoringCompliance -targetFolder $folderPath
        if ($passed) {
            $report += "• $folderName : OK (completion trigger found)`n"
        } else {
            $report += "• $folderName : MISSING completion trigger`n"
        }
    }

    Add-Type -AssemblyName PresentationCore, PresentationFramework
    [System.Windows.MessageBox]::Show($report, "Bulk Completion Check", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    Write-Log "--- Bulk Completion Check Finished ---" -Color "Blue"
}

function Handle-CheckScoresBulk {
    $parentFolder = Get-TargetFolderFromUI
    if (-not $parentFolder) { return }
    Write-Log "`n--- Check Scores (Bulk) ---" -Color "Blue"
    $subFolders = Get-ChildItem -Path $parentFolder -Directory -ErrorAction SilentlyContinue
    if (-Not $subFolders) {
        Show-Error "No subfolders found in $parentFolder."
        return
    }

    $report = "Scores Report (Sum of 'maxpoints'):`n`n"
    foreach ($folder in $subFolders) {
        $folderPath = $folder.FullName
        $folderName = $folder.Name
        $score = Get-ScormTotalScore -targetFolder $folderPath
        if ($null -ne $score) {
            $report += "• $folderName : $score`n"
        } else {
            $report += "• $folderName : ERROR (data.js missing or unreadable)`n"
        }
    }

    Add-Type -AssemblyName PresentationCore, PresentationFramework
    [System.Windows.MessageBox]::Show($report, "Bulk Scores Report", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
    Write-Log "--- Bulk Scores Check Finished ---" -Color "Blue"
}

# Reuse Verify-ScormScore (dialog used earlier)
function Verify-ScormScore {
    param([string]$targetFolder)
    Write-Log "Calculating total score for verification..." -Color "DarkGray"
    $score = Get-ScormTotalScore -targetFolder $targetFolder
    if ($null -eq $score) { return $false }
    $folderName = Split-Path -Path $targetFolder -Leaf
    $msg = "Folder: $folderName`n`nCalculated Total Score (Sum of 'maxpoints'): $score`n`nIs this score CORRECT?`n`n(Click Yes to Zip, No to Abort)"
    Add-Type -AssemblyName PresentationCore, PresentationFramework
    $Result = [System.Windows.MessageBox]::Show($msg, "Verify Score: $folderName", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($Result -eq [System.Windows.MessageBoxResult]::Yes) { Write-Step "Score verified as correct ($score)."; return $true } else { Show-Error "Score verification rejected by user (Score: $score). Aborting Zip for this folder."; return $false }
}

# -------------------------------------------------------------------------
# --- GUI (WPF/XAML) - 7 Buttons + Path + Log ---
# -------------------------------------------------------------------------
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="D-Forge" Width="1000" Height="650" WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
            <Label Content="Target Folder Path:" FontWeight="Bold" Width="150" VerticalAlignment="Center"/>
            <TextBox Name="TextBoxPath" Width="700" VerticalAlignment="Center" Padding="3" Text=""/>
        </StackPanel>

        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8" HorizontalAlignment="Center">
            <Button Name="ButtonZipOnly" Content="1. Zip Only (Compliance Check)" Margin="5" Padding="10" FontWeight="Bold" Width="220" Background="#6CB4EE" Foreground="White"/>
            <Button Name="ButtonBulkPublish" Content="2. Bulk Publish (Fix, Debug, Zip All)" Margin="5" Padding="10" FontWeight="Bold" Width="220" Background="#FFA500" Foreground="Black"/>
            <Button Name="ButtonSinglePublish" Content="3. Single Publish (Fix, Debug, Zip)" Margin="5" Padding="10" FontWeight="Bold" Width="220" Background="#4CAF50" Foreground="White"/>
        </StackPanel>

        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,0,0,8" HorizontalAlignment="Center">
            <Button Name="ButtonCheckCompletionSingle" Content="4. Check Completion (Single)" Margin="5" Padding="10" FontWeight="SemiBold" Width="220" Background="#8E44AD" Foreground="White"/>
            <Button Name="ButtonCheckScoreSingle" Content="5. Check Score (Single)" Margin="5" Padding="10" FontWeight="SemiBold" Width="220" Background="#2E86C1" Foreground="White"/>
            <Button Name="ButtonCheckCompletionBulk" Content="6. Check Completion (Bulk)" Margin="5" Padding="10" FontWeight="SemiBold" Width="220" Background="#F39C12" Foreground="Black"/>
            <Button Name="ButtonCheckScoresBulk" Content="7. Check Scores (Bulk)" Margin="5" Padding="10" FontWeight="SemiBold" Width="220" Background="#27AE60" Foreground="White"/>
        </StackPanel>

        <GroupBox Grid.Row="3" Header="Process Log" Margin="0,5,0,5">
            <RichTextBox Name="RichTextBoxLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto" Background="#F0F0F0" FontFamily="Consolas">
                <FlowDocument>
                    <Paragraph />
                </FlowDocument>
            </RichTextBox>
        </GroupBox>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$global:Window = [Windows.Markup.XamlReader]::Load($reader)

$global:TextBoxPath = $global:Window.FindName("TextBoxPath")
$global:RichTextBoxLog = $global:Window.FindName("RichTextBoxLog")

$ButtonZipOnly = $global:Window.FindName("ButtonZipOnly")
$ButtonBulkPublish = $global:Window.FindName("ButtonBulkPublish")
$ButtonSinglePublish = $global:Window.FindName("ButtonSinglePublish")
$ButtonCheckCompletionSingle = $global:Window.FindName("ButtonCheckCompletionSingle")
$ButtonCheckScoreSingle = $global:Window.FindName("ButtonCheckScoreSingle")
$ButtonCheckCompletionBulk = $global:Window.FindName("ButtonCheckCompletionBulk")
$ButtonCheckScoresBulk = $global:Window.FindName("ButtonCheckScoresBulk")

# Wire events
$ButtonZipOnly.Add_Click({ Handle-ZipOnly })
$ButtonBulkPublish.Add_Click({ Handle-BulkPublish })
$ButtonSinglePublish.Add_Click({ Handle-SinglePublish })

$ButtonCheckCompletionSingle.Add_Click({ Handle-CheckCompletionSingle })
$ButtonCheckScoreSingle.Add_Click({ Handle-CheckScoreSingle })
$ButtonCheckCompletionBulk.Add_Click({ Handle-CheckCompletionBulk })
$ButtonCheckScoresBulk.Add_Click({ Handle-CheckScoresBulk })

# Show window (blocking until closed)
$global:Window.ShowDialog() | Out-Null

# When the window closes, ensure process exits
Stop-Process -Id $PID -Force
