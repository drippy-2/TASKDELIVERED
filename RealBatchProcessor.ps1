# AutoCAD Batch DWG Processor - Real Version
# Processes multiple DWG files using accoreconsole.exe
# Features: Logging, .txt existence check, .bak cleanup, timeout handling

param(
    [string]$DWGFolder,
    [string]$ScriptFile = "C:\Scripts\process.scr",
    [string]$AutoCADPath = "C:\Program Files\Autodesk\AutoCAD 2024\accoreconsole.exe",
    [int]$TimeoutSeconds = 25
)

# Prompt for folder if not provided
if (-not $DWGFolder) {
    Add-Type -AssemblyName System.Windows.Forms
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select folder containing DWG files to batch process"
    $folderBrowser.RootFolder = "MyComputer"
    $result = $folderBrowser.ShowDialog()
    if ($result -eq "OK") { $DWGFolder = $folderBrowser.SelectedPath } 
    else { Write-Host "No folder selected. Exiting." -ForegroundColor Red; exit 0 }
}

# Validate paths
if (-not (Test-Path $DWGFolder)) { Write-Error "DWG folder not found: $DWGFolder"; exit 1 }
if (-not (Test-Path $ScriptFile)) { Write-Error "Script file not found: $ScriptFile"; exit 1 }
if (-not (Test-Path $AutoCADPath)) { Write-Error "AutoCAD executable not found: $AutoCADPath"; exit 1 }

# Log file
$LogFile = Join-Path $DWGFolder "Process_Log.txt"

# Get DWG files
$dwgFiles = Get-ChildItem -Path $DWGFolder -Filter "*.dwg"
$totalFiles = $dwgFiles.Count
if ($totalFiles -eq 0) { Write-Host "No DWG files found in $DWGFolder" -ForegroundColor Yellow; exit 0 }

# Initialize counters
$processedCount = 0
$errorCount = 0
$txtFoundCount = 0

# Log header
$sessionStartTime = Get-Date
"============================================" | Out-File -FilePath $LogFile
"           PROCESS LOG                       " | Out-File -FilePath $LogFile -Append
"============================================" | Out-File -FilePath $LogFile -Append
"Machine Name: $env:COMPUTERNAME" | Out-File -FilePath $LogFile -Append
"User Name:    $env:USERNAME" | Out-File -FilePath $LogFile -Append
"DWG Folder:   $DWGFolder" | Out-File -FilePath $LogFile -Append
"Script File:  $ScriptFile" | Out-File -FilePath $LogFile -Append
"Session Start: $($sessionStartTime.ToString('yyyy-MM-dd HH:mm:ss'))" | Out-File -FilePath $LogFile -Append
"Files to Process: $totalFiles" | Out-File -FilePath $LogFile -Append
"============================================" | Out-File -FilePath $LogFile -Append
"" | Out-File -FilePath $LogFile -Append

# Track previous DWG for .bak cleanup
$previousDwgFile = $null
$timeoutMs = $TimeoutSeconds * 1000

# Process each DWG
foreach ($dwg in $dwgFiles) {
    $fileName = $dwg.Name
    $filePath = $dwg.FullName
    $fileBaseName = $dwg.BaseName
    $fileDirectory = $dwg.DirectoryName
    $txtFilePath = Join-Path $fileDirectory "$fileBaseName.txt"
    $bakFilePath = Join-Path $fileDirectory "$fileBaseName.bak"

    # Delete previous .bak
    if ($previousDwgFile) {
        $prevBase = [System.IO.Path]::GetFileNameWithoutExtension($previousDwgFile)
        $prevDir = [System.IO.Path]::GetDirectoryName($previousDwgFile)
        $prevBak = Join-Path $prevDir "$prevBase.bak"
        if (Test-Path $prevBak) { Remove-Item -Path $prevBak -Force }
    }

    # Log start
    $startTime = Get-Date
    Add-Content -Path $LogFile -Value "------------------------------"
    Add-Content -Path $LogFile -Value "Processing File: $fileName"
    Add-Content -Path $LogFile -Value "Start: $($startTime.ToString('HH:mm:ss'))"
    Write-Host "Processing: $fileName" -ForegroundColor Yellow

    # Launch AutoCAD
    try {
        $acadProcess = Start-Process -FilePath $AutoCADPath `
            -ArgumentList "/i `"$filePath`" /s `"$ScriptFile`"" `
            -PassThru -NoNewWindow

        # Wait with timeout
        if ($acadProcess.WaitForExit($timeoutMs)) {
            $exitCode = $acadProcess.ExitCode
        }
        else {
            # Timeout reached – force close AutoCAD
            $acadProcess.Kill()
            $exitCode = -1
        }

        # End time & duration
        $endTime = Get-Date
        $duration = $endTime - $startTime

        # Check for .txt completion file
        $txtExists = Test-Path $txtFilePath
        $txtStatus = if ($txtExists) { "Yes" } else { "No" }
        if ($txtExists) { $txtFoundCount++ }

        # Log details EXACTLY as required
        Add-Content -Path $LogFile -Value "End: $($endTime.ToString('HH:mm:ss'))"
        Add-Content -Path $LogFile -Value "Duration: $([int]$duration.TotalSeconds) sec"
        Add-Content -Path $LogFile -Value ".txt exists? $txtStatus"
        Add-Content -Path $LogFile -Value "Exit Code: $exitCode"

        # SUCCESS or ERROR status
        if ($exitCode -eq 0 -and $txtExists) {
            Add-Content -Path $LogFile -Value "Status: SUCCESS"
            Write-Host "Status: SUCCESS" -ForegroundColor Green
            $processedCount++
        }
        else {
            Add-Content -Path $LogFile -Value "Status: ERROR"
            Write-Host "Status: ERROR" -ForegroundColor Red
            $errorCount++
        }
    }
    catch {
        $endTime = Get-Date
        Add-Content -Path $LogFile -Value "End: $($endTime.ToString('HH:mm:ss'))"
        Add-Content -Path $LogFile -Value "Duration: N/A"
        Add-Content -Path $LogFile -Value ".txt exists? N/A"
        Add-Content -Path $LogFile -Value "Exit Code: N/A"
        Add-Content -Path $LogFile -Value "Status: ERROR (Exception)"
        Write-Host "Status: ERROR (Exception)" -ForegroundColor Red
        $errorCount++
    }
    $previousDwgFile = $filePath
}

# Final .bak cleanup
if ($previousDwgFile) {
    $prevBase = [System.IO.Path]::GetFileNameWithoutExtension($previousDwgFile)
    $prevDir = [System.IO.Path]::GetDirectoryName($previousDwgFile)
    $prevBak = Join-Path $prevDir "$prevBase.bak"
    if (Test-Path $prevBak) { Remove-Item -Path $prevBak -Force }
}

# Summary
$sessionEndTime = Get-Date
$totalDuration = $sessionEndTime - $sessionStartTime

Add-Content -Path $LogFile -Value "============================================"
Add-Content -Path $LogFile -Value "           PROCESSING SUMMARY               "
Add-Content -Path $LogFile -Value "============================================"
Add-Content -Path $LogFile -Value "Session End: $($sessionEndTime.ToString('yyyy-MM-dd HH:mm:ss'))"
Add-Content -Path $LogFile -Value "Total Duration: $($totalDuration.TotalMinutes.ToString('F2')) minutes"
Add-Content -Path $LogFile -Value "Files Processed: $processedCount"
Add-Content -Path $LogFile -Value "Files with Errors: $errorCount"
Add-Content -Path $LogFile -Value "Files with .txt: $txtFoundCount"
Add-Content -Path $LogFile -Value "============================================"

# Console Summary
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "         PROCESSING COMPLETE                " -ForegroundColor Cyan
Write-Host "Files Processed:   $processedCount" -ForegroundColor Green
Write-Host "Files with Errors: $errorCount" -ForegroundColor Red
Write-Host "Files with .txt:   $txtFoundCount" -ForegroundColor Yellow
Write-Host "Total Duration:    $($totalDuration.TotalMinutes.ToString('F2')) minutes" -ForegroundColor Cyan
Write-Host "Log file saved to: $LogFile" -ForegroundColor Cyan
Write-Host "============================================"
