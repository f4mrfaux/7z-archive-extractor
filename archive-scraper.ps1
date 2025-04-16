# Simple Archive Extractor Script
# Extracts archive files (ZIP, RAR, 7Z, etc.) recursively from a directory.
# Requires 7-Zip to be installed and available in PATH.
# Version: 1.3
# Date: 2025-04-16

# Check if 7z is in PATH
try {
    $sevenZipPath = (Get-Command 7z -ErrorAction Stop).Source
    Write-Host "7-Zip found at: $sevenZipPath" -ForegroundColor Green
}
catch {
    Write-Host "7-Zip not found in PATH. Please install 7-Zip and ensure it's in your system PATH." -ForegroundColor Red
    Write-Host "Download from: https://www.7-zip.org/" -ForegroundColor Cyan
    exit 1
}

# Prompt for directory using CLI with input validation
do {
    $folderPath = Read-Host "Enter the full path to the folder you want to scan for archives"
    $folderPath = $folderPath.Trim('"') # Remove quotes if user pasted a quoted path

    if (-not (Test-Path -Path $folderPath -PathType Container)) {
        Write-Host "The provided path does not exist or is not a directory. Please try again." -ForegroundColor Red
        $folderPath = $null
    }
} while (-not $folderPath)

# Supported archive extensions
$archiveExtensions = @("*.zip", "*.rar", "*.7z", "*.tar", "*.gz", "*.bz2")
$archiveList = ($archiveExtensions -replace '^\*\.', '') -join ', '

Write-Host "Scanning for supported archive formats: $archiveList" -ForegroundColor Cyan

try {
    # Recursively find all archive files with progress display
    Write-Host "Searching for archives..." -ForegroundColor Yellow
    
    # Fixed path handling
    $archives = @()
    
    # Error handling for Get-ChildItem
    try {
        # Use -Force to handle hidden files, -ErrorAction SilentlyContinue to skip inaccessible paths
        $archives = Get-ChildItem -Path $folderPath -Recurse -Include $archiveExtensions -File -Force -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Extension -match '\.(zip|rar|7z|tar|gz|bz2)$' }
    }
    catch {
        Write-Host "Error scanning some paths. Continue with partial results." -ForegroundColor Yellow
    }
    
    if ($archives.Count -eq 0) {
        Write-Host "No archives found in the selected directory." -ForegroundColor Yellow
        exit 0
    }

    $total = $archives.Count
    $success = 0
    $fail = 0
    $current = 0

    Write-Host "Found $total archive(s). Beginning extraction..." -ForegroundColor Cyan
    Write-Host "Press Ctrl+C to abort the operation at any time." -ForegroundColor DarkYellow

    foreach ($archive in $archives) {
        $current++
        $extractTo = $archive.Directory.FullName
        $percent = [int](($current / $total) * 100)
        $progress = "[$current of $total] "

        # Show progress bar
        Write-Progress -Activity "Extracting archives..." -Status $progress -PercentComplete $percent -CurrentOperation $archive.FullName
        
        # Handling paths with special characters
        $archivePath = $archive.FullName
        $destinationPath = $extractTo

        try {
            # More robust command execution with special character handling
            $processArgs = "x `"$archivePath`" -o`"$destinationPath`" -aos -y -bso0 -bsp1"
            $proc = Start-Process -FilePath "7z.exe" -ArgumentList $processArgs -NoNewWindow -Wait -PassThru -ErrorAction Stop

            if ($proc.ExitCode -eq 0) {
                Write-Host "Success: $($archive.Name) extracted to $extractTo" -ForegroundColor Green
                $success++
            }
            else {
                Write-Host "Failed to extract $($archive.Name) (Exit code: $($proc.ExitCode))" -ForegroundColor Yellow
                $fail++

                # Provide more detailed error information
                switch ($proc.ExitCode) {
                    1 { Write-Host "Warning (Non-fatal error)" -ForegroundColor DarkYellow }
                    2 { Write-Host "Fatal error occurred" -ForegroundColor Red }
                    7 { Write-Host "Command line error" -ForegroundColor Red }
                    8 { Write-Host "Not enough memory" -ForegroundColor Red }
                    255 { Write-Host "User stopped the process" -ForegroundColor DarkYellow }
                    default { Write-Host "Unknown error" -ForegroundColor Red }
                }
            }
        }
        catch {
            Write-Host "Error extracting $($archive.Name): $_" -ForegroundColor Red
            $fail++
            continue
        }
    }

    # Final progress bar at 100%
    Write-Progress -Activity "Extraction complete" -Completed

    # Summary report
    Write-Host "Extraction Summary:" -ForegroundColor Cyan
    Write-Host "Successfully extracted: $success" -ForegroundColor Green
    Write-Host "Failed extractions: $fail" -ForegroundColor Red
    Write-Host "Total archives processed: $total"
    
    if ($fail -gt 0) {
        Write-Host "Tip: Some archives might be password protected or corrupted." -ForegroundColor Yellow
        Write-Host "Try extracting them manually with 7-Zip for more details." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Error during execution: $_" -ForegroundColor Red
    exit 1
}
finally {
    Write-Host "Script execution completed." -ForegroundColor Cyan
}

exit 0