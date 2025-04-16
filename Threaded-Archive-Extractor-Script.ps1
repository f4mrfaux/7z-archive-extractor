# Threaded Archive Extractor Script
# Extracts archive files (ZIP, RAR, 7Z, etc.) recursively from a directory.
# Using multi-threading for improved performance.
# Requires 7-Zip to be installed and available in PATH.
# Version: 2.1
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

# Get system info to determine optimal thread count - using newer CimInstance method
try {
    $cpuCores = (Get-CimInstance -ClassName Win32_Processor -ErrorAction Stop).NumberOfLogicalProcessors
}
catch {
    # Fallback to WMI if CimInstance isn't available
    $cpuCores = (Get-WmiObject -Class Win32_Processor -ErrorAction Stop).NumberOfLogicalProcessors
}

# Use 75% of available cores, with a minimum of 2 and maximum of 16
$maxThreads = [Math]::Max(2, [Math]::Min(16, [Math]::Floor($cpuCores * 0.75)))
Write-Host "System has $cpuCores logical processors, using $maxThreads threads for operations" -ForegroundColor Cyan

# Create a log file
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logFile = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "ArchiveExtractor-$timestamp.log"
Write-Host "Logging to: $logFile" -ForegroundColor Cyan

# Function to log messages
function Write-Log {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $logFile -Value $logMessage
    
    # Also write to console if not in verbose mode to avoid cluttering
    switch ($Level) {
        "INFO" { 
            # Don't echo regular info messages to console
        }
        "WARNING" { Write-Host $Message -ForegroundColor Yellow }
        "ERROR" { Write-Host $Message -ForegroundColor Red }
    }
}

# Supported archive extensions
$archiveExtensions = @("*.zip", "*.rar", "*.7z", "*.tar", "*.gz", "*.bz2")
$archiveList = ($archiveExtensions -replace '^\*\.', '') -join ', '

Write-Host "Scanning for supported archive formats: $archiveList" -ForegroundColor Cyan
Write-Log "Starting scan for formats: $archiveList"

try {
    # Initialize synchronized collections for thread-safe operations
    $syncArchives = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $syncHash = [hashtable]::Synchronized(@{})
    $syncHash.TotalFolders = 0
    $syncHash.ProcessedFolders = 0
    $syncHash.TotalArchives = 0
    $syncHash.RunspacePool = $null
    $syncHash.Jobs = @()
    $syncHash.SuccessCount = 0
    $syncHash.FailCount = 0
    $syncHash.LogFile = $logFile
    
    # First, count total folders for progress reporting
    Write-Host "Counting folders for progress tracking..." -ForegroundColor Yellow
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($folderPath)
    
    while ($queue.Count -gt 0) {
        $currentFolder = $queue.Dequeue()
        $syncHash.TotalFolders++
        
        try {
            $subFolders = Get-ChildItem -Path $currentFolder -Directory -Force -ErrorAction SilentlyContinue
            foreach ($subFolder in $subFolders) {
                $queue.Enqueue($subFolder.FullName)
            }
        }
        catch {
            # Log inaccessible folders
            Write-Log "Cannot access folder: $currentFolder - $($_.Exception.Message)" -Level "WARNING"
            continue
        }
    }
    
    Write-Host "Found $($syncHash.TotalFolders) folders to scan" -ForegroundColor Cyan
    Write-Log "Found $($syncHash.TotalFolders) folders to scan"
    
    # Set up runspace pool for parallel processing
    $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $syncHash.RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $maxThreads, $sessionState, $Host)
    $syncHash.RunspacePool.Open()
    
    # Create a timer for progress updates
    $timer = New-Object System.Timers.Timer
    $timer.Interval = 500 # Update every 500ms
    $timer.AutoReset = $true
    
    # Update progress bar when timer ticks
    $updateProgress = {
        $percentComplete = if ($syncHash.TotalFolders -gt 0) { 
            [Math]::Min(100, [Math]::Floor(($syncHash.ProcessedFolders / $syncHash.TotalFolders) * 100)) 
        } else { 
            0 
        }
        
        Write-Progress -Activity "Scanning folders for archives" -Status "Progress: $percentComplete%" `
                      -PercentComplete $percentComplete -CurrentOperation "$($syncHash.ProcessedFolders) of $($syncHash.TotalFolders) folders"
    }
    
    $timer.Elapsed.Add($updateProgress)
    $timer.Start()
    
    # Function to scan a folder for archives
    $scanFolder = {
        param(
            [string]$folder, 
            [string[]]$archiveExtensions, 
            [hashtable]$syncHash, 
            [System.Collections.Concurrent.ConcurrentBag[object]]$syncArchives
        )
        
        # Function to write to log from within the runspace
        function WriteToLog {
            param([string]$message, [string]$level = "INFO")
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$level] $message"
            Add-Content -Path $syncHash.LogFile -Value $logMessage -ErrorAction SilentlyContinue
        }
        
        try {
            # Look for archives in current folder
            $folderArchives = Get-ChildItem -Path $folder -Include $archiveExtensions -File -Force -ErrorAction SilentlyContinue
            if ($folderArchives) {
                foreach ($archive in $folderArchives) {
                    $syncArchives.Add($archive)
                    [System.Threading.Interlocked]::Increment([ref]$syncHash.TotalArchives)
                    WriteToLog "Found archive: $($archive.FullName)"
                }
            }
        }
        catch {
            # Log error for this folder
            WriteToLog "Error scanning folder $folder`: $($_.Exception.Message)" "WARNING"
        }
        
        # Increment the processed folder count
        [System.Threading.Interlocked]::Increment([ref]$syncHash.ProcessedFolders)
    }
    
    # Queue up scan jobs for each folder
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($folderPath)
    $processedFolders = [System.Collections.Generic.HashSet[string]]::new()
    
    Write-Log "Starting parallel folder scan with $maxThreads threads"
    
    while ($queue.Count -gt 0) {
        $currentFolder = $queue.Dequeue()
        
        # Skip if we've already processed this folder
        if ($processedFolders.Contains($currentFolder)) {
            continue
        }
        
        $processedFolders.Add($currentFolder)
        
        # Create and start a new runspace for this folder
        $job = [powershell]::Create().AddScript($scanFolder).AddParameters(@{
            folder = $currentFolder
            archiveExtensions = $archiveExtensions
            syncHash = $syncHash
            syncArchives = $syncArchives
        })
        
        $job.RunspacePool = $syncHash.RunspacePool
        $syncHash.Jobs += @{
            PowerShell = $job
            Handle = $job.BeginInvoke()
        }
        
        try {
            # Add subfolders to the queue
            $subFolders = Get-ChildItem -Path $currentFolder -Directory -Force -ErrorAction SilentlyContinue
            foreach ($subFolder in $subFolders) {
                if (-not $processedFolders.Contains($subFolder.FullName)) {
                    $queue.Enqueue($subFolder.FullName)
                }
            }
        }
        catch {
            # Log error for inaccessible subfolders
            Write-Log "Cannot access subfolders in $currentFolder`: $($_.Exception.Message)" -Level "WARNING"
            continue
        }
        
        # Throttle job creation if we have too many pending
        while ($syncHash.Jobs.Count -ge ($maxThreads * 2)) {
            # Clean up completed jobs
            $remainingJobs = @()
            foreach ($jobInfo in $syncHash.Jobs) {
                if ($jobInfo.Handle.IsCompleted) {
                    # End the invoke and dispose properly
                    $jobInfo.PowerShell.EndInvoke($jobInfo.Handle)
                    $jobInfo.PowerShell.Dispose()
                }
                else {
                    $remainingJobs += $jobInfo
                }
            }
            $syncHash.Jobs = $remainingJobs
            
            if ($syncHash.Jobs.Count -ge ($maxThreads * 2)) {
                Start-Sleep -Milliseconds 100
            }
        }
    }
    
    # Wait for all scan jobs to complete
    Write-Host "Waiting for all scan jobs to complete..." -ForegroundColor Yellow
    Write-Log "Waiting for scanning jobs to complete"
    
    while ($syncHash.Jobs.Count -gt 0) {
        $remainingJobs = @()
        
        foreach ($jobInfo in $syncHash.Jobs) {
            if ($jobInfo.Handle.IsCompleted) {
                # End the invoke and dispose properly
                $jobInfo.PowerShell.EndInvoke($jobInfo.Handle)
                $jobInfo.PowerShell.Dispose()
            }
            else {
                $remainingJobs += $jobInfo
            }
        }
        
        $syncHash.Jobs = $remainingJobs
        
        if ($syncHash.Jobs.Count -gt 0) {
            Start-Sleep -Milliseconds 100
        }
    }
    
    # Stop the timer and complete the progress bar
    $timer.Stop()
    $timer.Dispose()
    Write-Progress -Activity "Scanning folders for archives" -Completed
    
    # Convert the concurrent bag to an array
    $archives = $syncArchives.ToArray()
    
    if ($archives.Count -eq 0) {
        Write-Host "No archives found in the selected directory." -ForegroundColor Yellow
        Write-Log "No archives found. Exiting." -Level "INFO"
        exit 0
    }

    $total = $archives.Count
    Write-Host "Found $total archive(s). Beginning extraction..." -ForegroundColor Cyan
    Write-Host "Using $maxThreads parallel extraction threads" -ForegroundColor Cyan
    Write-Host "Press Ctrl+C to abort the operation at any time." -ForegroundColor DarkYellow
    Write-Log "Starting extraction of $total archives using $maxThreads threads"
    
    # Create a timer for extraction progress updates
    $extractTimer = New-Object System.Timers.Timer
    $extractTimer.Interval = 500 # Update every 500ms
    $extractTimer.AutoReset = $true
    
    # Update progress bar when timer ticks
    $updateExtractionProgress = {
        $currentProgress = $syncHash.SuccessCount + $syncHash.FailCount
        $percentComplete = if ($total -gt 0) { 
            [Math]::Min(100, [Math]::Floor(($currentProgress / $total) * 100)) 
        } else { 
            0 
        }
        
        Write-Progress -Activity "Extracting archives" -Status "Progress: $percentComplete%" `
                      -PercentComplete $percentComplete `
                      -CurrentOperation "Completed: $currentProgress of $total archives (Success: $($syncHash.SuccessCount), Failed: $($syncHash.FailCount))"
    }
    
    $extractTimer.Elapsed.Add($updateExtractionProgress)
    $extractTimer.Start()
    
    # Clear job list for extraction jobs
    $syncHash.Jobs = @()
    
    # Create a synchronized queue for archives to process
    $extractQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
    foreach ($archive in $archives) {
        $extractQueue.Enqueue($archive)
    }
    
    # Function to extract an archive
    $extractArchive = {
        param(
            [System.Collections.Concurrent.ConcurrentQueue[object]]$extractQueue, 
            [hashtable]$syncHash
        )
        
        # Function to write to log from within the runspace
        function WriteToLog {
            param([string]$message, [string]$level = "INFO")
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$level] $message"
            Add-Content -Path $syncHash.LogFile -Value $logMessage -ErrorAction SilentlyContinue
        }
        
        # Process archives from the queue until it's empty
        while ($true) {
            $archive = $null
            $gotItem = $extractQueue.TryDequeue([ref]$archive)
            
            if (-not $gotItem) {
                # Queue is empty, exit the loop
                break
            }
            
            $extractTo = $archive.Directory.FullName
            $archivePath = $archive.FullName
            
            try {
                # Process the archive
                $processArgs = "x `"$archivePath`" -o`"$extractTo`" -aos -y -bso0 -bsp1"
                
                # Log the extraction attempt
                WriteToLog "Extracting: $archivePath to $extractTo"
                
                $proc = Start-Process -FilePath "7z.exe" -ArgumentList $processArgs -NoNewWindow -Wait -PassThru -ErrorAction Stop
                
                if ($proc.ExitCode -eq 0) {
                    # Write to console in a thread-safe way
                    [Console]::WriteLine("Success: $($archive.Name) extracted to $extractTo")
                    WriteToLog "Successfully extracted: $archivePath" "INFO"
                    [System.Threading.Interlocked]::Increment([ref]$syncHash.SuccessCount)
                }
                else {
                    # Write to console in a thread-safe way
                    $errorMessage = "Failed to extract $($archive.Name) (Exit code: $($proc.ExitCode))"
                    [Console]::WriteLine($errorMessage)
                    
                    # Add detailed error information
                    $errorDetail = "Unknown error"
                    switch ($proc.ExitCode) {
                        1 { $errorDetail = "Warning (Non-fatal error)" }
                        2 { $errorDetail = "Fatal error occurred" }
                        7 { $errorDetail = "Command line error" }
                        8 { $errorDetail = "Not enough memory" }
                        255 { $errorDetail = "User stopped the process" }
                    }
                    
                    WriteToLog "Extraction failed: $archivePath - $errorDetail" "WARNING"
                    [System.Threading.Interlocked]::Increment([ref]$syncHash.FailCount)
                }
            }
            catch {
                # Write to console in a thread-safe way
                $errorMessage = "Error extracting $($archive.Name): $_"
                [Console]::WriteLine($errorMessage)
                WriteToLog "Exception during extraction: $archivePath - $($_.Exception.Message)" "ERROR"
                [System.Threading.Interlocked]::Increment([ref]$syncHash.FailCount)
            }
        }
    }
    
    # Start extraction worker threads
    for ($i = 0; $i -lt $maxThreads; $i++) {
        $job = [powershell]::Create().AddScript($extractArchive).AddParameters(@{
            extractQueue = $extractQueue
            syncHash = $syncHash
        })
        
        $job.RunspacePool = $syncHash.RunspacePool
        $syncHash.Jobs += @{
            PowerShell = $job
            Handle = $job.BeginInvoke()
        }
    }
    
    # Wait for all extraction jobs to complete
    Write-Host "Extracting archives... Please wait." -ForegroundColor Yellow
    Write-Log "Waiting for extraction jobs to complete"
    
    while ($syncHash.Jobs.Count -gt 0) {
        $remainingJobs = @()
        
        foreach ($jobInfo in $syncHash.Jobs) {
            if ($jobInfo.Handle.IsCompleted) {
                # End the invoke and dispose properly
                $jobInfo.PowerShell.EndInvoke($jobInfo.Handle)
                $jobInfo.PowerShell.Dispose()
            }
            else {
                $remainingJobs += $jobInfo
            }
        }
        
        $syncHash.Jobs = $remainingJobs
        
        if ($syncHash.Jobs.Count -gt 0) {
            Start-Sleep -Milliseconds 100
        }
    }
    
    # Stop the extraction timer and complete the progress bar
    $extractTimer.Stop()
    $extractTimer.Dispose()
    Write-Progress -Activity "Extracting archives" -Completed
    
    # Close the runspace pool
    $syncHash.RunspacePool.Close()
    $syncHash.RunspacePool.Dispose()
    
    # Summary report
    $summaryMessage = @"

Extraction Summary:
Successfully extracted: $($syncHash.SuccessCount)
Failed extractions: $($syncHash.FailCount)
Total archives processed: $total
"@
    
    Write-Host $summaryMessage -ForegroundColor Cyan
    Write-Log $summaryMessage
    
    if ($syncHash.FailCount -gt 0) {
        $tipMessage = @"
Tip: Some archives might be password protected or corrupted.
     Try extracting them manually with 7-Zip for more details.
     Check the log file for specific errors: $logFile
"@
        Write-Host $tipMessage -ForegroundColor Yellow
        Write-Log "Some extractions failed. See log for details."
    }
    else {
        Write-Host "All archives were extracted successfully!" -ForegroundColor Green
        Write-Log "All extractions completed successfully."
    }
}
catch {
    $errorMessage = "Error during execution: $_"
    Write-Host $errorMessage -ForegroundColor Red
    Write-Log $errorMessage -Level "ERROR"
    exit 1
}
finally {
    # Ensure all resources are properly disposed
    if ($null -ne $timer -and $timer -is [System.Timers.Timer]) {
        $timer.Stop()
        $timer.Dispose()
    }
    
    if ($null -ne $extractTimer -and $extractTimer -is [System.Timers.Timer]) {
        $extractTimer.Stop()
        $extractTimer.Dispose()
    }
    
    # Clean up any remaining jobs and runspaces
    if ($null -ne $syncHash -and $null -ne $syncHash.Jobs) {
        foreach ($jobInfo in $syncHash.Jobs) {
            if ($null -ne $jobInfo.PowerShell) {
                if ($jobInfo.Handle.IsCompleted) {
                    $jobInfo.PowerShell.EndInvoke($jobInfo.Handle)
                }
                else {
                    $jobInfo.PowerShell.Stop()
                }
                $jobInfo.PowerShell.Dispose()
            }
        }
    }
    
    if ($null -ne $syncHash -and $null -ne $syncHash.RunspacePool) {
        $syncHash.RunspacePool.Close()
        $syncHash.RunspacePool.Dispose()
    }
    
    Write-Host "Script execution completed. Log file: $logFile" -ForegroundColor Cyan
    Write-Log "Script execution completed."
}

exit 0