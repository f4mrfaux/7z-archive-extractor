# Threaded Archive Extractor

A powerful PowerShell script that recursively scans directories for archive files (ZIP, RAR, 7Z, etc.) and extracts them using multiple CPU threads for maximum performance.

## Features

- **Multi-threaded processing**: Automatically uses an optimal number of CPU cores for parallel scanning and extraction
- **Smart resource management**: Dynamically adjusts thread usage based on your system's capabilities
- **Comprehensive progress tracking**: Real-time progress bars for both scanning and extraction phases
- **Detailed logging**: Creates timestamped logs of all operations for troubleshooting
- **Robust error handling**: Gracefully handles inaccessible paths and extraction failures
- **Path compatibility**: Works with complex file paths containing special characters

## Requirements

- Windows PowerShell 5.1 or later
- [7-Zip](https://www.7-zip.org/) installed and available in PATH
- Windows 7/Server 2008 R2 or newer

## Supported Archive Formats

- ZIP (.zip)
- RAR (.rar)
- 7-Zip (.7z)
- TAR (.tar)
- GZip (.gz)
- BZip2 (.bz2)

## Installation

1. Download the script to your computer
2. Ensure 7-Zip is installed and in your system PATH
3. Save the script as `Threaded-Archive-Extractor.ps1` with UTF-8 (with BOM) encoding

## Usage

1. Open PowerShell
2. Navigate to the directory containing the script
3. Run the script:
```powershell
.\Threaded-Archive-Extractor.ps1
```
4. Enter the full path to the directory you want to scan when prompted
5. The script will handle the rest, showing progress as it works

## How It Works

1. **Initialization**: The script checks for 7-Zip and prompts for a directory to scan
2. **Resource Detection**: Automatically determines the optimal number of threads to use based on your CPU
3. **Folder Scanning**: Performs a multi-threaded scan of all subfolders looking for archive files
4. **Parallel Extraction**: Distributes archive extraction tasks across multiple threads
5. **Summary Report**: Provides a detailed report of successful and failed extractions with a log file for reference

## Advanced Features

- **Throttling**: Prevents system overload by limiting the maximum number of concurrent operations
- **Thread-safe operations**: Uses synchronized collections and interlocked operations for data integrity
- **Proper cleanup**: Ensures all resources are properly disposed even if errors occur
- **Performance optimization**: Balances speed and system responsiveness

## Troubleshooting

If you encounter issues:

1. Check the log file (location displayed when the script runs)
2. Ensure 7-Zip is correctly installed and in your PATH
3. Try running PowerShell as administrator if you're scanning system directories
4. For password-protected or corrupted archives, try extracting them manually with 7-Zip

## License

MIT License - Feel free to modify and distribute as needed

## Acknowledgments

- This script utilizes 7-Zip (https://www.7-zip.org/) for archive extraction
- Special thanks to the PowerShell community for runspace and threading techniques
