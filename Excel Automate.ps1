#Requires -Version 7.5.4
# Key fix: Force switch to script directory to avoid System32
$ScriptPath = $MyInvocation.MyCommand.Path
$ScriptDir = Split-Path -Parent $ScriptPath
Set-Location -Path $ScriptDir
Write-Host "Current working directory: $(Get-Location)" -ForegroundColor Cyan
$ErrorActionPreference = 'Continue'
try {
    # Step 1: Check and install Bun if not present
    if (-not (Get-Command bun -ErrorAction SilentlyContinue)) {
        Write-Host "Bun not found. Installing Bun..." -ForegroundColor Yellow
        irm bun.sh/install.ps1 | iex
        $env:Path += ";$env:USERPROFILE\.bun\bin"
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    }
    # Step 2: Initialize project if package.json doesn't exist
    if (-not (Test-Path "package.json")) {
        Write-Host "Initializing Bun project..." -ForegroundColor Yellow
        bun init -y
    }
    # Step 3: Install dependencies
    $packages = @('xlsx', 'exceljs', 'xlsx-populate', 'polars', 'chart.js', 'canvas', 'chartjs-node-canvas', 'chartjs-plugin-datalabels')
    $total = $packages.Count
    for ($i = 0; $i -lt $total; $i++) {
        $percent = [math]::Round((($i + 1) / $total) * 100)
        Write-Progress -Activity "Installing Dependencies" -Status "$($packages[$i]) ($($i+1)/$total)" -PercentComplete $percent
        bun add $packages[$i]
        Start-Sleep -Milliseconds 300
    }
    Write-Progress -Activity "Installing Dependencies" -Completed
    # Key addition: Clear the console after installation to clean up logs
    Clear-Host
    Write-Host "Dependencies installed successfully. Console cleared." -ForegroundColor Green
    # Step 4: Generate main.js only if it doesn't exist
    if (-not (Test-Path ".\main.js")) {
        $jsContent = @'
import XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import XlsxPopulate from 'xlsx-populate';
console.log('All libraries imported successfully!');
const filePath = process.argv[2]?.trim();
if (!filePath) {
    console.log('Usage: bun run main.js "C:\\path\\to\\your\\file.xlsx"');
    process.exit(1);
}
try {
    const workbook = XLSX.readFile(filePath);
    console.log(`SheetJS success! Sheets: ${workbook.SheetNames.join(', ')}`);
   
    // Example with ExcelJS: Load workbook
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    console.log('ExcelJS loaded successfully.');
} catch (err) {
    console.error('Failed to process Excel file:', err.message);
}
console.log('请重写main.js脚本');
'@
        Set-Content -Path ".\main.js" -Value $jsContent -Encoding UTF8 -Force
        Write-Host "`nmain.js has been successfully created in:" -ForegroundColor Green
        Write-Host " $(Resolve-Path .\main.js)" -ForegroundColor Green
    }
    # Step 5: Loop for prompting Excel file
    do {
        Write-Host "`nPlease drag and drop your Excel file here (or paste the full path; leave blank to exit):" -ForegroundColor Cyan
        $filePath = Read-Host " Path"
        if (-not $filePath) {
            break
        }
        $filePath = $filePath.Trim('"').Trim("'").Trim()
        if (Test-Path $filePath) {
            Write-Host "`nProcessing your Excel file..." -ForegroundColor Green
            bun run main.js "$filePath"
            Write-Host "`nProcessing completed. Please Re-Write `main.js` and drag in the Excel file again:" -ForegroundColor Magenta
        } else {
            Write-Host "File not found: $filePath" -ForegroundColor Red
        }
    } while ($true)
}
catch {
    Write-Host "`nCritical error occurred:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host $_.ScriptStackTrace
}
finally {
    Write-Host "`n`nScript execution completed." -ForegroundColor Magenta
    Write-Host "Press Enter to close this window..." -ForegroundColor Yellow
    Read-Host | Out-Null
}