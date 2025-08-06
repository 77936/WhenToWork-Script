# Check if ImportExcel module is available
function Test-ImportExcelModule {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
        try{
            Install-Module ImportExcel -Force -Scope CurrentUser
            Write-Host "ImportExcel module installed successfully." -ForegroundColor Green
            return $true
        }
        catch{
            Write-Error "Failed to install ImportExcel module. Please install manually: Install-Module ImportExcel"
            return $false
        }
    }
    Import-Module ImportExcel
    return $true
}

# Function to get Excel files using file dialog
function Get-ExcelFilesDialog {
    Write-Host "Opening file dialog to select Excel files..." -ForegroundColor Cyan
    
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        
        # Configure dialog
        $dialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
        $dialog.Multiselect = $true
        $dialog.Title = "Select Excel Files to Process"
        $dialog.InitialDirectory = Get-Location
        $dialog.CheckFileExists = $true
        $dialog.CheckPathExists = $true
        
        $result = $dialog.ShowDialog()
        
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedFiles = $dialog.FileNames
            Write-Host "Selected $($selectedFiles.Count) file(s)" -ForegroundColor Green
            
            # Display selected files
            foreach ($file in $selectedFiles) {
                Write-Host "  • $(Split-Path $file -Leaf)" -ForegroundColor White
            }
            
            return $selectedFiles
        } else {
            Write-Host "File selection cancelled by user" -ForegroundColor Yellow
            return @()
        }
        
    } catch {
        Write-Error "File dialog failed: $($_.Exception.Message)"
        Write-Host "This might happen if running in a non-Windows environment or without GUI support." -ForegroundColor Red
        
        # Fallback to manual input
        Write-Host "Falling back to manual file input..." -ForegroundColor Yellow
        $files = @()
        
        do {
            $file = Read-Host "Enter Excel file path (or 'done' to finish)"
            if ($file -ne "done" -and $file -ne "") {
                if (Test-Path $file) {
                    $extension = [System.IO.Path]::GetExtension($file).ToLower()
                    if ($extension -eq ".xlsx" -or $extension -eq ".xls") {
                        $files += $file
                        Write-Host "✓ Added: $(Split-Path $file -Leaf)" -ForegroundColor Green
                    } else {
                        Write-Host "✗ Not an Excel file: $file" -ForegroundColor Red
                    }
                } else {
                    Write-Host "✗ File not found: $file" -ForegroundColor Red
                }
            }
        } while ($file -ne "done")
        
        return $files
    }
}

# Function to get the first worksheet name
function Get-FirstWorksheet {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    try {
        # Get available worksheets
        Write-Host "Getting worksheet information..." -ForegroundColor Yellow
        $worksheets = Get-ExcelSheetInfo -Path $FilePath
        
        if (-not $worksheets -or $worksheets.Count -eq 0) {
            Write-Warning "No worksheets found in file"
            return $null
        }
        
        # Return the first worksheet
        $firstWorksheet = $worksheets[0].Name
        $rowCount = if ($worksheets[0].Rows) { $worksheets[0].Rows } else { "Unknown" }
        
        Write-Host "Auto-selecting first worksheet: '$firstWorksheet' ($rowCount rows)" -ForegroundColor Green
        
        if ($worksheets.Count -gt 1) {
            Write-Host "Note: File contains $($worksheets.Count) worksheets. Using first one only." -ForegroundColor Yellow
        }
        
        return $firstWorksheet
        
    } catch {
        Write-Error "Failed to get worksheet information: $($_.Exception.Message)"
        return $null
    }
}

# Function to validate and open Excel file
function Open-ExcelFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath
    )
    
    Write-Host "`n--- Opening Excel File ---" -ForegroundColor Cyan
    Write-Host "File: $(Split-Path $FilePath -Leaf)" -ForegroundColor White
    Write-Host "Full Path: $FilePath" -ForegroundColor Gray
    
    # Verify file exists
    if (-not (Test-Path $FilePath)) {
        Write-Error "Excel file not found: $FilePath"
        return $null
    }
    
    # Check file extension
    $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
    if ($extension -ne ".xlsx" -and $extension -ne ".xls") {
        Write-Error "File is not an Excel file: $FilePath"
        return $null
    }
    
    # Get file info
    $fileInfo = Get-Item $FilePath
    $sizeKB = [math]::Round($fileInfo.Length / 1KB, 2)
    Write-Host "File Size: $sizeKB KB" -ForegroundColor Gray
    Write-Host "Last Modified: $($fileInfo.LastWriteTime)" -ForegroundColor Gray
    
    # Get first worksheet automatically
    $selectedWorksheet = Get-FirstWorksheet -FilePath $FilePath
    
    if (-not $selectedWorksheet) {
        Write-Host "No worksheet available. Skipping file." -ForegroundColor Yellow
        return $null
    }
    
    try {
        # Read the Excel data
        Write-Host "Reading worksheet: $selectedWorksheet" -ForegroundColor Cyan
        $excelData = Import-Excel -Path $FilePath -WorksheetName $selectedWorksheet
        
        if ($null -eq $excelData -or $excelData.Count -eq 0) {
            Write-Warning "No data found in worksheet '$selectedWorksheet'"
            return $null
        }
        
        Write-Host "✓ Successfully loaded $($excelData.Count) rows" -ForegroundColor Green
        
        # Display column information
        if ($excelData.Count -gt 0) {
            $columns = $excelData[0].PSObject.Properties.Name
            Write-Host "Available columns ($($columns.Count)):" -ForegroundColor Yellow
            $columns | ForEach-Object { Write-Host "  • $_" -ForegroundColor White }
        }
        
        return @{
            Data = $excelData
            FilePath = $FilePath
            Worksheet = $selectedWorksheet
            Columns = $columns
            RowCount = $excelData.Count
        }
        
    } catch {
        Write-Error "Failed to open Excel file: $($_.Exception.Message)"
        Write-Host "Error details: $($_.Exception.InnerException.Message)" -ForegroundColor Red
        return $null
    }
}