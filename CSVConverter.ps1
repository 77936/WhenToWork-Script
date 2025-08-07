#Employee Hashtable Biweekly Names -> WhenToWork Names + Positions
$WorkerTable = @{
   "Nicole Shaw" = [PSCustomObject]@{
        Name = "Nicole Shaw"
        Position = "NB Pool Manager"
    }
   "Michael Coash" = [PSCustomObject]@{
        Name = "Michael Coash"
        Position = "NB Pool Manager"
    }
   "Eno Linsky (CH)" = [PSCustomObject]@{
        Name = "Eno Linsky"
        Position = "NB Pool Guard II"
    }
   "James Stewart" = [PSCustomObject]@{
        Name = "James Stewart"
        Position = "NB Rec Aide"
    }
   "Nicholas Nguyen" = [PSCustomObject]@{
        Name = "Nicholas Nguyen"
        Position = "NB Pool Guard II"
    }
   "Gaby Gonzalez" = [PSCustomObject]@{
        Name = "Gabriela Gonzalez"
        Position = "NB Pool Guard II"
    }
   "Colin Phung" = [PSCustomObject]@{
        Name = "Colin Phung"
        Position = "NB Pool Guard II"
    }
   "Derek Phan" = [PSCustomObject]@{
        Name = "Derek Pham"
        Position = "NB Pool Guard II"
    }
   "Maclovio Atilano" = [PSCustomObject]@{
        Name = "Maclovio Atilano"
        Position = "NB Pool Guard II"
    }
   "Victoria Coker" = [PSCustomObject]@{
        Name = "Victoria Coker"
        Position = "NB Pool Guard II"
    }
   "Rhilo Sotto" = [PSCustomObject]@{
        Name = "Rhilo Sotto"
        Position = "NB Pool Guard II"
    }
   "Axel Pedroza" = [PSCustomObject]@{
        Name = "Axel Pedroza"
        Position = "NB Pool Guard II"
    }
   "Hyein Choi" = [PSCustomObject]@{
        Name = "Hyein Choi"
        Position = "NB Pool Guard I"
    }
   "Sofie Salazar" = [PSCustomObject]@{
        Name = "Sofie Salazar"
        Position = "NB Pool Guard II"
    }
   "Manuel Alvarez" = [PSCustomObject]@{
        Name = "Manuel Alvarez"
        Position = "NB Pool Guard II"
    }
   "Trey Pavlik" = [PSCustomObject]@{
        Name = "Trey Pavlik"
        Position = "NB Pool Guard I"
    }
   "Santiago Nava Estevez (CM)" = [PSCustomObject]@{
        Name = "Santiago Esteves Nava"
        Position = "NB Pool Guard I"
    }
   "Dominic Lenguyen" = [PSCustomObject]@{
        Name = "Dominic Lenguyen"
        Position = "NB Pool Guard I"
    }
   "Yun Seo" = [PSCustomObject]@{
        Name = "Yun Seo"
        Position = "NB Pool Guard II"
    }
   "Alexander Forsman" = [PSCustomObject]@{
        Name = "Alexander Forsman"
        Position = "NB Pool Guard I"
    }
   "Alexander Dubin" = [PSCustomObject]@{
        Name = "Alexander Dubin"
        Position = "NB Pool Guard I"
    }
}

# Helper function
# Parse-CellAddress: Converts "A1" to Column 1, Row 1
function Parse-CellAddress {
    param([string]$CellAddress)
    
    if ($CellAddress -match '^([A-Z]+)(\d+)$') {
        $columnLetters = $matches[1]
        $rowNumber = [int]$matches[2]
        
        $columnNumber = 0
        for ($i = 0; $i -lt $columnLetters.Length; $i++) {
            $columnNumber = $columnNumber * 26 + ([byte][char]$columnLetters[$i] - [byte][char]'A' + 1)
        }
        
        return @{
            Column = $columnNumber
            Row = $rowNumber
            OriginalAddress = $CellAddress
        }
    } else {
        throw "Invalid cell address format: $CellAddress"
    }
}

# Helper function
# Get-CellAddress: Converts Column 1, Row 1 back to "A1"
function Get-CellAddress {
    param([int]$Column, [int]$Row)
    
    $columnLetter = ""
    $temp = $Column
    
    while ($temp -gt 0) {
        $temp--
        $columnLetter = [char]([byte][char]'A' + ($temp % 26)) + $columnLetter
        $temp = [math]::Floor($temp / 26)
    }
    
    return "$columnLetter$Row"
}

# Helper function to get cell value from ImportExcel data
function Get-CellValueFromImportedData {
    param(
        [array]$Data,
        [int]$Row,        # 0-based row index
        [int]$Column      # 1-based column index
    )
    
    # Check if row exists
    if ($Row -lt 0 -or $Row -ge $Data.Count) {
        return $null
    }
    
    # Get the row object
    $rowData = $Data[$Row]
    
    # ImportExcel creates properties like P1, P2, P3... for columns
    $columnProperty = "P$Column"
    
    if ($rowData.PSObject.Properties.Name -contains $columnProperty) {
        return $rowData.$columnProperty
    }
    
    return $null
}

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

# Function to display Excel file summary
function Show-ExcelSummary {
    param(
        [Parameter(Mandatory=$true)]
        $ExcelData
    )
    
    if ($null -eq $ExcelData -or $null -eq $ExcelData.Data) {
        Write-Host "No Excel data to display" -ForegroundColor Red
        return
    }
    
    Write-Host "`n=== Excel File Summary ===" -ForegroundColor Cyan
    Write-Host "File: $(Split-Path $ExcelData.FilePath -Leaf)" -ForegroundColor White
    Write-Host "Worksheet: $($ExcelData.Worksheet)" -ForegroundColor White
    Write-Host "Rows: $($ExcelData.RowCount)" -ForegroundColor White
    Write-Host "Columns: $($ExcelData.Columns.Count)" -ForegroundColor White
    
    # Show first few rows as preview
    if ($ExcelData.Data.Count -gt 0) {
        Write-Host "`nData Preview (first 3 rows):" -ForegroundColor Yellow
        $ExcelData.Data | Select-Object -First 3 | Format-Table -AutoSize
    }
}

# Function to parse time shifts for schedule mode *ADD LOCATION CODES TO "Category" Header*
function Parse-TimeShift {
    param(
        [string]$shift
    )
    
    # Return null for empty, or letters with no shift time, ex. OFF
    if ([string]::IsNullOrWhiteSpace($shift) -or $shift -match "[a-zA-Z]") {
        return $null
    }
    
    # Handles formats like "7:30-4 NB", "2-8:30 TCP", "9:00-5:30 ABC" and splits into shift time range & location code
    if ($CellValue -match '^([\d:]+\-[\d:]+)\s+(.+)$') {
        $cleanShift = $matches[1]
        $category = $matches[2]
    } else{
        $cleanShift = $shift
    }

    # Match time patterns like "9:30-4:30", "7am-3pm", "2-9pm", etc.
    $timePattern = '(\d{1,2}:?\d{0,2})\s*(?:am|pm)?\s*-\s*(\d{1,2}:?\d{0,2})\s*(am|pm)?'
    
    if ($cleanShift -match $timePattern) {
        $startTime = $matches[1]
        $endTime = $matches[2]
        $endPeriod = $matches[3]
        
        # Add missing colons for times without minutes
        if ($startTime -notmatch ':') {
            $startTime = $startTime + ':00'
        }
        if ($endTime -notmatch ':') {
            $endTime = $endTime + ':00'
        }
        
        # Handle AM/PM logic
        if ($endPeriod) {
            $endTime = $endTime + $endPeriod.ToLower()
            
            # If end time is PM and start time doesn't have AM/PM, assume start is AM
            if ($endPeriod.ToLower() -eq 'pm' -and $cleanShift -notmatch 'am' -and $startTime -notmatch 'am|pm') {
                # For early morning hours (6-11), assume AM
                $startHour = [int]($startTime -split ':')[0]
                if ($startHour -ge 6 -and $startHour -le 11) {
                    $startTime = $startTime + 'am'
                } else {
                    # For afternoon hours or when ambiguous, keep as is and let context decide
                    $startTime = $startTime + 'pm'
                }
            }
        } else {
            # No AM/PM specified, make educated guesses
            $startHour = [int]($startTime -split ':')[0]
            $endHour = [int]($endTime -split ':')[0]
            
            # Early hours (6-11) are likely AM
            if ($startHour -ge 6 -and $startHour -le 11) {
                $startTime = $startTime + 'am'
            }
            
            # If end hour is less than start hour, end is likely next period
            if ($endHour -le 12 -and $endHour -lt $startHour) {
                $endTime = $endTime + 'pm'
            } elseif ($endHour -ge 13) {
                # 24-hour format converted
                $endTime = ($endHour - 12).ToString() + ':' + ($endTime -split ':')[1] + 'pm'
            }
        }
        
        return @{
            StartTime = $startTime
            EndTime = $endTime
            Category = $category

        }
    }
    
    # If no pattern matches, return null
    Write-Warning "Could not parse shift: '$shift'"
    return $null
}

# TODO: Cell Traversal and add "Shift Description" and "Paid Hours"
function cellTraversal{
    param(
        [string]$startingCell,
        [string]$WorksheetName,
        [string]$FilePath
    )

    $result = [PSCustomObject]@{
        shift1 = ""
        shift2 = ""
        description = ""
        paidHours = ""
    }

    $cellInfo = Parse-CellAddress -CellAddress $startingCell
    $currentColumn = $cellInfo.Column
    $currentColumn = $cellInfo.Row

    Write-Host "Processing Excel file: $FilePath" -ForegroundColor Green
    Write-Host "Starting from cell: $StartingCell (Column $currentColumn, Row $currentRow)" -ForegroundColor Yellow
    
    $arrayRow = $currentRow - 1


}

# Main execution
function Main{
    Write-Host "=== Excel File Selector with Auto First Worksheet Selection ===" -ForegroundColor Cyan
    Write-Host "This script will open a file dialog to select Excel files and automatically use the first worksheet.`n" -ForegroundColor White

    # Check ImportExcel module
    if (-not (Test-ImportExcelModule)) {
        exit 1
    }

    # Get Excel files using dialog
    $filesToProcess = Get-ExcelFilesDialog

    if ($filesToProcess.Count -eq 0) {
        Write-Host "No Excel files selected. Exiting." -ForegroundColor Yellow
        exit 0
    }

    Write-Host "`n=== Processing Selected Files ===" -ForegroundColor Cyan

    # Process each selected file
    foreach ($file in $filesToProcess) {
        $excelData = Open-ExcelFile -FilePath $file
    
        if ($null -ne $excelData) {
            Show-ExcelSummary -ExcelData $excelData
        
            # Here you would add your parsing logic
            Write-Host "`n[Ready for parsing logic - Excel data is loaded and available]" -ForegroundColor Magenta
        
            # Ask if user wants to continue with next file (if multiple files)
            if ($filesToProcess.Count -gt 1 -and $file -ne $filesToProcess[-1]) {
                Write-Host "`nPress Enter to continue to next file, or 'q' to quit..." -ForegroundColor Yellow
                $continue = Read-Host
                if ($continue -eq 'q') {
                    Write-Host "Processing stopped by user." -ForegroundColor Yellow
                    break
                }
            }
        } else {
            Write-Host "Failed to process file: $(Split-Path $file -Leaf)" -ForegroundColor Red
        }
    }
}

Write-Host "`nProcessing complete!" -ForegroundColor Green
try {
    Main
} catch {
    Write-Error "Unexpected error: $($_.Exception.Message)"
}