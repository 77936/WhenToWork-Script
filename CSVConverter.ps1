# Done: Employee Hashtable Biweekly Names -> WhenToWork Names + Positions
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

# Done: Check if ImportExcel module is available
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

# Done: Function to open file dialog and select Excel file
function Select-ExcelFile {
    param(
        [string]$Title = "Select Excel File",
        [string]$InitialDirectory = [Environment]::GetFolderPath("Desktop")
    )
    
    Add-Type -AssemblyName System.Windows.Forms
    
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.Title = $Title
    $fileDialog.InitialDirectory = $InitialDirectory
    $fileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
    $fileDialog.FilterIndex = 1
    $fileDialog.Multiselect = $false
    
    $result = $fileDialog.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $fileDialog.FileName
    }
    else {
        Write-Host "No file selected." -ForegroundColor Yellow
        return $null
    }
}

# TODO: Helper function for column incremention

# TODO: Function to parse through shift cell group
function Parse-CellGroup{
    param(
        [Parameter(Mandatory=$true)]
        $Worksheet,
        [Parameter(Mandatory=$true)]
        [int]$StartRow,
        [Parameter(Mandatory=$true)]
        [int]$StartColumn
    )


}

# Done: Function to parse time shifts for schedule mode *ADD LOCATION CODES TO "Category" Header*
# Returns StartTime, EndTime, Category
function Parse-Time-Location {
    param(
        [string]$time
    )

    $category = $null
    
    # Return null for empty, or letters with no shift time, ex. OFF
    if ([string]::IsNullOrWhiteSpace($shift) -or $shift -match "[a-zA-Z]") {
        return $null
    }
    
    # Handles formats like "7:30-4 NB", "2-8:30 TCP", "9:00-5:30 ABC" and splits into shift time range & location code
    if ($time -match '^([\d:]+\-[\d:]+)\s+(.+)$') {
        $cleanShift = $matches[1]
        $category = $matches[2]
    } else{
        $cleanShift = $time
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




# Main function & Tester Main Functions
function Main{
    # Step 1: Check ImportExcel module
    if (-not (Test-ImportExcelModule)) {
        return
    }

    # Step 2: Let user select Excel file
    $excelFile = Select-ExcelFile
    if (-not $excelFile) {
        return
    }
    
    # Open Excel
    $data = Open-ExcelPackage -Path $excelFile
    $sheet = $data.Workbook.Worksheets[1] # first sheet

    # Check if namess are in the Worker Hashtable
    $startingColumn = "A"
    $startingRow = 5
    $offset = 4

    $name = ""
    $position = ""

    while ($true){
        $cellAddress = "$startingColumn$startingRow"
        $cellValue = $sheet.Cells[$cellAddress].Text.Trim()

        if ($WorkerTable.ContainsKey($cellValue)) {
        $worker = $WorkerTable[$cellValue]
        $name = $worker.Name
        $position = $worker.Position
        Write-Host "Found: $($name) - $($position) at $cellAddress"

        # TODO: Add Shift Parsing Here

        $startingRow += $offset
    } elseif([string]::IsNullOrWhiteSpace($cellValue)){
        # Breaks when blank and done with names
        Write-Host "Blank at $cellAddress. Stopping loop."
        break
    }else {
        # Person not found in hashtable and skipped
        Write-Host "No match for '$cellValue' at $cellAddress, skipping."
        $startingRow += $offset
        }
    }
    
    # Clean Up
    Close-ExcelPackage $data
}

# TODO: to test column increment
function TestColumnIncrementMain{
    $string = Read-Host "Input a letter to increment"

    $result = Increment-SingleLetter -input $string

    Write-Host "Incremented '$string' to '$result'"
}

function HashTableCheckMain{
    # Step 1: Check ImportExcel module
    if (-not (Test-ImportExcelModule)) {
        return
    }

    # Step 2: Let user select Excel file
    $excelFile = Select-ExcelFile
    if (-not $excelFile) {
        return
    }
    
    # Open Excel
    $data = Open-ExcelPackage -Path $excelFile
    $sheet = $data.Workbook.Worksheets[1] # first sheet

    # Check if namess are in the Worker Hashtable
    $name = ""
    $startingColumn = "A"
    $startingRow = 5
    $offset = 4

    while ($true){
        $cellAddress = "$startingColumn$startingRow"
        $cellValue = $sheet.Cells[$cellAddress].Text.Trim()

        if ($WorkerTable.ContainsKey($cellValue)) {
        $worker = $WorkerTable[$cellValue]
        Write-Host "Found: $($worker.Name) - $($worker.Position) at $cellAddress"
        $startingRow += $offset
    } elseif([string]::IsNullOrWhiteSpace($cellValue)){
        Write-Host "Blank at $cellAddress. Stopping loop."
        break
    }else {
        Write-Host "No match for '$cellValue' at $cellAddress, skipping."
        $startingRow += $offset
        }
    }
    
    # Clean Up
    Close-ExcelPackage $data
}

function CellCheckMain {
    # Step 1: Check ImportExcel module
    if (-not (Test-ImportExcelModule)) {
        return
    }

    # Step 2: Let user select Excel file
    $excelFile = Select-ExcelFile
    if (-not $excelFile) {
        return
    }

    # Step 3: Ask for cell address
    $cellAddress = Read-Host "Enter the cell address (e.g., A1, B2, C5)"
    if (-not $cellAddress) {
        Write-Host "No cell entered. Exiting." -ForegroundColor Yellow
        return
    }

    try {
        # Step 4: Read the worksheet as raw data
        $data = Open-ExcelPackage -Path $excelFile
        $sheet = $data.Workbook.Worksheets[1]  # first sheet
        $value = $sheet.Cells[$cellAddress].Text  # get text of the cell

        if ($value) {
            Write-Host "Value in ${cellAddress}: $value" -ForegroundColor Cyan
        }
        else {
            Write-Host "No value found at ${cellAddress}" -ForegroundColor Red
        }

        # Cleanup
        Close-ExcelPackage $data
    }
    catch {
        Write-Error "Failed to read Excel file: $_"
    }
}

# To run
try {
    Main
} catch {
    Write-Error "Unexpected error: $($_.Exception.Message)"
}