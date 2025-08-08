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

# Done: Helper function for column incremention
function ColumnIncrementHelper{
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern("^[A-Z]$")]
        [string]$letter
    )

    $i = [int][char]$letter
    if ($i -lt [int][char]'Z'){
        $nextletter = $i + 1
        $nextChar = [char]$nextLetter
        return $nextChar
    } else {
        return "A"
    }
}

# Done: Calculate Paid Hours
# returns $paidHours
function PaidHourCalculator{
    param(
        [Parameter (Mandatory = $true)]
        [string]$StartTime,

        [Parameter (Mandatory = $true)]
        [string]$EndTime
    )

    try{
        $start = [DateTime]::Parse($StartTime)
        $end   = [DateTime]::Parse($EndTime)

        # Prob not used but handles overnight shifts
        if ($end -le $start){
            $end = $end.AddDays(1)
        }

        # Calculate total timespan
        $timespan = $end - $start
        $totalHours = $timespan.TotalHours
        $totalMinutes = $timespan.TotalMinutes

        # Extract whole hours and remaining minutes
        $wholeHours = [Math]::Floor($totalHours)
        $remainingMinutes = $totalMinutes % 60

        # Apply minute rounding scale
        $roundedMinuteHours = switch ($remainingMinutes) {
            {$_ -ge 1  -and $_ -le 6}   {0.1}
            {$_ -ge 7  -and $_ -le 12}  {0.2}
            {$_ -ge 13 -and $_ -le 18}  {0.3}
            {$_ -ge 19 -and $_ -le 24}  {0.4}
            {$_ -ge 25 -and $_ -le 30}  {0.5}
            {$_ -ge 31 -and $_ -le 36}  {0.6}
            {$_ -ge 37 -and $_ -le 42}  {0.7}
            {$_ -ge 43 -and $_ -le 48}  {0.8}
            {$_ -ge 49 -and $_ -le 54}  {0.9}
            {$_ -ge 55}                 {1.0}
            default                     {0.0}
        }

        # Calculate total paid hours before break deduction
        $paidHours = $wholeHours + $roundedMinuteHours

        # Subtract 30 minute break (0.5 Hours) if shift is over 6 hours
        if ($paidHours -gt 6){
            $paidHours -= 0.5
        }

        return $paidHours
    } catch{
        Write-Error "Error parsing time values. Please use format like '9:00am' or '5:30pm'. Error: $($_.Exception.Message)"
        return $null
    }
}

# Function to parse time shifts for schedule mode *ADD LOCATION CODES TO "Category" Header*
# Returns StartTime, EndTime, Category
function Parse-Time-Location {
    param(
        [string]$time
    )

    $category = $null
    
    # Return null for empty, or letters with no shift time, ex. OFF
    if ([string]::IsNullOrWhiteSpace($time) -or $time -match "^[a-zA-Z]+$") {
        return $null
    }
    
    # Split time & category if present
    if ($time -match '^(.+?)\s+([A-Za-z]+)$') {
        $cleanShift = $matches[1]
        $category   = $matches[2]
    } else {
        $cleanShift = $time
        $category   = "NB"
    }

    # Match time patterns like "9:30-4:30", "7am-3pm", "2-9pm", "9:30-4:30pm"
    $timePattern = '(\d{1,2}:?\d{0,2}\s*(?:am|pm)?)\s*-\s*(\d{1,2}:?\d{0,2})\s*(am|pm)?'
    
    if ($cleanShift -match $timePattern) {
        $startTime = $matches[1]
        $endTime   = $matches[2]
        $endPeriod = $matches[3]
        
        # Add missing colons for times without minutes
        if ($startTime -notmatch ':') { $startTime += ':00' }
        if ($endTime -notmatch ':')   { $endTime   += ':00' }
        
        # Handle AM/PM logic
        if ($endPeriod) {
            $endTime = $endTime + $endPeriod.ToLower()
            
            if ($endPeriod.ToLower() -eq 'pm' -and $cleanShift -notmatch 'am' -and $startTime -notmatch 'am|pm') {
                $startHour = [int]($startTime -split ':')[0]
                if ($startHour -ge 6 -and $startHour -le 11) {
                    $startTime += 'am'
                } else {
                    $startTime += 'pm'
                }
            }
        } else {
            # No AM/PM specified — educated guess
            $startHour = [int]($startTime -split ':')[0]
            $endHour   = [int]($endTime -split ':')[0]

            # Decide AM/PM for start
            if ($startHour -ge 6 -and $startHour -le 11) {
                $startTime += 'am'
            } else {
                $startTime += 'pm'
            }

            $startPeriod = $startTime.Substring($startTime.Length - 2)

            # Decide AM/PM for end
            if ($endHour -lt $startHour) {
                if ($startHour -eq 12 -and $startPeriod -eq 'pm') {
                    # Special case: noon shift, later end hour is still PM
                    $endTime += 'pm'
                }
                elseif ($startPeriod -eq 'am') {
                    $endTime += 'pm'
                } else {
                    $endTime += 'am'
                }
            }
            elseif ($endHour -eq 12) {
                # Special noon/midnight handling
                if ($startPeriod -eq 'am') {
                    $endTime += 'pm'  # noon
                } else {
                    $endTime += 'am'  # midnight
                }
            }
            else {
                $endTime += $startPeriod
            }
        }

        # Call PaidHoursCalculation
        $paidHours = PaidHourCalculator -StartTime $startTime -EndTime $endTime

        
        return @{
            StartTime = $startTime
            EndTime   = $endTime
            PaidHours = $paidHours
            Category  = $category
        }
    }
    
    Write-Warning "Could not parse shift: '$time'"
    return $null
}

# Function to parse through shift cell group
# Param $Worksheet, $StartRow, StartColumn
# Returns $cellsTexts (array of 4 cells)
function Parse-CellGroup {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,
        
        [Parameter(Mandatory = $true)]
        [int]$StartRow,
        
        [Parameter(Mandatory = $true)]
        [string]$StartColumn
    )
    
    $cellTexts = @()
    
    # Get text from 4 consecutive cells
    for ($i = 0; $i -lt 4; $i++) {
        $currentRow = $StartRow + $i
        $cellAddress = "$StartColumn$currentRow"
        $cellText = $Worksheet.Cells[$cellAddress].Text.Trim()
        $cellTexts += $cellText
    }
    
    return $cellTexts
}

# Add Biweekly Check
# Function to process shift data from cell group
# Param $Worksheet, $StartRow, $StartColumn
# Returns PSCustomObject with shift data
function Process-ShiftGroup {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,
        
        [Parameter(Mandatory = $true)]
        [int]$StartRow,
        
        [Parameter(Mandatory = $true)]
        [string]$StartColumn
    )
    
    # Get the 4 cell texts
    $cellTexts = Parse-CellGroup -Worksheet $Worksheet -StartRow $StartRow -StartColumn $StartColumn
    
    # Initialize result object
    $result = [PSCustomObject]@{
        Shift1 = $null
        Shift2 = $null
        Description = ""
    }
    
    # Process Cell 1 (always check for shift)
    if (-not [string]::IsNullOrWhiteSpace($cellTexts[0])) {
        $result.Shift1 = Parse-Time-Location -time $cellTexts[0]
    }
    
    # Process Cell 2
    if (-not [string]::IsNullOrWhiteSpace($cellTexts[1])) {
        if ($cellTexts[1] -match '^\d') {
            # Starts with number - it's a shift
            $result.Shift2 = Parse-Time-Location -time $cellTexts[1]
        } else {
            # Doesn't start with number - it's a description
            $result.Description = $cellTexts[1]
            return $result
        }
    }
    
    # Process Cell 3 (only if we got here - meaning cell 2 was a shift or empty)
    if (-not [string]::IsNullOrWhiteSpace($cellTexts[2])) {
        if ($cellTexts[2] -match '^\d') {
            # Starts with number - could be a third shift or paid hours
            # For now, treat as description since you mentioned checking if it doesn't start with number
        } else {
            # Doesn't start with number - add to description for both shifts
            $result.Description = $cellTexts[2]
        }
    }
    
    Process Cell 4 (paid hours)
    if (-not [string]::IsNullOrWhiteSpace($cellTexts[3]) -and $cellTexts[3] -match '^\d') {
        
        # Add Biweekly check
        $result.PaidHours = $cellTexts[3]
    }
    
    return $result
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
    $startingNameColumn = "A"
    $startingNameRow = 5
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

# Testers

function TestProcessShiftGroup {
    # Step 1: Check ImportExcel module
    if (-not (Test-ImportExcelModule)) {
        return
    }

    # Step 2: Let user select Excel file
    $excelFile = Select-ExcelFile
    if (-not $excelFile) {
        return
    }

    # Step 3: Ask for starting position
    $startColumn = Read-Host "Enter the starting column (e.g., C, D, E)"
    if (-not $startColumn) {
        Write-Host "No column entered. Using default 'C'" -ForegroundColor Yellow
        $startColumn = "C"
    }

    $startRowInput = Read-Host "Enter the starting row number (e.g., 5)"
    if (-not $startRowInput) {
        Write-Host "No row entered. Using default '5'" -ForegroundColor Yellow
        $startRow = 5
    } else {
        $startRow = [int]$startRowInput
    }

    try {
        # Step 4: Open Excel and get worksheet
        $data = Open-ExcelPackage -Path $excelFile
        $sheet = $data.Workbook.Worksheets[1]  # first sheet

        # Step 5: Process the shift group
        Write-Host "`nProcessing shift group starting at $startColumn$startRow..." -ForegroundColor Cyan
        
        # First show raw cell contents
        $cellTexts = Parse-CellGroup -Worksheet $sheet -StartRow $startRow -StartColumn $startColumn
        Write-Host "`n=== Raw Cell Contents ===" -ForegroundColor Yellow
        for ($i = 0; $i -lt $cellTexts.Count; $i++) {
            $cellAddress = "$startColumn$($startRow + $i)"
            Write-Host "$cellAddress : '$($cellTexts[$i])'"
        }
        
        # Then process the shift data
        $shiftResult = Process-ShiftGroup -Worksheet $sheet -StartRow $startRow -StartColumn $startColumn

        Write-Host "`n=== Processed Shift Data ===" -ForegroundColor Green
        
        if ($shiftResult.Shift1) {
            Write-Host "Shift 1:"
            Write-Host "  Start: $($shiftResult.Shift1.StartTime)"
            Write-Host "  End: $($shiftResult.Shift1.EndTime)"
            Write-Host "  Category: $($shiftResult.Shift1.Category)"
        } else {
            Write-Host "Shift 1: None"
        }
        
        if ($shiftResult.Shift2) {
            Write-Host "Shift 2:"
            Write-Host "  Start: $($shiftResult.Shift2.StartTime)"
            Write-Host "  End: $($shiftResult.Shift2.EndTime)"
            Write-Host "  Category: $($shiftResult.Shift2.Category)"
        } else {
            Write-Host "Shift 2: None"
        }
        
        Write-Host "Description: '$($shiftResult.Description)'"
        Write-Host "Paid Hours: '$($shiftResult.PaidHours)'"

        # Cleanup
        Close-ExcelPackage $data
    }
    catch {
        Write-Error "Failed to test shift group processing: $_"
    }
}

function TestCellGroupParse {
    # Step 1: Check ImportExcel module
    if (-not (Test-ImportExcelModule)) {
        return
    }

    # Step 2: Let user select Excel file
    $excelFile = Select-ExcelFile
    if (-not $excelFile) {
        return
    }

    # Step 3: Ask for starting position
    $startColumn = Read-Host "Enter the starting column (e.g., C, D, E)"
    if (-not $startColumn) {
        Write-Host "No column entered. Using default 'C'" -ForegroundColor Yellow
        $startColumn = "C"
    }

    $startRowInput = Read-Host "Enter the starting row number (e.g., 5)"
    if (-not $startRowInput) {
        Write-Host "No row entered. Using default '5'" -ForegroundColor Yellow
        $startRow = 5
    } else {
        $startRow = [int]$startRowInput
    }

    try {
        # Step 4: Open Excel and get worksheet
        $data = Open-ExcelPackage -Path $excelFile
        $sheet = $data.Workbook.Worksheets[1]  # first sheet

        # Step 5: Parse the cell group
        Write-Host "`nGetting cell texts starting at $startColumn$startRow..." -ForegroundColor Cyan
        $cellTexts = Parse-CellGroup -Worksheet $sheet -StartRow $startRow -StartColumn $startColumn

        Write-Host "`n=== Cell Texts ===" -ForegroundColor Green
        for ($i = 0; $i -lt $cellTexts.Count; $i++) {
            $cellAddress = "$startColumn$($startRow + $i)"
            Write-Host "$cellAddress : '$($cellTexts[$i])'"
        }

        # Cleanup
        Close-ExcelPackage $data
    }
    catch {
        Write-Error "Failed to test cell group parsing: $_"
    }
}

function TestTimeParse{
    $time = Read-Host "Input a time to parse"

    $result = Parse-Time-Location -time $time

    Write-Host "Start time = '$($result.StartTime)'"
    Write-Host "End time = '$($result.EndTime)'"
    Write-Host "Category = '$($result.Category)'"
    Write-Host "Paid Hours = '$($result.PaidHours)'"
}

function TestColumnIncrementMain{
    $string = Read-Host "Input a letter to increment"

    $result = ColumnIncrementHelper -letter $string

    Write-Host "Incremented '$string' to '$result'"
}

function TestHashTable{
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

function TestCellCheck {
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
    TestTimeParse
} catch {
    Write-Error "Unexpected error: $($_.Exception.Message)"
}