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

# Done: Calculate Column Letter to Dates for Shift
# Function to read date range from A2 and assign dates to columns starting from C
# Returns hashtable with column letters as keys and dates as values
function Assign-ColumnDates {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet
    )
    
    # Read the date range from cell A2
    $dateRangeText = $Worksheet.Cells["A2"].Text.Trim()
    
    if ([string]::IsNullOrWhiteSpace($dateRangeText)) {
        Write-Error "No date range found in cell A2"
        return $null
    }
    
    Write-Host "Found date range: $dateRangeText" -ForegroundColor Cyan
    
    # Parse the date range (format: "August 2, 2025 - August 15, 2025")
    $datePattern = '^(.+?)\s*-\s*(.+)$'
    
    if ($dateRangeText -match $datePattern) {
        $startDateText = $matches[1].Trim()
        $endDateText = $matches[2].Trim()
        
        try {
            # Parse start and end dates
            $startDate = [DateTime]::Parse($startDateText)
            $endDate = [DateTime]::Parse($endDateText)
            
            Write-Host "Start Date: $($startDate.ToString('MM/dd/yyyy'))" -ForegroundColor Green
            Write-Host "End Date: $($endDate.ToString('MM/dd/yyyy'))" -ForegroundColor Green
            
            # Calculate number of days in the range (should be 14 for 2-week schedule)
            $totalDays = ($endDate - $startDate).Days + 1
            Write-Host "Total days in range: $totalDays" -ForegroundColor Yellow
            
            # Validate it's a 2-week schedule
            if ($totalDays -ne 14) {
                Write-Warning "Expected 14 days for 2-week schedule, but found $totalDays days"
            }
            
            # Create hashtable to store column-date mappings
            $columnDateMap = @{}
            
            # Start from column C
            $currentColumn = "C"
            $currentDate = $startDate
            
            # Assign dates to columns (C through P for 14 days)
            for ($day = 0; $day -lt $totalDays -and $day -lt 14; $day++) {
                $dateString = $currentDate.ToString("MM/dd/yyyy")
                $columnDateMap[$currentColumn] = $dateString
                
                Write-Host "Column $currentColumn = $dateString" -ForegroundColor White
                
                # Move to next date and column
                $currentDate = $currentDate.AddDays(1)
                
                # Only increment column if we're not on the last day
                if ($day -lt ($totalDays - 1) -and $day -lt 13) {
                    $currentColumn = ColumnIncrementHelper -letter $currentColumn
                }
            }
            
            Write-Host "Successfully mapped $($columnDateMap.Count) columns to dates" -ForegroundColor Green
            return $columnDateMap
            
        } catch {
            Write-Error "Failed to parse dates: $($_.Exception.Message)"
            Write-Host "Expected format: 'August 2, 2025 - August 15, 2025'" -ForegroundColor Yellow
            return $null
        }
        
    } else {
        Write-Error "Date range format not recognized. Expected format: 'August 2, 2025 - August 15, 2025'"
        return $null
    }
}

# Done: Function to convert column letter to date using the column-date hashtable
# Param: Column letter (string), ColumnDateMap (hashtable from Assign-ColumnDates)
# Returns: Date string in MM/dd/yyyy format, or null if column not found
function Get-DateFromColumn {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ColumnLetter,
        
        [Parameter(Mandatory = $true)]
        $ColumnDateMap
    )
    
    # Convert to uppercase to ensure consistency
    $ColumnLetter = $ColumnLetter.ToUpper().Trim()
    
    # Try direct lookup first
    if ($ColumnDateMap.ContainsKey($ColumnLetter)) {
        return $ColumnDateMap[$ColumnLetter]
    }
    
    # If direct lookup fails, try iterating through keys to find match
    foreach ($key in $ColumnDateMap.Keys) {
        $cleanKey = $key.ToString().Trim().ToUpper()
        if ($cleanKey -eq $ColumnLetter) {
            return $ColumnDateMap[$key]
        }
    }
    
    # If still not found, try with string comparison
    foreach ($key in $ColumnDateMap.Keys) {
        if ([string]$key -eq $ColumnLetter) {
            return $ColumnDateMap[$key]
        }
    }
    
    Write-Warning "Column '$ColumnLetter' not found in date mapping"
    return $null
}

# Done: Function to parse time shifts for schedule mode *ADD LOCATION CODES TO "Category" Header*
# Returns StartTime, EndTime, Category
function Parse-Time-Location {
    param(
        [string]$time
    )

    # Allows for date addition to the shift in a later function
    $date = $null

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
        $startTime = $matches[1].Trim()
        $endTime   = $matches[2].Trim()
        $endPeriod = $matches[3]
        
        # Handle AM/PM for start time first (before adding minutes)
        $startHasAmPm = $startTime -match '(am|pm)$'
        $startAmPm = ""
        if ($startHasAmPm) {
            $startAmPm = $startTime.Substring($startTime.Length - 2)
            $startTime = $startTime.Substring(0, $startTime.Length - 2)
        }
        
        # Add missing colons for times without minutes
        if ($startTime -notmatch ':') { 
            $startTime = $startTime + ':00'
        }
        if ($endTime -notmatch ':') { 
            $endTime = $endTime + ':00'
        }
        
        # Re-add AM/PM to start time after fixing format
        if ($startHasAmPm) {
            $startTime = $startTime + $startAmPm
        }
        
        # Handle AM/PM logic for end time
        if ($endPeriod) {
            $endTime = $endTime + $endPeriod.ToLower()
            
            # If end has AM/PM and start doesn't, infer start time AM/PM
            if (-not $startHasAmPm) {
                $startHour = [int]($startTime -split ':')[0]
                if ($endPeriod.ToLower() -eq 'pm' -and $startHour -ge 6 -and $startHour -le 11) {
                    $startTime += 'am'
                } elseif ($endPeriod.ToLower() -eq 'pm') {
                    $startTime += 'pm'
                } else {
                    # End is AM
                    $startTime += 'am'
                }
            }
        } else {
            # No AM/PM specified for end time — educated guess
            $startHour = [int]($startTime -split ':')[0]
            $endHour   = [int]($endTime -split ':')[0]

            # If start time doesn't have AM/PM, decide it first
            if (-not $startHasAmPm) {
                if ($startHour -ge 6 -and $startHour -le 11) {
                    $startTime += 'am'
                } else {
                    $startTime += 'pm'
                }
            }

            $startPeriod = $startTime.Substring($startTime.Length - 2)

            # Decide AM/PM for end based on start
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
            Date = $date
            StartTime = $startTime
            EndTime   = $endTime
            PaidHours = $paidHours
            Category  = $category
        }
    }
    
    Write-Warning "Could not parse shift: '$time'"
    return $null
}

# Done: Function to parse through shift cell group
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

# Done: Function to process shift data from cell group
# Param $Worksheet, $StartRow, $StartColumn
# Returns PSCustomObject with shift data
function Process-ShiftGroup {
    param(
        [Parameter(Mandatory = $true)]
        $Worksheet,
        [Parameter(Mandatory = $true)]
        [hashtable]$ColumnDateMap,
        
        [Parameter(Mandatory = $true)]
        [int]$StartRow,
        
        [Parameter(Mandatory = $true)]
        [string]$StartColumn
    )
    
    $date = Get-DateFromColumn -ColumnLetter $StartColumn -ColumnDateMap $ColumnDateMap
    # Get the 4 cell texts
    $cellTexts = Parse-CellGroup -Worksheet $Worksheet -StartRow $StartRow -StartColumn $StartColumn
    
    # Initialize result object
    $result = [PSCustomObject]@{
        Shift1 = $null
        Shift2 = $null
        Description = ""
        Date = $date
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
            
            # Check Cell 4 (paid hours) for single shift scenario
            if (-not [string]::IsNullOrWhiteSpace($cellTexts[3]) -and $cellTexts[3] -match '^\d+\.?\d*$') {
                $cell4Hours = [double]$cellTexts[3]
                
                # If we only have one shift, check pattern
                if ($result.Shift1 -and -not $result.Shift2) {
                    $calculatedHours = $result.Shift1.PaidHours
                    $difference = [Math]::Abs($calculatedHours - $cell4Hours)
                    
                    if ($difference -le 0.1) {
                        # Matches - keep calculated hours
                        Write-Host "Cell 4 hours match calculated hours, keeping calculated value" -ForegroundColor Green
                    } elseif ([Math]::Abs($difference - 0.5) -le 0.1) {
                        # Difference is 0.5 hours (break) - use Cell 4 hours
                        Write-Host "Cell 4 hours differ by 0.5 (break time), using Cell 4 value: $cell4Hours" -ForegroundColor Yellow
                        $result.Shift1.PaidHours = $cell4Hours
                    } else {
                        # Invalid difference - clear shift due to mistaken parsing
                        Write-Host "Cell 4 hours ($cell4Hours) don't match expected pattern (diff: $difference), clearing Shift1 due to mistaken parsing" -ForegroundColor Red
                        $result.Shift1 = $null
                    }
                }
            }
            
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
    
    # Process Cell 4 (paid hours validation and correction)
    if (-not [string]::IsNullOrWhiteSpace($cellTexts[3]) -and $cellTexts[3] -match '^\d+\.?\d*$') {
        $cell4Hours = [double]$cellTexts[3]
        
        # Calculate expected total paid hours from shifts
        $expectedTotal = 0
        if ($result.Shift1) { $expectedTotal += $result.Shift1.PaidHours }
        if ($result.Shift2) { $expectedTotal += $result.Shift2.PaidHours }
        
        # Check scenarios and update accordingly
        if ($result.Shift1 -and $result.Shift2) {
            # Two shifts: Check if Cell 4 matches combined total
            $difference = [Math]::Abs($expectedTotal - $cell4Hours)
            
            if ($difference -le 0.1) {
                # Matches - keep calculated hours
                Write-Host "Cell 4 hours match calculated total, keeping calculated values" -ForegroundColor Green
            } elseif ([Math]::Abs($difference - 0.5) -le 0.1) {
                # Difference is 0.5 hours (break) - add 0.5 to the longer shift
                Write-Host "Cell 4 hours differ by 0.5 (break time), adding 0.5 to longer shift" -ForegroundColor Yellow
                
                # Find which shift has more paid hours
                if ($result.Shift1.PaidHours -gt $result.Shift2.PaidHours) {
                    $result.Shift1.PaidHours += 0.5
                } elseif ($result.Shift2.PaidHours -gt $result.Shift1.PaidHours) {
                    $result.Shift2.PaidHours += 0.5
                } else {
                    # Equal hours - add to Shift1 by default
                    $result.Shift1.PaidHours += 0.5
                }
            } else {
                # Invalid difference - clear shifts
                Write-Host "Cell 4 hours ($cell4Hours) don't match expected pattern, clearing shifts" -ForegroundColor Red
                $result.Shift1 = $null
                $result.Shift2 = $null
                $result.Date = $null
            }
        } elseif ($result.Shift1 -and -not $result.Shift2) {
            # Single shift: Check difference patterns
            $calculatedHours = $result.Shift1.PaidHours
            $difference = [Math]::Abs($calculatedHours - $cell4Hours)
            
            if ($difference -le 0.1) {
                # Matches - keep calculated hours
                Write-Host "Cell 4 hours match calculated hours, keeping calculated value" -ForegroundColor Green
            } elseif ([Math]::Abs($difference - 0.5) -le 0.1) {
                # Difference is 0.5 hours (break) - use Cell 4 hours
                Write-Host "Cell 4 hours differ by 0.5 (break time), using Cell 4 value: $cell4Hours" -ForegroundColor Yellow
                $result.Shift1.PaidHours = $cell4Hours
            } else {
                # Invalid difference - clear shift due to mistaken parsing
                Write-Host "Cell 4 hours ($cell4Hours) don't match expected pattern (diff: $difference), clearing Shift1 due to mistaken parsing" -ForegroundColor Red
                $result.Shift1 = $null
                $result.Date = $null
            }
        }
        
        # Store the Cell 4 value for reference
        $result | Add-Member -NotePropertyName "Cell4Hours" -NotePropertyValue $cell4Hours
    }
    
    return $result
}

# TODO: Consolidate Shift Data
# Create a CSV row that fits the headers
function CreateCSVRow {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Worker,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$ShiftData,
        
        [Parameter(Mandatory = $false)]
        [int]$ShiftNumber = 1  # 1 for Shift1, 2 for Shift2
    )
    
    # Determine which shift to process
    $shift = $null
    if ($ShiftNumber -eq 1 -and $ShiftData.Shift1) {
        $shift = $ShiftData.Shift1
    } elseif ($ShiftNumber -eq 2 -and $ShiftData.Shift2) {
        $shift = $ShiftData.Shift2
    } else {
        # No valid shift found
        return $null
    }
    
    # Create the CSV row object
    $csvRow = [PSCustomObject]@{
        "Employee Name" = $Worker.Name
        "Position Name" = $Worker.Position
        "Date" = $ShiftData.Date
        "Start Time" = $shift.StartTime
        "End Time" = $shift.EndTime
        "Duration" = $shift.PaidHours
        "Category" = $shift.Category
        "Shift Description" = $ShiftData.Description
    }
    
    return $csvRow
}

# Helper function to create all CSV rows from shift data
function CreateAllCSVRows {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Worker,
        
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$ShiftData
    )
    
    $csvRows = @()
    
    # Create row for Shift1 if it exists
    if ($ShiftData.Shift1) {
        $row1 = CreateCSVRow -Worker $Worker -ShiftData $ShiftData -ShiftNumber 1
        if ($row1) {
            $csvRows += $row1
        }
    }
    
    # Create row for Shift2 if it exists
    if ($ShiftData.Shift2) {
        $row2 = CreateCSVRow -Worker $Worker -ShiftData $ShiftData -ShiftNumber 2
        if ($row2) {
            $csvRows += $row2
        }
    }
    
    return $csvRows
}






# Main function & Tester Main Functions

 # Main function - Rewritten with better debugging and robust processing
function Main {
    # Step 1: Check ImportExcel module
    if (-not (Test-ImportExcelModule)) {
        return
    }

    # Step 2: Let user select Excel file
    $excelFile = Select-ExcelFile
    if (-not $excelFile) {
        return
    }
    
    try {
        # Step 3: Open Excel
        $data = Open-ExcelPackage -Path $excelFile
        $sheet = $data.Workbook.Worksheets[1] # first sheet

        # Step 4: Get column date mapping
        Write-Host "Getting column date mapping..." -ForegroundColor Cyan
        $columnDateMap = Assign-ColumnDates -Worksheet $sheet
        if (-not $columnDateMap) {
            Write-Error "Failed to get column date mapping"
            Close-ExcelPackage $data
            return
        }

        # Step 5: Create an array to hold all CSV data
        $allCsvData = @()

        # Step 6: Define the column sequence for 14 days (C through P)
        $dayColumns = @("C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P")

        # Step 7: Process each worker
        $startingNameColumn = "A"
        $startingNameRow = 5
        $offset = 4
        $workerCount = 0

        Write-Host "Starting worker processing..." -ForegroundColor Cyan

        while ($true) {
            $cellAddress = "$startingNameColumn$startingNameRow"
            $cellValue = $sheet.Cells[$cellAddress].Text.Trim()

            Write-Host "`nChecking cell $cellAddress : '$cellValue'" -ForegroundColor White

            if ($WorkerTable.ContainsKey($cellValue)) {
                # Found a worker in our hashtable
                $worker = $WorkerTable[$cellValue]
                $workerCount++
                Write-Host "[$workerCount] Processing: $($worker.Name) - $($worker.Position) at row $startingNameRow" -ForegroundColor Green

                $shiftsFoundForWorker = 0

                # Process shifts for each of the 14 days
                for ($dayIndex = 0; $dayIndex -lt $dayColumns.Length; $dayIndex++) {
                    $currentColumn = $dayColumns[$dayIndex]
                    $currentDate = $columnDateMap[$currentColumn]
                    
                    Write-Host "  Day $($dayIndex + 1): Column $currentColumn ($currentDate)" -ForegroundColor Yellow
                    
                    # Process the 4-cell group for this worker and day
                    try {
                        $shiftData = Process-ShiftGroup -Worksheet $sheet -ColumnDateMap $columnDateMap -StartRow $startingNameRow -StartColumn $currentColumn
                        
                        # Create CSV rows for any shifts found
                        if ($shiftData -and ($shiftData.Shift1 -or $shiftData.Shift2)) {
                            $csvRows = CreateAllCSVRows -Worker $worker -ShiftData $shiftData
                            
                            if ($csvRows -and $csvRows.Count -gt 0) {
                                $allCsvData += $csvRows
                                $shiftsFoundForWorker += $csvRows.Count
                                Write-Host "    ✓ Added $($csvRows.Count) shift(s) for $currentDate" -ForegroundColor Green
                                
                                # Display shift details for confirmation
                                foreach ($row in $csvRows) {
                                    Write-Host "      - $($row.'Start Time') to $($row.'End Time') ($($row.Duration) hrs) [$($row.Category)]" -ForegroundColor Cyan
                                }
                            }
                        } else {
                            Write-Host "    - No shifts found for $currentDate" -ForegroundColor DarkGray
                        }
                    } catch {
                        Write-Host "    ⚠ Error processing day $($dayIndex + 1): $($_.Exception.Message)" -ForegroundColor Red
                    }
                }

                Write-Host "  Total shifts found for $($worker.Name): $shiftsFoundForWorker" -ForegroundColor Magenta
                $startingNameRow += $offset

            } elseif ([string]::IsNullOrWhiteSpace($cellValue)) {
                # Blank cell - we've reached the end of workers
                Write-Host "✓ Blank cell at $cellAddress. Finished processing all workers." -ForegroundColor Cyan
                Write-Host "Total workers processed: $workerCount" -ForegroundColor Green
                break

            } else {
                # Person not found in hashtable - skip this row
                Write-Host "⚠ No match for '$cellValue' at $cellAddress, skipping to next row." -ForegroundColor Yellow
                $startingNameRow += $offset
                
                # Safety check to prevent infinite loop
                if ($startingNameRow > 200) {
                    Write-Host "⚠ Reached row 200, stopping to prevent infinite loop" -ForegroundColor Red
                    break
                }
            }
        }

        # Step 8: Export to CSV
        Write-Host "`n=== PROCESSING COMPLETE ===" -ForegroundColor Magenta
        Write-Host "Total CSV rows collected: $($allCsvData.Count)" -ForegroundColor White

        if ($allCsvData.Count -gt 0) {
            # Create output filename
            $outputPath = [System.IO.Path]::ChangeExtension($excelFile, ".csv")
            
            # Export all data to CSV
            $allCsvData | Export-Csv -Path $outputPath -NoTypeInformation
            
            # Display summary
            Write-Host "`n=== EXPORT SUMMARY ===" -ForegroundColor Magenta
            Write-Host "✓ CSV exported to: $outputPath" -ForegroundColor Green
            Write-Host "✓ Total shifts exported: $($allCsvData.Count)" -ForegroundColor Cyan
            
            # Show breakdown by worker
            $workerSummary = $allCsvData | Group-Object "Employee Name"
            Write-Host "`nShifts per worker:" -ForegroundColor Yellow
            foreach ($worker in $workerSummary) {
                Write-Host "  $($worker.Name): $($worker.Count) shifts" -ForegroundColor White
            }
            
            # Show date range
            $dates = $allCsvData | Select-Object -ExpandProperty "Date" | Sort-Object -Unique
            if ($dates.Count -gt 0) {
                Write-Host "`nDate range: $($dates[0]) to $($dates[-1])" -ForegroundColor Yellow
            }

            # Show sample of first few rows
            Write-Host "`nFirst 3 exported rows:" -ForegroundColor Yellow
            $allCsvData | Select-Object -First 3 | Format-Table -AutoSize

        } else {
            Write-Host "`n⚠ No shift data found to export" -ForegroundColor Yellow
            Write-Host "Debugging information:" -ForegroundColor White
            Write-Host "  - Workers processed: $workerCount" -ForegroundColor Gray
            Write-Host "  - Please check:" -ForegroundColor Gray
            Write-Host "    • Worker names match the hashtable entries exactly" -ForegroundColor Gray
            Write-Host "    • Shift data is in the expected format" -ForegroundColor Gray
            Write-Host "    • Date range is properly formatted in cell A2" -ForegroundColor Gray
            
            # Show what workers were found
            if ($workerCount -eq 0) {
                Write-Host "`nTrying to debug worker detection..." -ForegroundColor Yellow
                $testRow = 5
                for ($i = 0; $i -lt 10; $i++) {
                    $testCell = "A$testRow"
                    $testValue = $sheet.Cells[$testCell].Text.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($testValue)) {
                        $isInTable = $WorkerTable.ContainsKey($testValue)
                        Write-Host "  Row $testRow ($testCell): '$testValue' - In table: $isInTable" -ForegroundColor Gray
                    }
                    $testRow += 4
                }
            }
        }

        # Step 9: Clean up
        Close-ExcelPackage $data

    } catch {
        Write-Error "Error in Main function: $($_.Exception.Message)"
        Write-Host "Stack trace: $($_.ScriptStackTrace)" -ForegroundColor Red
        
        # Make sure to clean up even if there's an error
        if ($data) {
            Close-ExcelPackage $data
        }
    }

    Write-Host "`nProcessing complete!" -ForegroundColor Green
}

# Testers

# Test function for Assign-ColumnDates
function TestAssignColumnDates {
    # Step 1: Check ImportExcel module
    if (-not (Test-ImportExcelModule)) {
        return
    }

    # Step 2: Let user select Excel file
    $excelFile = Select-ExcelFile
    if (-not $excelFile) {
        return
    }

    try {
        # Step 3: Open Excel and get worksheet
        $data = Open-ExcelPackage -Path $excelFile
        $sheet = $data.Workbook.Worksheets[1]  # first sheet

        # Step 4: Test the column date assignment
        Write-Host "`nTesting column date assignment..." -ForegroundColor Cyan
        $columnDates = Assign-ColumnDates -Worksheet $sheet

        if ($columnDates) {
            Write-Host "`n=== Column Date Mapping ===" -ForegroundColor Magenta
            foreach ($column in ($columnDates.Keys | Sort-Object)) {
                Write-Host "$column : $($columnDates[$column])"
            }
            
            # Show example of how to use the mapping
            Write-Host "`n=== Usage Example ===" -ForegroundColor Cyan
            Write-Host "To get the date for column D: $($columnDates['D'])"
            Write-Host "To check if column E exists: $($columnDates.ContainsKey('E'))"
        }

        # Cleanup
        Close-ExcelPackage $data
    }
    catch {
        Write-Error "Failed to test column date assignment: $_"
    }
}

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
        
        $ColumnDateMap = Assign-ColumnDates -Worksheet $sheet 
        Write-Host "$startColumn : '$($startColumn)'"

        # Then process the shift data
        $shiftResult = Process-ShiftGroup -Worksheet $sheet -ColumnDateMap $ColumnDateMap -StartRow $startRow -StartColumn $startColumn

        Write-Host "`n=== Processed Shift Data ===" -ForegroundColor Green
        
        if ($shiftResult.Shift1) {
            Write-Host "Shift 1:"
            Write-Host "  Start: $($shiftResult.Shift1.StartTime)"
            Write-Host "  End: $($shiftResult.Shift1.EndTime)"
            Write-Host "  Category: $($shiftResult.Shift1.Category)"
            Write-Host "  Paid Hours: $($shiftResult.Shift1.PaidHours)"
        } else {
            Write-Host "Shift 1: None"
        }
        
        if ($shiftResult.Shift2) {
            Write-Host "Shift 2:"
            Write-Host "  Start: $($shiftResult.Shift2.StartTime)"
            Write-Host "  End: $($shiftResult.Shift2.EndTime)"
            Write-Host "  Category: $($shiftResult.Shift2.Category)"
            Write-Host "  Paid Hours: $($shiftResult.Shift2.PaidHours)"
        } else {
            Write-Host "Shift 2: None"
        }
        
        Write-Host "Date: '$($shiftResult.Date)'"
        Write-Host "Description: '$($shiftResult.Description)'"

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
    Main
} catch {
    Write-Error "Unexpected error: $($_.Exception.Message)"
}