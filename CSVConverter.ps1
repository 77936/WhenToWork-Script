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