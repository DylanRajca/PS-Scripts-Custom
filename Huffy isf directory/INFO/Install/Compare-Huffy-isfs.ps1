# Synopsis - Compare Huffy daily loading report with ISF database and return new bookings.

# File paths

$huffyIsfDirectory = "$env:USERPROFILE\Desktop\Huffy isf directory"
$loadingReportPath = "$huffyIsfDirectory\*.xlsx"
$databasePath = "$huffyIsfDirectory\ISF database\Huffy ISF database.csv"
$databaseBackup = "$env:USERPROFILE\OneDrive - Green Worldwide Shipping, LLC\Huffy ISF database-backup.csv"
$newBookingsDirectory = "$huffyIsfDirectory\New booking reports"
$pastReports = "$huffyIsfDirectory\Past loading reports"

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
##### MAIN #####
#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function compareFiles {
    if ($null -ne $files) {

        # Initialize date variable.
        $date = Get-Date -UFormat %D 
        $date = [string]$date -replace "/", '-'
        
        # Iterate through loading reports.
        for ($i = 0; $i -lt $files.Count; $i++) {
            $fileName = $files[$i].BaseName
            $start = Read-Host -Prompt "Would you like to compare $fileName with Huffy ISF database? (y/n)"
            if ($start -eq 'y') {
    
                ### Conditional for 'Huffy daily loading report'
                if ($fileName -match "Huffy Daily Loading Report") {
            
                    # Open Excel and create open new bookings report
                    if ($i -eq 0) {
                        $XL = New-Object -comobject Excel.Application
                        $XL.visible = $false
                        $newBookings = $XL.Workbooks.Add()
                        $newBookingsSheet = $newBookings.worksheets.item(1)

                        # Initialize database.
                        Write-Host "Initializing database..."
                        $database = $XL.Workbooks.Open($databasePath)
                        $databaseSheet = $database.worksheets.item(1)
                        $databaseArray = @()
                        
                        # Save index of last row in database 
                        [int]$lastRowDatabase = ($databaseSheet.UsedRange.Rows.count + 1) - 1
                        for ($n = 1; $n -le $lastRowDatabase; $n++) {
                            $value = ($databaseSheet.Rows.item($n).cells.item(1).text)
                            $value = $value.Trim()
                            $databaseArray += , $value
                        }
                        $databaseArray = $databaseArray | Select-Object -Unique
                    }
        
                    # Assign $book and $sheet variables to the current workbook.
                    $book = $XL.Workbooks.Open($files[$i].FullName)
                    $sheet = $book.worksheets.item(1)

                    # Save index of last column/row used
                    [int]$lastRowValue = ($sheet.UsedRange.Rows.count + 1) - 1
                    [int]$lastColValue = ($sheet.UsedRange.columns.count + 1) - 1

                    # Save header/BL column number
                    $blColumn = $null
                    $loading_report_header = @()
                    for ($n = 1; $n -le $lastColValue; $n++) {
                        $value = ($sheet.Rows.item(1).cells.item($n).text)
                        $loading_report_header += , $value
                        if ($value -eq 'CBL Number') {
                            $blColumn = $n
                        } 
                    }

                    # Conditional to check if BL column exists
                    if ($null -ne $blColumn) {
                
                        # Create BL array (listOne)
                        Write-Host "Checking $fileName for new bookings..."
                        $listOne = @()
                        for ($n = 1; $n -le $lastRowValue; $n++) {
                            $value = ($sheet.Rows.item($n).cells.item($blColumn).text)
                            $value = $value.Trim()
                            $listOne += , $value
                        }
                        $listOne | Group-Object | % { $hash = @{} } { $hash[$_.Name] = $_.Count } { $hash } | Out-Null
                        $listOne = $listOne | Select-Object -Unique

                        ### Report/database comparison
                        try {
                            $difference = Compare-Object $listOne $databaseArray | where { $_.sideindicator -eq "<=" } | % { $_.inputobject }
                            $differenceCount = $difference.Count
                            if ($differenceCount -eq 1) {
                                Write-Host "$differenceCount new booking found."
                            }
                            else {
                                Write-Host "$differenceCount new bookings found."
                            }
                        }
                        catch {
                            Write-Host "ERROR: There was an error comparing files. Please exit and try again."
                        }

                        ### Iterating through new bookings and adding data to new bookings report/database
                        Write-Host "Generating new bookings report and updating database..."
                        for ($f = 0; $f -lt $difference.Count; $f++) {
                            $value = $difference[$f]
                            $count = $hash[$value]
                            $rowArray = @()

                            # Capturing cell address of BL in loading report and generating row number.
                            $getName = $sheet.Range("C$blColumn").find($difference[$f])
                            $cellAddress = $getName.Address($false, $false)
                            $row = -join $cellAddress[1..($cellAddress.Length - 1)]

                            
                            # Highlight cell address
                            $getName.Interior.colorindex = 6

                            # Capturing row data from loading report.
                            for ($n = 1; $n -le $lastColValue; $n++) {
                                $rowValue = ($sheet.Rows.item($row).cells.item($n).Text)
                                $rowArray += , $rowValue
                            }

                            # Adding header to new bookings report.
                            for ($n = 0; $n -lt $loading_report_header.Count; $n++) {
                                $newBookingsSheet.cells.item(1, ($n + 1)) = $loading_report_header[$n]
                            }

                            # Adding row data to new bookings report.
                            for ($r = 0; $r -lt $count; $r++) {
                                for ($n = 0; $n -lt $rowArray.Count; $n++) {
                                    if ($f -eq 0) {

                                        # Setting first row for bookings under header.
                                        $firstRow = 1
                                    }
                                    $newBookingsSheet.cells.item(($firstRow + 1), ($n + 1)) = $rowArray[$n]
                                }
                                $firstRow = $firstRow + 1
                            }

                            # Add BL to database
                            $lastRowDatabase = ($lastRowDatabase + 1)
                            $databaseSheet.cells.item($lastRowDatabase, 1) = $difference[$f]
                            $databaseSheet.cells.item($lastRowDatabase, 3) = $date

                            # Add $difference array list to txt file in new bookings directory
                            "$differenceCount new bookings ($date)`n", $difference | Out-File "$newBookingsDirectory\New -$fileName.txt"
                        }
                    }
                    else {
                        Write-Host "ERROR: Please make sure 'CBL Number' column exists on daily loading report."
                    }
                    Write-Host "Complete."

                    # Save & close workbook
                    $book.Save()
                    $book.close($true)
            
                    # Move loading report to past bookings folder
                    Move-Item $files[0] -Destination $pastReports

                    # Save new bookings/database/database back up & quit excel
                    $newBookings.SaveAs("$newBookingsDirectory\New -$fileName.xlsx")
                    $database.SaveAs($databasePath)
                    $database.SaveAs($databaseBackup)
                    $XL.Quit()

                }
        
                ### Conditional if daily loading report does not exist
                else {
                    Write-Host "ERROR: Please make sure a 'Huffy Daily Loading Report' exists in $huffyIsfDirectory."
                }
            }
            elseif ($start -ne "y" -and $start -ne "n") {
                $i = ($i - 1)
            }
        }
    }
    else {
        Write-Host "ERROR: Please upload a Huffy loading report to $huffyIsfDirectory."
        $continue = Read-Host -Prompt "Would you like to try again? (y/n)"
        if ($continue -eq "y") {
            compareFiles
        }
    }
}

# Grab loading reports in directory.
try {
    $files = Get-ChildItem $loadingReportPath
}
catch {}

compareFiles

