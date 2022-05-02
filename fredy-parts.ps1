<#
.Synopsis
Parse input.xlsx data into invoice line summary, then output to output.txt.
Author: Dylan Rajca, 02/01/22
#>

function returnExcelValues($inputPath) {
    Write-host "Extracting data from input.xlsx..."
    $XL = New-Object -comobject Excel.Application
    $XL.visible = $false
    
    # Assign $book and $sheet variables to input.sv
    $book = $XL.Workbooks.Open($inputPath)
    $sheet = $book.worksheets.item(1)

    ## -- CI DATA -- ##
    $ciItemNums = $sheet.UsedRange.Rows.Columns[1].Value2 | Where-Object { $_ -ne "CI-ITEM-NO" -and $null -ne $_ }
    $Ci_Hts = $sheet.UsedRange.Rows.Columns[3].Value2 | Where-Object { $_ -ne "CI-HTS/SCH.B" -and $null -ne $_ }
    $cost = $sheet.UsedRange.Rows.Columns[5].Value2 | Where-Object { $_ -ne "Cost" -and $null -ne $_ }

    ## -- PL DATA -- ##
    $plItemNums = $sheet.UsedRange.Rows.Columns[7].Value2 | Where-Object { $_ -ne "BL-ITEM-NO" -and $null -ne $_ }
    $weight = $sheet.UsedRange.Rows.Columns[9].Value2 | Where-Object { $_ -ne "WEIGHT" -and $null -ne $_ }


    # Close excel and return data
    $book.close($true)
    $XL.Quit()

    $excelData = [ordered]@{
        "Ci_Items" = $ciItemNums;
        "Ci_Hts"   = $Ci_Hts
        "Cost"     = $cost
        "Pl_Items" = $plItemNums
        "Weight"   = $weight
    }

    return $excelData
}

function parseData ($data) {

    #--- Validation Check ---#
    $validationStatus = validateData $data

    if ($validationStatus.Status -eq "valid") {

        Write-Host "Parsing CIV data..."

        #--- CIV ---#

        $htsValuesHash = [ordered]@{}

        for ($i = 0; $i -lt $data.Ci_Hts.Count; $i++) {
            $htsCode = $data.Ci_Hts[$i]
            $indexCost = [float]$data.Cost[$i]
            $indexCost = [math]::Round($indexCost, 2)
            $indexItem = $data.Ci_Items[$i]

            if ($htsValuesHash.$htsCode) {
                $htsValuesHash.$htsCode.Cost += $indexCost
                $htsValuesHash.$htsCode.Items += "$indexItem"
            }
            else {

                # Initialize HTS object values
                $htsValuesHash.$htsCode = [ordered]@{}
                $htsValuesHash.$htsCode.Cost += $indexCost
                $htsValuesHash.$htsCode.Items = @() 
                $htsValuesHash.$htsCode.Items += "$indexItem" 
                $htsValuesHash.$htsCode.Weight = 0.00
            }
        }

        #--- PL/BL ---#
        # PL Items loop
        for ($i = 0; $i -lt $data.Pl_Items.Count; $i++) {
            $plWeight = [float]$data.Weight[$i]
            $plWeight = [math]::Round($plWeight, 2)

            $plItem = $data.Pl_Items[$i]

            # $htsValueHash (HTS = (Cost = x); (Items = X)) loop
            foreach ($hash in $htsValuesHash.GetEnumerator()) {
                $htsKey = $hash.Key
                $htsItems = $hash.Value[1]


                # htsKey.Items loop
                $parentLoopBreak = $false
                for ($j = 0; $j -lt $htsItems.Count; $j++) {
                
                    if ($plItem -eq $htsItems[$j]) {            
                        $htsValuesHash.$htsKey.Weight += $plWeight
                        $parentLoopBreak = $true
                        break
                    }
                }

                # If $plItem is matched to $htsItems, break out of $htsValueHash loop
                if ($true -eq $parentLoopBreak) {
                    break          
                }
            }
        }
        return $htsValuesHash
    }

    else {
        return $validationStatus
    }
}

function validateData ($data) {
    Write-Host "Validating input..."

    $validationStatus = [ordered]@{
        "Status" = '';
    }

    # CI-Item-No duplicate check
    $ciItemGroup = $data.Ci_Items | Group-Object | Sort-Object -Property Count -Descending
    $ciItemGroup | ForEach-Object {
        if ($_.Count -gt 1) {
            $validationStatus.Duplicate_Items += , $_.Name
        }
    }
    
    if ($null -ne $validationStatus.Duplicate_Items) {
        Write-Host "Warning: Duplicate item numbers were found under CI-ITEM-NO column (check if HTS codes match)..." 
    }

    # Compare plTarrifs with ciTariffs and add any discrepancies to report.
    $missingCiItems = Compare-Object $data.Ci_Items $data.Pl_Items | Where-Object { $_.sideindicator -eq "<=" } | ForEach-Object { $_.inputobject }
    $missingPlItems = Compare-Object $data.Pl_Items $data.Ci_Items | Where-Object { $_.sideindicator -eq "<=" } | ForEach-Object { $_.inputobject }

    if ($null -ne $missingCiItems) {
        $validationStatus.Status = "not-valid"
        $validationStatus.Missing_Ci = $missingCiItems;
    }

    if ($null -ne $missingPlItems) {
        $validationStatus.Status = "not-valid"
        $validationStatus.Missing_Pl = $missingPlItems;
    }

    # Count Validation Check
    if (($data.Ci_Items.Count -ne $data.Ci_Hts.Count) -or ($data.Ci_Items.Count -ne $data.Cost.Count) -or ($data.Pl_Items.Count -ne $data.Weight.Count)) {
        $validationStatus.Status = "not-valid";
        $validationStatus.Input_Count = @($data.Ci_Items.Count, $data.Ci_Hts.Count, $data.Cost.Count, $data.Pl_Items.Count, $data.Weight.Count);
    }

    if ($validationStatus.Status -ne "not-valid") {
        $validationStatus.Status = "valid"
    }

    return $validationStatus
}

###--- Main ---###
function main {

    add-type -AssemblyName Microsoft.VisualBasic

    ###--- File path Validation ---###

    if (Test-Path -Path "$env:USERPROFILE\Documents\Invoice-Parser Directory") {
        $invParserPath = "$env:USERPROFILE\Documents\Invoice-Parser Directory"  
    }
    elseif (Test-Path -Path "$env:USERPROFILE\Desktop\Invoice-Parser Directory") {
        $invParserPath = "$env:USERPROFILE\Desktop\Invoice-Parser Directory" 
    }
    elseif (Test-Path -Path "$env:USERPROFILE\Desktop\OneDrive - Green Worldwide Shipping, LLC\Documents\Invoice-Parser Directory") {
        $invParserPath = "$env:USERPROFILE\Desktop\OneDrive - Green Worldwide Shipping, LLC\Documents\Invoice-Parser Directory"
    }
    elseif (Test-Path -Path "$env:USERPROFILE\OneDrive - Green Worldwide Shipping, LLC\Desktop\Invoice-Parser Directory") {
        $invParserPath = "$env:USERPROFILE\OneDrive - Green Worldwide Shipping, LLC\Desktop\Invoice-Parser Directory"
    }
    else {
        Read-Host -Prompt "There was an error locating 'Invoice-Parser Directory'. Make sure it is in the Documents or Desktop folder. (Press ENTER to quit)"
        return
    }

    $inputPath = "$invParserPath\input\input.xlsx"
    $outputPath = "$invParserPath\output\output.txt"

    ## -- Extract and parse input.xlsx values -- ##
    $data = parseData(returnExcelValues($inputPath))

    #--- output ---#
    # Clear output.txt, and append lines
    "" | Out-File $outputPath

    # Check if $data was returned as $validationStatus or $htsValuesHash
    if ($data.Status) { 

        if ($null -ne $data.Input_Count) {
            Add-Content -Path $outputPath -Value "!! ERROR !!`nData input count does not match. Note that CI column lengths must equal and BL column lengths must equal, respectively.`n`nCI-Item-No Count: $($data.Input_Count[0])`nCI-HTS/SCH.B Count: $($data.Input_Count[1])`nCost Count: $($data.Input_Count[2])`n----------------------`nBl-Item-No Count: $($data.Input_Count[3])`nWeight Count: $($data.Input_Count[4])`n`n"
        }
        if ($null -ne $data.Missing_Ci) {
            Add-Content -Path $outputPath -Value "!! ERROR !!`nThe following item numbers were found on the Commercial Invoice, but not on the Bill of Lading:`n$($data.Missing_Ci)`n`n"
        }
        if ($null -ne $data.Missing_Pl) {
            Add-Content -Path $outputPath -Value "!! ERROR !!`nThe following item numbers were found on the Bill of Lading, but not on the Commercial Invoice:`n$($data.Missing_Pl)"
        }

        Read-Host -Prompt "There was an error validating input.xlsx (Press ENTER to open output.txt)"
    }
    else {
        $count = 1
        $totalCost = 0
        $totalWeight = 0

        foreach ($hash in $data.GetEnumerator()) {
            $hts = $hash.Key
            $cost = "$" + "$($hash.Value[0])"
            $weight = "$($hash.Value[2])" + "kg"
            $totalCost = $totalCost + $hash.Value[0]
            $totalWeight = $totalWeight + $hash.Value[2]

            Add-Content -Path $outputPath -Value "$count.) $hts, $cost, $weight"
            $count = ($count + 1)
        }
        Add-Content -Path $outputPath -Value "`nTotal cost = $totalCost`nTotal weight = $totalWeight`n`n* DOUBLE CHECK INVOICE FOR ACCURACY!! *"
        Read-Host -Prompt "Complete! (Press ENTER to open output.txt)"
    }

    Invoke-Item $outputPath
}

main

