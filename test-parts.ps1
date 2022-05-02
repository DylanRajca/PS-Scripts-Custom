<#
.Synopsis
Parse input.xlsx data into a proper invoice line, then output to output.txt
Author: Dylan Rajca, 02/01/22
#>

###--- File paths ---###

if (Test-Path -Path "$env:USERPROFILE\Documents\Test directory") {
    $testDirectory = "$env:USERPROFILE\Documents\Test directory"  
}
elseif (Test-Path -Path "$env:USERPROFILE\Desktop\Test directory") {
    $testDirectory = "$env:USERPROFILE\Desktop\Test directory" 
}
else {
    Write-Host "Directory does not exist!"
}

$inputPath = "$testDirectory\input\input.xlsx"
$outputPath = "$testDirectory\output\output.txt"

###--- Main ---###
function addTotal($data) {
    $dataArray = @()
    foreach ($item in $data) {
        try {
            if ($item[0] -eq "$") {
                $dataArray += , -join $item[1..($item.Length - 1)]
            }
            else {
                $dataArray += , $item
            }
        }
        catch { }
    }
    $sum = $dataArray | ForEach-Object -begin { $sum = 0 }-process { $sum += $_ } -end { $sum }
    return $sum
}

function returnInput {
    Write-host "Extracting data from input.xlsx..."
    $XL = New-Object -comobject Excel.Application
    $XL.visible = $false
    
    # Assign $book and $sheet variables to input.sv
    $book = $XL.Workbooks.Open($inputPath)
    $sheet = $book.worksheets.item(1)

    # Save index of last row in database and set first row
    [int]$lastRow = ($book.UsedRange.Rows.count + 1) - 1
    [int]$firstRow = 2

    ## -- CI DATA -- ##
    $ciTariffs = $sheet.UsedRange.Rows.Columns[1].Value2 | where { $_ -notmatch "CI-HTS-CODES" -and $null -ne $_ }
    $qty = $sheet.UsedRange.Rows.Columns[3].Value2 | where { $_ -notmatch "Qty" -and $null -ne $_ }
    $cost = $sheet.UsedRange.Rows.Columns[5].Value2 | where { $_ -notmatch "Cost" -and $null -ne $_ }

    # Close excel and return data
    $book.close($true)
    $XL.Quit()
    return $ciTariffs, $qty, $cost
}

function ciParse ($data) {
    Write-Host "Parsing CIV data..."

    #--- ciTariffs ---#
    $ciTariffs = $data[0] | foreach { if ($null -ne $_) { $_.trim() } }
    $ciTariffsUnique = $ciTariffs | Select-Object -Unique

    # Save indexes of ciTariffs
    $indexHash = [ordered]@{}

    foreach ($item in $ciTariffsUnique) {
        $indexArray = @()
        for ($i = 0; $i -lt $ciTariffs.count; $i++) {
            if ($ciTariffs[$i] -eq $item) {
                $indexArray += , $i 
            }
        } 

        # Create tariff/indices hash table
        $indexHash += @{$item = $indexArray }
    }

    # Save keys in indexHash
    $hashKeys = @($indexHash.Keys)

    #--- costs / qty ---#
    $qty = $data[1]
    $cost = $data[2]

    # Generate tariff, costs/qty hash table
    $qtyHash = [ordered]@{}
    $costHash = [ordered]@{}

    # Loop through ciTariffs in $hashKeys
    foreach ($item in $hashKeys) {

        # Match cost/qty values with ciTariffs, from the indexes in $indexHash
        $qtyArray = @()
        $costArray = @()
        $valueLength = $indexHash[$item].Count
        
        # Loop through values (indices) in $indexHash and add values to $costArray, $qtyArray
        for ($i = 0; $i -lt $valueLength; $i++) {
            $qtyArray += , $qty[$indexHash[$item][$i]]
            $costArray += , $cost[$indexHash[$item][$i]]
        }
        
        # Total sum of values in $qtyArray and create tariff/qty hash table
        $totalQty = addTotal $qtyArray
        $qtyHash += @{$item = $totalQty }

        # Total sum of values in $costArray and create tariff/cost hash table
        $totalCost = addTotal $costArray
        $costHash += @{$item = $totalCost }

        # Total cost and qty for invoice
        $invoiceCost = addTotal $cost
        $invoiceQty = addTotal $qty
    }


    return $hashKeys, $costHash, $qtyHash, $invoiceCost, $invoiceQty
}

function plParse ($data) {
    Write-Host "Parsing PL data..."

    #--- plTariffs ---#
    $plTariffs = $data[3] | foreach { if ($null -ne $_) { $_.trim() } }
    $plTariffsUnique = $plTariffs | Select-Object -Unique

    # Save indexes of plTariffs
    $indexHash = [ordered]@{}

    foreach ($item in $plTariffsUnique) {
        $indexArray = @()
        for ($i = 0; $i -lt $plTariffs.count; $i++) {
            if ($plTariffs[$i] -match $item) {
                $indexArray += , $i 
            }
        } 

        # Create tariff/indices hash table
        $indexHash += @{$item = $indexArray }
    }

    # Save keys in indexHash
    $hashKeys = @($indexHash.Keys)

    #--- grossWeights / netWeight ---#
    $netWeight = $data[4]
    $grossWeight = $data[5]

    # Generate tariff, grossWeights/netWeight hash table
    $netWeightHash = [ordered]@{}
    $grossWeightHash = [ordered]@{}

    # Loop through plTariffs in $hashKeys
    foreach ($item in $hashKeys) {

        # Match grossWeight/netWeight values with plTariffs, from the indexes in $indexHash
        $netWeightArray = @()
        $grossWeightArray = @()
        $valueLength = $indexHash[$item].Count
        
        # Loop through values (indices) in $indexHash and add values to $grossWeightArray, $netWeightArray
        for ($i = 0; $i -lt $valueLength; $i++) {
            $netWeightArray += , $netWeight[$indexHash[$item][$i]]
            $grossWeightArray += , $grossWeight[$indexHash[$item][$i]]
        }
        
        # Total sum of values in $netWeightArray and create tariff/netWeight hash table
        $totalNetWeight = addTotal $netWeightArray
        $netWeightHash += @{$item = $totalNetWeight }

        # Total sum of values in $grossWeightArray and create tariff/grossWeight hash table
        $totalGrossWeight = addTotal $grossWeightArray
        $grossWeightHash += @{$item = $totalGrossWeight }

        # Total grossWeight and netWeight for invoice
        $invoiceGrossWeight = addTotal $grossWeight
        $invoiceNetWeight = addTotal $netWeight
    }


    return $hashKeys, $grossWeightHash, $netWeightHash, $invoiceGrossWeight, $invoiceNetWeight
}

function main {
    $data = returnInput

    ## -- Extract CI data -- ##
    $ciData = ciParse $data
    $ciHashKeys = $ciData[0]
    $costHash = $ciData[1]
    $qtyHash = $ciData[2]
    $costTotal = $ciData[3]
    $qtyTotal = $ciData[4]

    #--- output ---#
    
    # Clear output.txt, and append lines
    "" | Out-File $outputPath
    $count = 1
    foreach ($tariff in $ciHashKeys) {
        $cost = "$" + "$($costHash[$tariff])"
        $qty = "$($qtyHash[$tariff])NO"
        Add-Content -Path $outputPath -Value "$count.) $tariff | $qty | $cost"
        $count = ($count + 1)
    }
    Add-Content -Path $outputPath -Value "`nTotal cost = $costTotal`nTotal qty = $qtyTotal`n`n* DOUBLE CHECK INVOICE FOR ACCURACY!! *"

    Read-Host -Prompt "Complete! Press ENTER to open output.txt"
    Invoke-Item $outputPath
}
main

