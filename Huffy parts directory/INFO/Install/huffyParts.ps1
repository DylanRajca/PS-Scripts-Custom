###--- File paths ---###

$huffyPartsDirectory = "$env:USERPROFILE\Desktop\Huffy parts directory" 
$inputPath = "$huffyPartsDirectory\input\input.xlsx"
$outputPath = "$huffyPartsDirectory\output\output.txt"

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
    $XL = New-Object -comobject Excel.Application
    $XL.visible = $false
    
    # Assign $book and $sheet variables to input.sv
    $book = $XL.Workbooks.Open($inputPath)
    $sheet = $book.worksheets.item(1)

    # Save index of last row in database and set first row
    [int]$lastRow = ($book.UsedRange.Rows.count + 1) - 1
    [int]$firstRow = 2

    # grab tariff, qty, cost data
    $tariffs = $sheet.UsedRange.Rows.Columns[1].Value2 | where { $_ -notmatch "Hts Numbers" }
    $qty = $sheet.UsedRange.Rows.Columns[3].Value2 | where { $_ -notmatch "Qty" }
    $cost = $sheet.UsedRange.Rows.Columns[5].Value2 | where { $_ -notmatch "Cost" }

    # Close excel and return data
    $book.close($true)
    $XL.Quit()
    return $tariffs, $qty, $cost
}

function main {
    $data = returnInput

    #--- tariffs ---#
    $tariffs = $data[0]
    $tariffsUnique = $tariffs | Select-Object -Unique

    # Save indexes of tariffs
    $indexHash = [ordered]@{}

    foreach ($item in $tariffsUnique) {
        $indexArray = @()
        for ($i = 0; $i -lt $tariffs.count; $i++) {
            if ($tariffs[$i] -match $item) {
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

    # Loop through tariffs in $hashKeys
    foreach ($item in $hashKeys) {

        # Match cost/qty values with tariffs, from the indexes in $indexHash
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

    #--- output ---#
    
    # Clear output.txt, and append lines
    "" | Out-File $outputPath
    $count = 1
    foreach ($tariff in $hashKeys) {
        $cost = "$" + "$($costHash[$tariff])"
        $qty = "$($qtyHash[$tariff])pcs"
        Add-Content -Path $outputPath -Value "$count. $tariff, $cost, $qty"
        $count = ($count + 1)
    }
    Add-Content -Path $outputPath -Value "`nTotal cost = $invoiceCost`nTotal qty = $invoiceQty`n`n* DOUBLE CHECK INVOICE FOR ACCURACY!! *"
}
main