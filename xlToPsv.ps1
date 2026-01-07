$excelPath  = "C:\Path\To\Your\File.xlsx"
$outputPath = "C:\Path\To\output.txt"

$columnMap = @{
    "MaleCommander.ID"        = "DialogueID"
    "MaleCommander.Dialogue"  = "Line"
    "MaleCommander.Emotion"   = "Mood"

    "FemaleYoungEager.ID"     = "DialogueID"
    "FemaleYoungEager.Text"   = "Line"
    "FemaleYoungEager.Tone"   = "Mood"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open($excelPath)
Remove-Item $outputPath -ErrorAction SilentlyContinue

# Determine output column order (stable & predictable)
$outputColumns = $columnMap.Values | Select-Object -Unique
Add-Content $outputPath ("Sheet|" + ($outputColumns -join "|"))

foreach ($sheet in $workbook.Worksheets) {
    $sheetName = $sheet.Name
    $usedRange = $sheet.UsedRange
    $rowCount  = $usedRange.Rows.Count
    $colCount  = $usedRange.Columns.Count

    # Build Excel header → column index map
    $excelHeaders = @{}
    for ($col = 1; $col -le $colCount; $col++) {
        $header = $usedRange.Cells.Item(1, $col).Text
        if ($header) {
            $excelHeaders[$header] = $col
        }
    }

    # Find mappings relevant to this sheet
    $sheetMappings = $columnMap.Keys |
        Where-Object { $_ -like "$sheetName.*" }

    if (-not $sheetMappings) { continue }

    for ($row = 2; $row -le $rowCount; $row++) {
        $rowData = @{}

        foreach ($key in $sheetMappings) {
            $parts = $key.Split(".")
            $columnName = $parts[1]
            $outputName = $columnMap[$key]

            if ($excelHeaders.ContainsKey($columnName)) {
                $value = $usedRange.Cells.Item($row, $excelHeaders[$columnName]).Text
                $rowData[$outputName] = $value
            }
        }

        # Emit values in fixed output column order
        $values = foreach ($outCol in $outputColumns) {
            $rowData[$outCol]
        }

        Add-Content $outputPath ($sheetName + "|" + ($values -join "|"))
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
