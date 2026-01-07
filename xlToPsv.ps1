$excelPath  = "C:\Path\To\Your\File.xlsx"
$outputPath = "C:\Path\To\output.txt"

# ExcelHeader = OutputHeader
$columnMap = @{
    "ID"        = "DialogueID"
    "VoiceType" = "Voice"
    "Dialogue"  = "Line"
    "Emotion"   = "Mood"
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open($excelPath)

Remove-Item $outputPath -ErrorAction SilentlyContinue

# Write header row once (optional)
$headerLine = "Sheet|" + ($columnMap.Values -join "|")
Add-Content -Path $outputPath -Value $headerLine

foreach ($sheet in $workbook.Worksheets) {
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    # Map Excel headers to column indexes
    $headerIndex = @{}
    for ($col = 1; $col -le $colCount; $col++) {
        $header = $usedRange.Cells.Item(1, $col).Text
        if ($columnMap.ContainsKey($header)) {
            $headerIndex[$header] = $col
        }
    }

    if ($headerIndex.Count -eq 0) { continue }

    for ($row = 2; $row -le $rowCount; $row++) {
        $values = foreach ($excelHeader in $columnMap.Keys) {
            if ($headerIndex.ContainsKey($excelHeader)) {
                $usedRange.Cells.Item($row, $headerIndex[$excelHeader]).Text
            } else {
                ""
            }
        }

        $line = $sheet.Name + "|" + ($values -join "|")
        Add-Content -Path $outputPath -Value $line
    }
}

$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
