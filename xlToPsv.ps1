$excelPath = "C:\Path\To\Your\File.xlsx"
$outputPath = "C:\Path\To\output.txt"

# Columns you want, by header name as they appear in row 1
$wantedColumns = @("ID", "VoiceType", "Dialogue", "Emotion")

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$workbook = $excel.Workbooks.Open($excelPath)

# Start with a clean file
Remove-Item $outputPath -ErrorAction SilentlyContinue

foreach ($sheet in $workbook.Worksheets) {
    $usedRange = $sheet.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    # Map headers to column indexes
    $headerMap = @{}
    for ($col = 1; $col -le $colCount; $col++) {
        $header = $usedRange.Cells.Item(1, $col).Text
        if ($wantedColumns -contains $header) {
            $headerMap[$header] = $col
        }
    }

    # Skip sheet if none of the desired columns exist
    if ($headerMap.Count -eq 0) { continue }

    for ($row = 2; $row -le $rowCount; $row++) {
        $values = foreach ($colName in $wantedColumns) {
            if ($headerMap.ContainsKey($colName)) {
                $usedRange.Cells.Item($row, $headerMap[$colName]).Text
            } else {
                ""
            }
        }

        # Optional: prepend sheet name
        $line = ($sheet.Name + "|" + ($values -join "|"))

        Add-Content -Path $outputPath -Value $line
    }
}

$workbook.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
