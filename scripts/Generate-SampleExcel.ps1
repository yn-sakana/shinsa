param(
    [string]$OutputDir = (Join-Path $PSScriptRoot '..\data\sample\onedrive\table')
)

$ErrorActionPreference = 'Stop'

# Read JSON source files
$ankenJson = Get-Content (Join-Path $OutputDir 'anken.source.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$contactsJson = Get-Content (Join-Path $OutputDir 'contacts.source.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$kenshuJson = Get-Content (Join-Path $OutputDir 'kenshu.source.json') -Raw -Encoding UTF8 | ConvertFrom-Json

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # --- anken.xlsx with structured table ---
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'anken'

    $headers = @($ankenJson[0].PSObject.Properties.Name)
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }
    for ($r = 0; $r -lt $ankenJson.Count; $r++) {
        $rec = $ankenJson[$r]
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $val = $rec.PSObject.Properties[$headers[$c]].Value
            $ws.Cells.Item($r + 2, $c + 1).Value2 = [string]$val
        }
    }

    $dataRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($ankenJson.Count + 1, $headers.Count))
    $listObj = $ws.ListObjects.Add(1, $dataRange, $null, 1)  # xlSrcRange=1, xlYes=1
    $listObj.Name = 'T_anken'

    $outPath = Join-Path $OutputDir 'anken.xlsx'
    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $wb.SaveAs($outPath, 51)  # xlOpenXMLWorkbook=51
    $wb.Close($false)
    Write-Host "Created: $outPath (table: T_anken, $($ankenJson.Count) rows)"

    # --- contacts.xlsx with structured table ---
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'contacts'

    $headers = @($contactsJson[0].PSObject.Properties.Name)
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }
    for ($r = 0; $r -lt $contactsJson.Count; $r++) {
        $rec = $contactsJson[$r]
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $val = $rec.PSObject.Properties[$headers[$c]].Value
            $ws.Cells.Item($r + 2, $c + 1).Value2 = [string]$val
        }
    }

    $dataRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($contactsJson.Count + 1, $headers.Count))
    $listObj = $ws.ListObjects.Add(1, $dataRange, $null, 1)
    $listObj.Name = 'T_contacts'

    $outPath = Join-Path $OutputDir 'contacts.xlsx'
    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $wb.SaveAs($outPath, 51)
    $wb.Close($false)
    Write-Host "Created: $outPath (table: T_contacts, $($contactsJson.Count) rows)"

    # --- kenshu.xlsx with structured table ---
    $wb = $excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = 'kenshu'

    $headers = @($kenshuJson[0].PSObject.Properties.Name)
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }
    for ($r = 0; $r -lt $kenshuJson.Count; $r++) {
        $rec = $kenshuJson[$r]
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $val = $rec.PSObject.Properties[$headers[$c]].Value
            $ws.Cells.Item($r + 2, $c + 1).Value2 = [string]$val
        }
    }

    $dataRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($kenshuJson.Count + 1, $headers.Count))
    $listObj = $ws.ListObjects.Add(1, $dataRange, $null, 1)
    $listObj.Name = 'T_kenshu'

    $outPath = Join-Path $OutputDir 'kenshu.xlsx'
    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $wb.SaveAs($outPath, 51)
    $wb.Close($false)
    Write-Host "Created: $outPath (table: T_kenshu, $($kenshuJson.Count) rows)"

    Write-Host "`nDone. Update config.local.json source_path and source_table to use these files."
}
finally {
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $excel) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
