param([string]$WorkbookPath = "", [switch]$SkipReference)
$ErrorActionPreference = 'Stop'
$workspace = $PSScriptRoot
$outputPath = if ($WorkbookPath) { $WorkbookPath } else { Join-Path $workspace 'data\output\PDF1\PDF1_summary.xlsx' }
$previewDirectory = Join-Path $workspace 'tmp\summary_review'
New-Item -ItemType Directory -Force -Path $previewDirectory | Out-Null
$excel = $null
$book = $null
$reference = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3
    if (-not $SkipReference) {
    $reference = $excel.Workbooks.Open('C:\Users\GOKUL PAWAR\Downloads\KA2 Center Stand Summery Report 2025.xls', 0, $true)
    $reference.Worksheets.Item(1).ExportAsFixedFormat(0, (Join-Path $previewDirectory 'sample_reference.pdf'))
    $reference.Close($false)
    $reference = $null
    }
    $book = $excel.Workbooks.Open($outputPath, 0, $false)
    $excel.CalculateFullRebuild()
    $book.Save()
    $book.Worksheets.Item(1).ExportAsFixedFormat(0, (Join-Path $previewDirectory 'summary_preview.pdf'))
    Write-Output 'Excel recalculation and PDF previews completed.'
}
finally {
    if ($reference) { $reference.Close($false) }
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
