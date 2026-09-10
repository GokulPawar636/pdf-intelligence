param([Parameter(Mandatory=$true)][string]$WorkbookPath)
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
function Assert-Near($Actual, $Expected, $Label) {
    if ($null -eq $Actual -or [Math]::Abs([double]$Actual - $Expected) -gt 0.0000001) {
        throw "$Label expected $Expected, received $Actual"
    }
}
function Assert-Text($Actual, $Expected, $Label) {
    if ($Actual -ne $Expected) { throw "$Label expected $Expected, received $Actual" }
}
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    # Read-only: mutations are deliberately discarded; the release artifact is preserved.
    $book = $excel.Workbooks.Open((Resolve-Path -LiteralPath $WorkbookPath).Path,0,$true)
    $sheet = $book.Worksheets.Item('Summary Report')
    $excel.CalculateFullRebuild()
    Assert-Near $sheet.Range('H7').Value2 18.081 'Initial mean'
    Assert-Text $sheet.Range('I7').Value2 'N/A' 'Initial SD'
    $sheet.Range('G7').Value2 = 18.0
    $sheet.Columns.Item(8).Insert(-4161) | Out-Null
    $sheet.Range('H5').Value2 = 'Weekly Plan'
    $sheet.Range('H6').Value2 = 'New reading'
    $sheet.Range('H7').Value2 = 18.3
    $excel.CalculateFullRebuild()
    Assert-Near $sheet.Range('I7').Value2 18.15 'Mean after inserting column'
    Assert-Near $sheet.Range('J7').Value2 ([Math]::Sqrt(0.045)) 'SD after inserting column'
    Assert-Near $sheet.Range('K7').Value2 (18.15 + 3 * [Math]::Sqrt(0.045)) 'UCL'
    Assert-Near $sheet.Range('L7').Value2 (18.15 - 3 * [Math]::Sqrt(0.045)) 'LCL'
    if ($sheet.Range('C21').Value2 -notlike '*13 readings*') { throw 'Overview reading count did not update' }
    $sheet.Range('F7').Value2 = '+/-0.1'
    $excel.CalculateFullRebuild()
    Assert-Near $sheet.Range('D7').Value2 17.9 'Edited lower tolerance'
    Assert-Near $sheet.Range('E7').Value2 18.1 'Edited upper tolerance'
    Assert-Text $sheet.Range('O7').Value2 'Out of tolerance' 'Edited tolerance result'
    if ($sheet.Range('C22').Value2 -notlike '*Out of tolerance: 1*') { throw 'Overview status count did not update' }
    $sheet.Range('C7').Value2 = 19.0
    $sheet.Range('G7').Value2 = 19.0
    $sheet.Range('H7').Value2 = 19.05
    $sheet.Range('F7').Value2 = '+0.5/-0.2'
    $excel.CalculateFullRebuild()
    Assert-Near $sheet.Range('D7').Value2 18.8 'Asymmetric lower limit'
    Assert-Near $sheet.Range('E7').Value2 19.5 'Asymmetric upper limit'
    Assert-Text $sheet.Range('O7').Value2 'OK' 'Updated result'
    Assert-Near $book.Worksheets.Item('Differences').Range('F5').Value2 0.0 'Linked difference'
    Assert-Near $book.Worksheets.Item('Differences').Range('E20').Value2 19.05 'Latest-reading link includes added column'
    Assert-Near $book.Worksheets.Item('Differences').Range('F20').Value2 0.05 'Latest-reading difference'
    $sheet.Range('H7').Value2 = 'bad reading'
    $excel.CalculateFullRebuild()
    Assert-Text $sheet.Range('I7').Value2 'Review' 'Invalid reading mean'
    Assert-Text $sheet.Range('O7').Value2 'Review' 'Invalid reading status'
    $sheet.Range('H7').ClearContents()
    $excel.CalculateFullRebuild()
    Assert-Text $sheet.Range('J7').Value2 'N/A' 'Cleared reading SD'
    $sheet.Columns.Item(8).Delete() | Out-Null
    $excel.CalculateFullRebuild()
    Assert-Near $sheet.Range('H7').Value2 19.0 'Mean after deleting added column'
    Assert-Text $sheet.Range('I7').Value2 'N/A' 'SD after deleting added column'
    $sheet.Columns.Item(7).Insert(-4161) | Out-Null
    $sheet.Range('G7').Value2 = 18.8
    $excel.CalculateFullRebuild()
    Assert-Near $sheet.Range('I7').Value2 18.9 'New first reading included'
    Write-Output 'PASS: Excel recalculation, inserted/deleted columns, SD/UCL/LCL, tolerance edits, overview, linked differences, and invalid readings.'
} finally {
    if ($book) { $book.Close($false) }
    if ($excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
