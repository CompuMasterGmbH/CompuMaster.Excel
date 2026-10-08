<#
.SYNOPSIS
Creates HTML color test workbooks and captures reference colors with Microsoft Excel.
.DESCRIPTION
Uses a private Excel instance. Reference colors are captured after saving and reopening
each workbook. Run only when intentionally refreshing the checked-in reference data.
#>
[CmdletBinding()]
param (
    [string] $ThemeDirectory = 'C:\Program Files\Microsoft Office\root\Document Themes 16',
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\ExcelOpsTest\test_data'),
    [string] $PreviewDirectory,
    [switch] $Overwrite
)

$ErrorActionPreference = 'Stop'
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
$sourcePath = Join-Path $PSScriptRoot '..\ExcelOpsTest\test_data\HtmlExportThemeExcelLegacy.xlsx'
$variants = @(
    @{ Id = 'Legacy'; Theme = $null; Label = 'Office legacy (existing test palette)' },
    @{ Id = 'Office'; Theme = 'Office Theme.thmx'; Label = 'Office' },
    @{ Id = 'Office2013'; Theme = 'Office 2013 - 2022 Theme.thmx'; Label = 'Office 2013 - 2022' },
    @{ Id = 'Ion'; Theme = 'Ion.thmx'; Label = 'Ion' },
    @{ Id = 'Red'; Theme = 'Office Theme.thmx'; Label = 'Office with custom red Accent 2' }
)
$referenceName = 'HtmlExportThemeExcelReferences.csv'
$outputNames = @($variants | ForEach-Object { 'HtmlExportThemeExcel' + $_.Id + '.xlsx' }) + $referenceName
foreach ($name in $outputNames) {
    if ((Test-Path -LiteralPath (Join-Path $OutputDirectory $name)) -and -not $Overwrite) {
        throw "Output already exists: $name. Use -Overwrite only for an intentional reference refresh."
    }
}
foreach ($variant in $variants) {
    if ($variant.Theme -and -not (Test-Path -LiteralPath (Join-Path $ThemeDirectory $variant.Theme))) {
        throw "Required Excel theme not found: $($variant.Theme)"
    }
}
[void][System.IO.Directory]::CreateDirectory($OutputDirectory)
if ($PreviewDirectory) {
    $PreviewDirectory = [System.IO.Path]::GetFullPath($PreviewDirectory)
    [void][System.IO.Directory]::CreateDirectory($PreviewDirectory)
}

# Reuse the palette from the existing Excel-generated legacy reference workbook.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$sourceArchive = [System.IO.Compression.ZipFile]::OpenRead([System.IO.Path]::GetFullPath($sourcePath))
try {
    $themeReader = [System.IO.StreamReader]::new($sourceArchive.GetEntry('xl/theme/theme1.xml').Open())
    try { [xml] $legacyTheme = $themeReader.ReadToEnd() } finally { $themeReader.Dispose() }
} finally { $sourceArchive.Dispose() }
$themeNamespaces = [System.Xml.XmlNamespaceManager]::new($legacyTheme.NameTable)
$themeNamespaces.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
$legacyPalette = @($legacyTheme.SelectNodes('/a:theme/a:themeElements/a:clrScheme/*', $themeNamespaces) | ForEach-Object {
    $colorNode = $_.FirstChild
    $rgb = if ($colorNode.LocalName -eq 'sysClr') { $colorNode.GetAttribute('lastClr') } else { $colorNode.GetAttribute('val') }
    $red = [Convert]::ToInt32($rgb.Substring(0, 2), 16)
    $green = [Convert]::ToInt32($rgb.Substring(2, 2), 16)
    $blue = [Convert]::ToInt32($rgb.Substring(4, 2), 16)
    $red + ($green -shl 8) + ($blue -shl 16)
})

function ConvertTo-CssHex([long] $Color) {
    # Excel COM colors store the red byte first (OLE_COLOR), unlike CSS hex strings.
    return '#{0:X2}{1:X2}{2:X2}' -f ($Color -band 255), (($Color -shr 8) -band 255), (($Color -shr 16) -band 255)
}

function Release-ComObject($Value) {
    if ($null -ne $Value -and [System.Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
    }
}

$excel = $null
$workbook = $null
$sheet = $null
$references = [System.Collections.Generic.List[object]]::new()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    foreach ($variant in $variants) {
        $fileName = 'HtmlExportThemeExcel' + $variant.Id + '.xlsx'
        $filePath = Join-Path $OutputDirectory $fileName
        $workbook = $excel.Workbooks.Add(-4167)
        if ($variant.Theme) { $workbook.ApplyTheme((Join-Path $ThemeDirectory $variant.Theme)) }
        if ($variant.Id -eq 'Legacy') {
            for ($themeIndex = 1; $themeIndex -le $legacyPalette.Count; $themeIndex++) {
                $workbook.Theme.ThemeColorScheme.Colors($themeIndex).RGB = $legacyPalette[$themeIndex - 1]
            }
        }
        if ($variant.Id -eq 'Red') { $workbook.Theme.ThemeColorScheme.Colors(6).RGB = 255 }
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = 'Accent reference'
        [void] $sheet.Cells.Clear()
        [void] $sheet.Range('A1:F1').Merge()
        $sheet.Range('A1').Value2 = 'HTML-Farbreferenz: ' + $variant.Label
        $sheet.Range('A1').Font.Size = 18
        $sheet.Range('A1').Font.Bold = $true
        [void] $sheet.Range('A2:F2').Merge()
        $sheet.Range('A2').Value2 = 'Referenz aus Microsoft Excel ' + $excel.Version + ' (Build ' + $excel.Build + ')'
        [void] $sheet.Range('A3:F3').Merge()
        $sheet.Range('A3').Value2 = 'Sollfarben gelten fuer das gespeicherte Design. Schrift und Fuellung verwenden echte Akzentfarben.'
        [void] $sheet.Range('A4:F4').Merge()
        $sheet.Range('A4').Value2 = '0 = Grundfarbe, -0.25 = dunkler, +0.60 = heller. Feste RGB-Kontrollen bleiben bei Designwechsel gleich.'
        $headings = @('Testfall', 'Schriftprobe', 'Fuellprobe', 'Tint', 'Soll Schrift B', 'Soll Fuellung C')
        for ($column = 1; $column -le $headings.Count; $column++) {
            $sheet.Cells.Item(6, $column).Value2 = $headings[$column - 1]
        }
        $sheet.Range('A6:F6').Font.Bold = $true
        $sheet.Range('A6:F6').Interior.Color = 15132390
        $cases = [System.Collections.Generic.List[object]]::new()
        $row = 7
        foreach ($accent in 1..6) {
            foreach ($tint in @(0.0, -0.25, 0.60)) {
                $caseId = 'Accent' + $accent + '-' + @{'0' = 'Base'; '-0.25' = 'Dark'; '0.6' = 'Light'}[$tint.ToString([System.Globalization.CultureInfo]::InvariantCulture)]
                $sheet.Cells.Item($row, 1).Value2 = $caseId
                $sheet.Range('D' + $row).Formula = $tint.ToString([System.Globalization.CultureInfo]::InvariantCulture)
                $fontCell = $sheet.Range('B' + $row)
                $fillCell = $sheet.Range('C' + $row)
                $fontCell.Interior.Pattern = 1
                $fontCell.Interior.Color = 16777215
                $fontCell.Font.ThemeColor = $accent + 4
                $fontCell.Font.TintAndShade = $tint
                $fillCell.Interior.Pattern = 1
                $fillCell.Interior.ThemeColor = $accent + 4
                $fillCell.Interior.TintAndShade = $tint
                $fillCell.Font.Color = 0
                $fontHex = ConvertTo-CssHex $fontCell.DisplayFormat.Font.Color
                $fillHex = ConvertTo-CssHex $fillCell.DisplayFormat.Interior.Color
                $fontLabel = 'Text Akzent ' + $accent + ' (' + $fontHex + ') auf Weiss'
                if ($variant.Id -eq 'Red' -and $accent -eq 2 -and $tint -eq 0) {
                    $fontLabel = 'Roter Text (#FF0000) auf weissem Untergrund'
                }
                $fontCell.Value2 = $fontLabel
                $fillCell.Value2 = 'Schwarzer Text auf Akzent ' + $accent + ' (' + $fillHex + ')'
                $sheet.Cells.Item($row, 5).Value2 = $fontHex
                $sheet.Cells.Item($row, 6).Value2 = $fillHex
                $cases.Add(@{ CaseId = $caseId + '-Font'; Address = 'B' + $row })
                $cases.Add(@{ CaseId = $caseId + '-Fill'; Address = 'C' + $row })
                Release-ComObject $fontCell
                Release-ComObject $fillCell
                $row++
            }
        }
        $sheet.Cells.Item($row, 1).Value2 = 'FixedRgbRed'
        $sheet.Cells.Item($row, 2).Value2 = 'Fester roter Text (#FF0000) auf Weiss'
        $sheet.Cells.Item($row, 3).Value2 = 'Weisser Text auf festem Rot (#FF0000)'
        $sheet.Range('B' + $row).Font.Color = 255
        $sheet.Range('B' + $row).Interior.Color = 16777215
        $sheet.Range('C' + $row).Font.Color = 16777215
        $sheet.Range('C' + $row).Interior.Color = 255
        $sheet.Cells.Item($row, 5).Value2 = '#FF0000'
        $sheet.Cells.Item($row, 6).Value2 = '#FF0000'
        $cases.Add(@{ CaseId = 'FixedRgbRed-Font'; Address = 'B' + $row })
        $cases.Add(@{ CaseId = 'FixedRgbRed-Fill'; Address = 'C' + $row })
        $row++
        $sheet.Cells.Item($row, 1).Value2 = 'NoFill'
        $sheet.Cells.Item($row, 2).Value2 = 'Schwarzer Text ohne Fuellung'
        $sheet.Range('B' + $row).Font.Color = 0
        $sheet.Range('B' + $row).Interior.Pattern = -4142
        $sheet.Cells.Item($row, 5).Value2 = '#000000'
        $cases.Add(@{ CaseId = 'NoFill'; Address = 'B' + $row })
        $sheet.Range('A6:F' + $row).Font.Size = 11
        $sheet.Range('A6:F' + $row).WrapText = $true
        $sheet.Range('A6:F' + $row).RowHeight = 34
        $sheet.Range('A:A').ColumnWidth = 20
        $sheet.Range('B:C').ColumnWidth = 51
        $sheet.Range('D:D').ColumnWidth = 10
        $sheet.Range('E:F').ColumnWidth = 18
        $sheet.Range('D7:D' + $row).NumberFormatLocal = '0' + $excel.International(3) + '00'
        $sheet.Range('A1:F4').RowHeight = 25
        $sheet.PageSetup.Orientation = 2
        $sheet.PageSetup.Zoom = $false
        $sheet.PageSetup.FitToPagesWide = 1
        $sheet.PageSetup.FitToPagesTall = 1
        $sheet.PageSetup.PrintArea = '$A$1:$F$' + $row
        $excel.ActiveWindow.SplitRow = 6
        $excel.ActiveWindow.FreezePanes = $true
        [void] $sheet.Range('A1').Select()
        $workbook.SaveAs($filePath, 51)
        $workbook.Close($false)
        Release-ComObject $sheet
        $sheet = $null
        Release-ComObject $workbook
        $workbook = $null

        # The saved workbook, not transient authoring state, is the test input and color oracle.
        $workbook = $excel.Workbooks.Open($filePath, 0, $true)
        $sheet = $workbook.Worksheets.Item(1)
        $excel.CalculateFull()
        $sha256 = (Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash
        foreach ($case in $cases) {
            $cell = $sheet.Range($case.Address)
            $fillHex = if ($cell.DisplayFormat.Interior.Pattern -eq -4142) { '' } else { ConvertTo-CssHex $cell.DisplayFormat.Interior.Color }
            $references.Add([pscustomobject][ordered]@{
                Workbook = $fileName
                WorkbookSha256 = $sha256
                Theme = $variant.Label
                Sheet = $sheet.Name
                Address = $case.Address
                CaseId = $case.CaseId
                Text = $cell.Value2
                FontRgb = (ConvertTo-CssHex $cell.DisplayFormat.Font.Color)
                FillRgb = $fillHex
                ExcelVersion = $excel.Version
                ExcelBuild = $excel.Build
            })
            Release-ComObject $cell
        }
        if ($PreviewDirectory) {
            $sheet.ExportAsFixedFormat(0, (Join-Path $PreviewDirectory ($variant.Id + '.pdf')))
        }
        Write-Output "$fileName`: $($cases.Count) reference cells captured with Excel"
        $workbook.Close($false)
        Release-ComObject $sheet
        $sheet = $null
        Release-ComObject $workbook
        $workbook = $null
    }
    $references | Export-Csv -LiteralPath (Join-Path $OutputDirectory $referenceName) -NoTypeInformation -Encoding utf8BOM
} finally {
    if ($null -ne $workbook) { $workbook.Close($false) }
    Release-ComObject $sheet
    Release-ComObject $workbook
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($null -ne $excel) { $excel.Quit(); Release-ComObject $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
