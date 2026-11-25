param(
    [Parameter(Mandatory)]
    [string] $OutputFile
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

# IDs in the documented Excel XlChartType enumeration, including selector values.
$types = @(-4169, -4152, -4151, -4120, -4102, -4101, -4100, -4098, -2, 1, 4, 5, 15) + (51..140)
$interopPath = Join-Path $PSScriptRoot '../ExcelOps-MicrosoftExcel/libs-NonWindowsPlatforms/Microsoft.Office.Interop.Excel.dll'
$chartType = [Reflection.Assembly]::LoadFrom([IO.Path]::GetFullPath($interopPath)).GetType('Microsoft.Office.Interop.Excel.XlChartType')
$recentTypes = @(
    'xlOtherCombinations', 'xlTreemap', 'xlHistogram', 'xlWaterfall', 'xlSunburst', 'xlBoxwhisker',
    'xlPareto', 'xlFunnel', 'xlColumnClusteredEx', 'xlColumnStackedEx', 'xlColumnStacked100Ex',
    'xlLineEx', 'xlLineStackedEx', 'xlLineStacked100Ex', 'xlPieEx', 'xlDoughnutEx', 'xlBarClusteredEx',
    'xlBarStackedEx', 'xlBarStacked100Ex', 'xlAreaEx', 'xlAreaStackedEx', 'xlAreaStacked100Ex',
    'xlXYScatterEx', 'xlBubbleEx', 'xlRegionMap'
)
$additionalNames = @{
    -4152 = 'xlCombo'; -2 = 'xlSuggestedChart'; 113 = 'xlComboColumnClusteredLine'
    114 = 'xlComboColumnClusteredLineSecondaryAxis'; 115 = 'xlComboAreaStackedColumnClustered'
}
$tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd([IO.Path]::DirectorySeparatorChar)
$scratch = [IO.Directory]::CreateDirectory((Join-Path $tempRoot ('ExcelOps-ChartSurvey-' + [guid]::NewGuid().ToString('N')))).FullName

function Read-Part($package, [string] $name) {
    $reader = [IO.StreamReader]::new($package.GetEntry($name).Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Read-ChartReferences([string] $path) {
    $package = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $result = @{}
        $names = @{}
        $book = Read-Part $package 'xl/workbook.xml'
        foreach ($name in $book.SelectNodes('//*[local-name()="definedName"]')) {
            $names[$name.GetAttribute('name')] = $name.InnerText
        }
        foreach ($drawing in $package.Entries | Where-Object { $_.FullName -match '^xl/drawings/drawing[0-9]+\.xml$' }) {
            $drawingXml = Read-Part $package $drawing.FullName
            $rels = Read-Part $package ('xl/drawings/_rels/' + $drawing.Name + '.rels')
            $targets = @{}
            foreach ($link in $rels.DocumentElement.ChildNodes) { $targets[$link.Id] = $link.Target }
            foreach ($frame in $drawingXml.SelectNodes('//*[local-name()="graphicFrame"]')) {
                $name = $frame.SelectSingleNode('.//*[local-name()="cNvPr"]').GetAttribute('name')
                if (-not $name.StartsWith('ChartType_') -or $result.ContainsKey($name)) { continue }
                $reference = $frame.SelectSingleNode('.//*[local-name()="chart"]')
                if ($null -eq $reference) { continue }
                $id = $reference.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
                $partUri = [uri]::new([uri]('https://package.invalid/' + $drawing.FullName), [string]$targets[$id])
                $chart = Read-Part $package $partUri.AbsolutePath.TrimStart('/')
                $formulas = foreach ($formula in $chart.SelectNodes('//*[local-name()="f"]')) {
                    $expression = $formula.InnerText
                    if ($names.ContainsKey($expression)) { $expression = $names[$expression] }
                    $role = $formula.ParentNode.LocalName
                    if ($role -in @('numRef', 'strRef', 'multiLvlStrRef')) { $role = $formula.ParentNode.ParentNode.LocalName }
                    $role + '=' + $expression
                }
                $layouts = @($chart.SelectNodes('//*[@layoutId]') | ForEach-Object { $_.GetAttribute('layoutId') } | Sort-Object -Unique)
                if ($layouts.Count -eq 0) {
                    $layouts = @($chart.SelectNodes('//*[local-name()="plotArea"]/*') |
                        Where-Object { $_.LocalName.EndsWith('Chart') } | ForEach-Object LocalName | Sort-Object -Unique)
                }
                $result[$name] = [pscustomobject]@{
                    Format = if ($chart.DocumentElement.NamespaceURI.EndsWith('/chartex')) { 'ChartEx' } else { 'Classic' }
                    Layout = $layouts -join '; '
                    References = $formulas -join '; '
                }
            }
        }
        return $result
    } finally {
        $package.Dispose()
    }
}

function Write-Inputs($sheet, [string[]] $headers, [scriptblock] $values) {
    for ($column = 0; $column -lt $headers.Count; $column++) {
        $letter = [string][char]([int][char]'A' + $column)
        $sheet.Range($letter + '1').Value2 = $headers[$column]
        for ($row = 2; $row -le 7; $row++) {
            $address = $letter + [string]$row
            $value = & $values $column $row
            if ($value -is [string]) {
                $sheet.Range($address).Value2 = [string]$value
            } else {
                $sheet.Range($address).Value2 = [double]$value
            }
        }
    }
}

$excel = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $book = $excel.Workbooks.Add(-4167)
    $data = $book.Worksheets.Item(1)
    $data.Name = 'Chart Data'
    Write-Inputs $data @('Category', 'First', 'Second', 'Third') {
        param($column, $row)
        if ($column -eq 0) { return ('Item ' + $row) }
        return [double]($column * 100 + $row * 9)
    }
    $numbers = $book.Worksheets.Add([Type]::Missing, $data)
    $numbers.Name = 'Numeric Data'
    Write-Inputs $numbers @('X', 'Y', 'Size', 'Other') {
        param($column, $row)
        return [double]($column * 100 + $row * 9)
    }
    $stockSheets = @{}
    foreach ($spec in @(
        @{ Id = 88; Fields = @('Date', 'High', 'Low', 'Close') },
        @{ Id = 89; Fields = @('Date', 'Open', 'High', 'Low', 'Close') },
        @{ Id = 90; Fields = @('Date', 'Volume', 'High', 'Low', 'Close') },
        @{ Id = 91; Fields = @('Date', 'Volume', 'Open', 'High', 'Low', 'Close') }
    )) {
        $stock = $book.Worksheets.Add()
        $stock.Name = 'Stock ' + $spec.Id
        $stockSheets[$spec.Id] = $stock
        $fields = $spec.Fields
        Write-Inputs $stock $fields {
            param($column, $row)
            $base = @{ Date = 45660; Volume = 1000; Open = 90; High = 110; Low = 80; Close = 95 }
            return [double]($base[$fields[$column]] + $row)
        }
        $stock.Range('A2:A7').NumberFormat = 'yyyy-mm-dd'
    }
    $map = $book.Worksheets.Add()
    $map.Name = 'Map Data'
    $countries = @('Germany', 'France', 'Canada', 'Brazil', 'Australia', 'Japan')
    Write-Inputs $map @('Country', 'Measure') {
        param($column, $row)
        if ($column -eq 0) { return $countries[$row - 2] }
        return [double]$row
    }

    $results = foreach ($type in $types) {
        $typeName = if ($type -ge 116) { $recentTypes[$type - 116] } elseif ($additionalNames.ContainsKey($type)) { $additionalNames[$type] } else { [Enum]::GetName($chartType, [int]$type) }
        $entry = [ordered]@{ ExcelVersion = [string]$excel.Version; ExcelBuild = [string]$excel.Build; RequestedType = $type; RequestedName = $typeName
            ActualType = ''; Status = ''; Format = ''; Layout = ''; Before = ''; AfterRowInsert = ''; AfterColumnDelete = '' }
        if ($type -in -2, 116) {
            $entry.Status = 'Selector, not a concrete chart type'
            [pscustomobject]$entry
            continue
        }
        $chartSheet = $data
        $address = 'A1:D7'
        if ($type -in -4169, 15, 72, 73, 74, 75, 87, 138, 139) { $chartSheet = $numbers }
        if ($type -in 15, 87, 139) { $address = 'A1:C7' }
        if ($type -in 118, 122) { $chartSheet = $numbers; $address = 'B1:B7' }
        if ($type -eq 140) { $chartSheet = $map; $address = 'A1:B7' }
        if ($stockSheets.ContainsKey($type)) {
            $chartSheet = $stockSheets[$type]
            $lastColumn = @{ 88 = 'D'; 89 = 'E'; 90 = 'E'; 91 = 'F' }
            $address = 'A1:' + $lastColumn[$type] + '7'
        }
        $shape = $null
        try {
            $chartSheet.Activate()
            [void]$chartSheet.Range($address).Select()
            if ($type -eq -4152) {
                $shape = $chartSheet.Shapes.AddChart2(-1, 51, 240, 20, 220, 150)
                $shape.Chart.SeriesCollection().Item(2).ChartType = 4
            } elseif ($type -in 83, 84, 85, 86, 87, 88, 89, 90, 91, 113, 114, 115 -or ($type -ge 124 -and $type -le 139)) {
                $shape = $chartSheet.Shapes.AddChart2(-1, 51, 240, 20, 220, 150)
                $shape.Chart.ChartType = $type
            } else {
                $shape = $chartSheet.Shapes.AddChart2(-1, $type, 240, 20, 220, 150)
            }
            $shape.Name = 'ChartType_' + $type
            $entry.ActualType = [int]$shape.Chart.ChartType
            $entry.Status = if ($entry.ActualType -eq $type) { 'Created' } else { 'Excel normalized the type' }
        } catch [System.Runtime.InteropServices.COMException], [System.ArgumentException] {
            if ($null -ne $shape) { $shape.Delete() }
            $entry.Status = 'Creation rejected with this setup'
        }
        [pscustomobject]$entry
    }

    $beforePath = Join-Path $scratch 'before.xlsx'
    $insertPath = Join-Path $scratch 'insert.xlsx'
    $deletePath = Join-Path $scratch 'delete.xlsx'
    $book.SaveAs($beforePath, 51)
    foreach ($sheet in $book.Worksheets) { [void]$sheet.Range('4:4').EntireRow.Insert() }
    $book.SaveAs($insertPath, 51)
    foreach ($sheet in $book.Worksheets) { [void]$sheet.Range('B:B').EntireColumn.Delete() }
    $book.SaveAs($deletePath, 51)
    $book.Close($false)
    $book = $null

    $before = Read-ChartReferences $beforePath
    $afterInsert = Read-ChartReferences $insertPath
    $afterDelete = Read-ChartReferences $deletePath
    foreach ($result in $results) {
        $name = 'ChartType_' + $result.RequestedType
        if (-not $before.ContainsKey($name)) { continue }
        $result.Format = $before[$name].Format
        $result.Layout = $before[$name].Layout
        $result.Before = $before[$name].References
        if ($afterInsert.ContainsKey($name)) { $result.AfterRowInsert = $afterInsert[$name].References }
        if ($afterDelete.ContainsKey($name)) { $result.AfterColumnDelete = $afterDelete[$name].References }
    }
    $results | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding utf8BOM
    $results | Group-Object Status | Select-Object Count, Name
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    if ([IO.Path]::GetDirectoryName($scratch) -eq $tempRoot -and [IO.Path]::GetFileName($scratch).StartsWith('ExcelOps-ChartSurvey-')) {
        Remove-Item -LiteralPath $scratch -Recurse -Force
    }
}
