param(
    [string] $SupportIndexPath,
    [string] $XlsxFunctionsPath,
    [string] $OutputPath
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
if (-not $OutputPath) { $OutputPath = Join-Path $PSScriptRoot 'excel-function-catalog.csv' }

$supportPage = [Uri]'https://support.microsoft.com/en-us/excel/excel-functions-alphabetical'
$xlsxPage = [Uri]'https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/5d1b6d44-6fc1-4ecd-8fef-0b27406cc2bf'

function Read-Source([string] $LocalPath, [Uri] $Address) {
    if ($LocalPath) { return Get-Content -LiteralPath $LocalPath -Raw -Encoding UTF8 }
    return (Invoke-WebRequest -UseBasicParsing -Uri $Address -TimeoutSec 30).Content
}

function Get-PlainText([string] $Html) {
    return [Net.WebUtility]::HtmlDecode(($Html -replace '<[^>]+>', '')).Trim()
}

function Convert-FunctionName([string] $Value) {
    return ($Value -replace '^(?:_xlfn\.|_xlws\.)+', '').ToUpperInvariant()
}

$supportHtml = Read-Source $SupportIndexPath $supportPage
$xlsxHtml = Read-Source $XlsxFunctionsPath $xlsxPage
$documentedTokens = @{}
$futureTable = [regex]::Match($xlsxHtml, '(?is)<table\b[^>]*>(?:(?!</table>).)*Future function(?:(?!</table>).)*</table>')
if (-not $futureTable.Success) { throw 'The Microsoft XLSX future-function table could not be located.' }
foreach ($match in [regex]::Matches($futureTable.Value, '(?is)<td\b[^>]*>(.*?)</td>')) {
    $token = Get-PlainText $match.Groups[1].Value
    if ($token -cmatch '^(?:(?:_xlfn\.|_xlws\.))*[A-Z][A-Z0-9_.]*$') {
        $documentedTokens[(Convert-FunctionName $token)] = $token
    }
}
if ($documentedTokens.Count -lt 100) { throw 'The XLSX future-function table contains fewer entries than expected.' }

$implementation = @{}
$operations = Get-Content -LiteralPath (Join-Path $root 'ExcelOps-OpenXml/OpenXmlExcelDataOperations.FormulaOperations.vb') -Raw
$dispatch = [regex]::Matches($operations, '(?s)(?:Friend Function EvaluateFunction|Private Function EvaluateBuiltInFunction)\(.*?End Function')
foreach ($method in $dispatch) {
    foreach ($case in [regex]::Matches($method.Value, '(?m)^\s*(?:Case\s+|If upperName = )("[^"]+"(?:\s*,\s*"[^"]+")*)')) {
        foreach ($name in [regex]::Matches($case.Groups[1].Value, '"([^"]+)"')) {
            $implementation[(Convert-FunctionName $name.Groups[1].Value)] = $true
        }
    }
}
$parser = Get-Content -LiteralPath (Join-Path $root 'ExcelOps-OpenXml/OpenXmlExcelDataOperations.FormulaParser.vb') -Raw
foreach ($name in [regex]::Matches($parser, 'token\.Equals\("([A-Za-z0-9_.]+)"[^\r\n]*StringComparison\.OrdinalIgnoreCase')) {
    if ((Convert-FunctionName $name.Groups[1].Value) -in 'LET','LAMBDA') {
        $implementation[(Convert-FunctionName $name.Groups[1].Value)] = $true
    }
}
if ($implementation.PSBase.Count -lt 30) { throw 'The Open XML function dispatch could not be read.' }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$sampleCounts = @{}
$observedTokens = @{}
$fixtureDirectory = Join-Path $root 'ExcelOpsTest/test_data/FormulaConformance'
foreach ($file in Get-ChildItem -LiteralPath $fixtureDirectory -Filter '*.xlsx' | Sort-Object Name) {
    $package = [IO.Compression.ZipFile]::OpenRead($file.FullName)
    try {
        foreach ($sheet in $package.Entries.Where({ $_.FullName -match '^xl/worksheets/sheet\d+\.xml$' })) {
            $reader = [IO.StreamReader]::new($sheet.Open())
            try { $xml = [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
            foreach ($formula in $xml.SelectNodes("//*[local-name()='c']/*[local-name()='f']")) {
                $text = $formula.InnerText -replace '"(?:[^"]|"")*"', '' -replace "'(?:[^']|'')*'", ''
                $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                foreach ($call in [regex]::Matches($text, '(?i)(?<![A-Z0-9_.])([A-Z_][A-Z0-9_.]*)\s*\(')) {
                    $token = $call.Groups[1].Value
                    $name = Convert-FunctionName $token
                    if ($seen.Add($name)) { $sampleCounts[$name] = 1 + [int]$sampleCounts[$name] }
                    if ($token.StartsWith('_xlfn.', [StringComparison]::OrdinalIgnoreCase)) { $observedTokens[$name] = $token }
                }
            }
        }
    } finally { $package.Dispose() }
}

$table = [regex]::Match($supportHtml, '(?is)<table\b[^>]*>(?:(?!</table>).)*Function name(?:(?!</table>).)*</table>')
if (-not $table.Success) { throw 'The Microsoft Excel worksheet-function index could not be located.' }

$index = @{}
foreach ($row in [regex]::Matches($table.Value, '(?is)<tr\b[^>]*>(.*?)</tr>')) {
    $cells = [regex]::Matches($row.Groups[1].Value, '(?is)<td\b[^>]*>(.*?)</td>')
    if ($cells.Count -ne 2) { continue }
    $label = $cells[0].Groups[1].Value
    $anchor = [regex]::Match($label, '(?is)<a\b[^>]*\bhref="([^"]+)"[^>]*>(.*?)</a>')
    if (-not $anchor.Success) { continue }
    $url = [Uri]::new($supportPage, [Net.WebUtility]::HtmlDecode($anchor.Groups[1].Value))
    if ($url.Scheme -ne 'https' -or $url.Host -ne 'support.microsoft.com') { continue }
    $version = [regex]::Match((Get-PlainText $label), '\((\d{4}|Microsoft 365|Office 365)\)').Groups[1].Value
    if (-not $version -and $label -match 'alt="[^"]*(?:Office 365|Microsoft 365)[^"]*"') { $version = 'Microsoft 365' }
    $category = (Get-PlainText ([regex]::Match($cells[1].Groups[1].Value, '(?is)<strong>(.*?)</strong>').Groups[1].Value)).TrimEnd(':')
    $names = (Get-PlainText $anchor.Groups[2].Value) -split ','
    foreach ($displayName in $names) {
        $name = $displayName.Trim()
        # The Microsoft index currently adds a stray letter to the documented BETA.INV name.
        if ($name -ceq 'BETA.INVn' -and $url.AbsolutePath.EndsWith('/beta-inv-function')) { $name = 'BETA.INV' }
        if ($name -cnotmatch '^[A-Z][A-Z0-9_.]*$') { throw "Unexpected worksheet-function name '$name' in the Microsoft index." }
        $index[$name] = [pscustomobject][ordered]@{
            Function = $name
            Category = $category
            Introduced = $version
            StorageToken = if ($observedTokens.ContainsKey($name)) { $observedTokens[$name] } elseif ($documentedTokens.ContainsKey($name)) { $documentedTokens[$name] } else { '' }
            OpenXmlEvaluator = if ($implementation.ContainsKey($name)) { 'Present' } else { '' }
            ReferenceFormulaCells = [int]$sampleCounts[$name]
            MicrosoftReference = $url.AbsoluteUri
        }
    }
}
if ($index.PSBase.Count -lt 480) { throw "Only $($index.PSBase.Count) worksheet functions could be read from the Microsoft index." }

[string[]] $functionNames = $index.PSBase.Keys
[Array]::Sort($functionNames, [StringComparer]::Ordinal)
$rows = @($functionNames | ForEach-Object { $index[$_] })
$csv = $rows | ConvertTo-Csv -NoTypeInformation
[IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputPath), (($csv -join [Environment]::NewLine) + [Environment]::NewLine), [Text.UTF8Encoding]::new($true))
Write-Output "Catalogued $($rows.Count) worksheet functions. $(@($rows.Where({$_.OpenXmlEvaluator})).Count) evaluator entries and $(@($rows.Where({$_.ReferenceFormulaCells -gt 0})).Count) functions present in reference formula cells."
