param(
    [string] $CatalogPath = (Join-Path $PSScriptRoot 'excel-function-catalog.csv'),
    [string] $FixtureDirectory = (Join-Path (Split-Path -Parent $PSScriptRoot) 'ExcelOpsTest/test_data/FormulaConformance'),
    [string] $OutputPath = (Join-Path $PSScriptRoot 'excel-function-variants.csv')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$catalog = @{}
Import-Csv -LiteralPath $CatalogPath | ForEach-Object { $catalog[$_.Function] = $_ }
$tokens = [regex]::new(@'
"([^"]|"")*"|'([^']|'')*'|(?<function>[A-Za-z_][A-Za-z0-9_.]*)\s*\(|[(){}\[\],]
'@, [Text.RegularExpressions.RegexOptions]::Compiled)
$groups = @{}
$sheetNamespace = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$officeNamespace = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

function Read-Part($Package, [string] $Name) {
    $reader = [IO.StreamReader]::new($Package.GetEntry($Name).Open())
    try { return [xml]$reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Get-ArgumentKind([string] $Text) {
    $value = $Text.Trim()
    if (-not $value) { return 'Omitted' }
    if ($value -eq '""') { return 'EmptyText' }
    if ($value -match '^"([^"]|"")*"$') { return 'Text' }
    if ($value -match '^(TRUE|FALSE)(?:\(\))?$') { return $Matches[1].ToUpperInvariant() }
    if ($value -match '^#[A-Z0-9_/?.]+[!?]?$') { return $value }
    if ($value -match '^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:E[+-]?[0-9]+)?$') {
        $number = 0.0
        if ([double]::TryParse($value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$number) -and
            $number -ge -10 -and $number -le 10 -and [Math]::Truncate($number) -eq $number) {
            return 'Number(' + $number.ToString([Globalization.CultureInfo]::InvariantCulture) + ')'
        }
        return 'Number'
    }

    if ($value.StartsWith('{') -and $value.EndsWith('}')) { return 'ArrayConstant' }
    if ($value -match '^(?:(?<qualifier>.+)!)?\$?[A-Z]{1,3}\$?[0-9]+(?<range>:\$?[A-Z]{1,3}\$?[0-9]+)?(?<spill>#)?$') {
        $qualifier = $Matches['qualifier']
        $scope = if ([regex]::IsMatch([string]$qualifier, '\[[^\]]+\]')) { 'External' } elseif ($qualifier) { 'Sheet' } else { '' }
        $shape = if ($Matches['spill']) { 'SpillReference' } elseif ($Matches['range']) { 'Range' } else { 'Cell' }
        return $scope + $shape
    }
    if ($value -match '^[\p{L}_\\][\p{L}\p{Nd}_.\\]*$') { return 'Name' }
    return 'Expression'
}

function Get-FormulaCalls([string] $Formula) {
    $stack = [Collections.Generic.Stack[object]]::new()
    foreach ($token in $tokens.Matches($Formula)) {
        $group = if ($stack.Count) { $stack.Peek() }
        if ($null -ne $group -and $group.Close -eq ']' -and $token.Value -notin '[', ']') { continue }
        $name = $token.Groups['function']
        if ($name.Success -or $token.Value -in '(', '{', '[') {
            $close = switch ($token.Value) {
                '{' { '}' }
                '[' { ']' }
                default { ')' }
            }
            $stack.Push(@{
                Close = $close
                Start = $token.Index + $token.Length
                Arguments = [Collections.Generic.List[string]]::new()
                Name = $name
            })
            continue
        }
        if ($null -eq $group) { continue }
        if ($token.Value -eq ',' -and $group.Name.Success) {
            $group.Arguments.Add((Get-ArgumentKind ($Formula.Substring($group.Start, $token.Index - $group.Start))))
            $group.Start = $token.Index + $token.Length
            continue
        }
        if ($token.Value -ne $group.Close) { continue }
        [void]$stack.Pop()
        if (-not $group.Name.Success) { continue }
        $last = $Formula.Substring($group.Start, $token.Index - $group.Start)
        if ($group.Arguments.Count -or $last.Trim()) { $group.Arguments.Add((Get-ArgumentKind $last)) }
        $canonical = ($group.Name.Value -replace '^(?:_xlfn\.|_xlws\.)+', '').ToUpperInvariant()
        if (-not $catalog.ContainsKey($canonical)) { continue }
        [pscustomobject]@{
            Function = $canonical
            Direct = $group.Name.Index -eq 0 -and $token.Index + $token.Length -eq $Formula.Length
            Arity = $group.Arguments.Count
            ArgumentKinds = $group.Arguments -join ' | '
        }
    }
}

function Get-CachedResult($Cell, $Formula) {
    $range = $Formula.GetAttribute('ref')
    $shape = if ($Cell.HasAttribute('cm')) {
        if ($range.Contains(':')) { 'Spill' } else { 'DynamicScalar' }
    } elseif ($Formula.GetAttribute('t') -eq 'array') { 'LegacyArray' }
    else { 'Scalar' }
    $value = $Cell['v']
    $type = if ($null -eq $value) { 'Missing' } else {
        switch ($Cell.GetAttribute('t')) {
            'e' { 'Error' }
            'b' { 'Boolean' }
            { $_ -in 'str', 's', 'inlineStr' } { 'Text' }
            default { 'Number' }
        }
    }
    return @{
        Shape = $shape
        Type = $type
        Error = if ($type -eq 'Error') { $value.InnerText } else { '' }
        Range = $range
    }
}

foreach ($file in Get-ChildItem -LiteralPath $FixtureDirectory -Filter '*.xlsx' | Sort-Object -Property Name) {
    $package = [IO.Compression.ZipFile]::OpenRead($file.FullName)
    try {
        $workbook = Read-Part $package 'xl/workbook.xml'
        $date1904 = $workbook.DocumentElement['workbookPr']?.GetAttribute('date1904') -in '1', 'true'
        $dateSystem = if ($date1904) { '1904' } else { '1900' }
        $relationships = @{}
        foreach ($relationship in (Read-Part $package 'xl/_rels/workbook.xml.rels').DocumentElement.ChildNodes) {
            $relationships[$relationship.Id] = [Uri]::new([Uri]'https://xlsx.local/xl/workbook.xml', $relationship.Target).AbsolutePath.TrimStart('/')
        }
        foreach ($sheet in $workbook.GetElementsByTagName('sheet', $sheetNamespace)) {
            $path = $relationships[$sheet.GetAttribute('id', $officeNamespace)]
            if ($path -notmatch '^xl/worksheets/') { continue }
            $worksheet = Read-Part $package $path
            foreach ($formula in $worksheet.GetElementsByTagName('f', $sheetNamespace)) {
                $cell = $formula.ParentNode
                $cached = Get-CachedResult $cell $formula
                foreach ($call in Get-FormulaCalls $formula.InnerText.Trim()) {
                    $context = if ($call.Direct) { 'Direct' } else { 'Nested' }
                    $shape = if ($call.Direct) { $cached.Shape } else { '' }
                    $type = if ($call.Direct) { $cached.Type } else { '' }
                    $errorValue = if ($call.Direct) { $cached.Error } else { '' }
                    $key = $call.Function, $context, $call.Arity, $call.ArgumentKinds, $shape, $type, $errorValue, $dateSystem -join ';'
                    if ($groups.ContainsKey($key)) {
                        $groups[$key].Cases++
                        continue
                    }
                    $groups[$key] = [pscustomobject][ordered]@{
                        Function = $call.Function
                        Context = $context
                        Arity = $call.Arity
                        ArgumentKinds = $call.ArgumentKinds
                        ResultShape = $shape
                        CachedAnchorType = $type
                        CachedAnchorError = $errorValue
                        DateSystem = $dateSystem
                        ExampleWorkbook = $file.Name
                        ExampleSheet = $sheet.Name
                        ExampleCell = $cell.GetAttribute('r')
                        ExampleResultRange = if ($call.Direct) { $cached.Range } else { '' }
                        Cases = 1
                    }
                }
            }
        }
    } finally { $package.Dispose() }
}

[string[]] $keys = $groups.Keys
[Array]::Sort($keys, [StringComparer]::Ordinal)
$csv = $keys | ForEach-Object { $groups[$_] } | ConvertTo-Csv -NoTypeInformation
[IO.File]::WriteAllText([IO.Path]::GetFullPath($OutputPath), (($csv -join "`r`n") + "`r`n"), [Text.UTF8Encoding]::new($true))
$functionCount = @($groups.Values.Function | Sort-Object -Unique).Count
"Recorded $($groups.Count) formula variants across $functionCount cataloged functions."
