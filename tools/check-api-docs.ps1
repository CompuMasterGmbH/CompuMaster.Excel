param(
    [switch] $Strict
)

$ErrorActionPreference = "Stop"

$repoRoot = Resolve-Path (Join-Path $PSScriptRoot "..")

$sourceRoots = @(
    "ExcelOps",
    "ExcelOps-EpplusFreeFixCalcsEdition",
    "ExcelOps-EpplusPolyform",
    "ExcelOps-FreeSpireXls",
    "ExcelOps-SpireXls",
    "ExcelOps-MicrosoftExcel",
    "ExcelOps-Tools-MsAndEpplusFreeFixCalcsEdition",
    "MsExcelComInterop",
    "CM.Data.EpplusFixCalcsEdition",
    "CM.Data.EpplusPolyformEdition"
)

$excludedDocumentationScopes = @(
    "Epplus-FixCalcsEdition/EPPlus/",
    "TestAndDemoExcelOps/"
)

$maxMissingDocumentation = 0
$maxOverridesWithoutInheritdoc = 0

if (-not $Strict) {
    $maxMissingDocumentation = 0
    $maxOverridesWithoutInheritdoc = 0
}

function Get-RelativePath([string] $path) {
    return [System.IO.Path]::GetRelativePath($repoRoot, (Resolve-Path $path))
}

function Get-XmlDocBlock([string[]] $lines, [int] $declarationIndex) {
    $index = $declarationIndex - 1

    while ($index -ge 0) {
        $trimmed = $lines[$index].Trim()
        if ($trimmed.Length -eq 0 -or $trimmed.StartsWith("<")) {
            $index--
            continue
        }

        break
    }

    $docLines = New-Object System.Collections.Generic.List[string]
    while ($index -ge 0 -and $lines[$index].TrimStart().StartsWith("'''")) {
        $docLines.Insert(0, $lines[$index])
        $index--
    }

    return $docLines
}

function Is-Documented([System.Collections.Generic.List[string]] $docBlock) {
    return $docBlock.Count -gt 0
}

function Test-Inheritdoc([System.Collections.Generic.List[string]] $docBlock) {
    return (($docBlock -join "`n") -match "<inheritdoc\s*/>")
}

function Get-DeclarationWithoutInlineAttributes([string] $line) {
    $declaration = $line.Trim()

    while ($declaration -match '^<[^>]+>\s*') {
        $declaration = $declaration.Substring($matches[0].Length).TrimStart()
    }

    return $declaration
}

$files = New-Object System.Collections.Generic.List[string]
foreach ($sourceRoot in $sourceRoots) {
    $rootPath = Join-Path $repoRoot $sourceRoot
    if (Test-Path -LiteralPath $rootPath) {
        Get-ChildItem -LiteralPath $rootPath -Recurse -Filter "*.vb" -File |
            Where-Object {
                $relativeFile = (Get-RelativePath $_.FullName).Replace("\", "/")
                $_.FullName -notmatch "\\(bin|obj)\\" -and
                    -not ($excludedDocumentationScopes | Where-Object { $relativeFile.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase) })
            } |
            ForEach-Object { $files.Add($_.FullName) }
    }
}

$memberModifierPattern = '(?:Shared|Overrides|Overridable|MustOverride|MustInherit|NotInheritable|ReadOnly|WriteOnly|Partial|Default|Shadows|Overloads|Widening|Narrowing|Custom|Async|Iterator|Declare|Auto|Ansi|Unicode)'
$declarationPattern = "^(Public|Protected Friend|Protected)\s+(?:$memberModifierPattern\s+)*(Class|Structure|Enum|Interface|Module|Delegate|Event|Property|Function|Sub|Operator)\b"
$fieldPattern = "^(Public|Protected Friend|Protected)\s+(?:(?:Shared|ReadOnly|Const|WithEvents|Shadows)\s+)*[A-Za-z_][A-Za-z0-9_]*(?:\([^)]*\))?\s*(?:,|As\b|=)"

$missingDocumentation = New-Object System.Collections.Generic.List[object]
$overridesWithoutInheritdoc = New-Object System.Collections.Generic.List[object]

foreach ($file in $files) {
    $relativeFile = Get-RelativePath $file
    $lines = Get-Content -LiteralPath $file
    $insidePublicOrProtectedEnum = $false

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $declarationLine = Get-DeclarationWithoutInlineAttributes $line

        if ($insidePublicOrProtectedEnum) {
            if ($declarationLine -match "^End\s+Enum\b") {
                $insidePublicOrProtectedEnum = $false
                continue
            }

            if ($declarationLine.Length -gt 0 -and
                -not $declarationLine.StartsWith("'") -and
                $declarationLine -match "^[A-Za-z_][A-Za-z0-9_]*\b") {

                $docBlock = Get-XmlDocBlock $lines $i
                if (-not (Is-Documented $docBlock)) {
                    $missingDocumentation.Add([pscustomobject]@{
                        File = $relativeFile
                        Line = $i + 1
                        Declaration = $declarationLine
                    })
                }
            }
        }

        if ($declarationLine -notmatch $declarationPattern -and $declarationLine -notmatch $fieldPattern) {
            continue
        }

        $docBlock = Get-XmlDocBlock $lines $i
        $declaration = $declarationLine

        if (-not (Is-Documented $docBlock)) {
            $missingDocumentation.Add([pscustomobject]@{
                File = $relativeFile
                Line = $i + 1
                Declaration = $declaration
            })
            continue
        }

        if ($declaration -match "\bOverrides\b" -and -not (Test-Inheritdoc $docBlock)) {
            $overridesWithoutInheritdoc.Add([pscustomobject]@{
                File = $relativeFile
                Line = $i + 1
                Declaration = $declaration
            })
        }

        if ($declaration -match "\bEnum\b") {
            $insidePublicOrProtectedEnum = $true
        }
    }
}

Write-Host "API documentation check"
Write-Host "Missing XML documentation: $($missingDocumentation.Count) (allowed: $maxMissingDocumentation)"
Write-Host "Overrides without inheritdoc: $($overridesWithoutInheritdoc.Count) (allowed: $maxOverridesWithoutInheritdoc)"

$failed = $false

if ($missingDocumentation.Count -gt $maxMissingDocumentation) {
    $failed = $true
    Write-Host ""
    Write-Host "Public/protected API members without XML documentation:"
    $missingDocumentation | Select-Object -First 50 | Format-Table -AutoSize | Out-String | Write-Host
}

if ($overridesWithoutInheritdoc.Count -gt $maxOverridesWithoutInheritdoc) {
    $failed = $true
    Write-Host ""
    Write-Host "Overrides with XML documentation but without <inheritdoc/>:"
    $overridesWithoutInheritdoc | Select-Object -First 50 | Format-Table -AutoSize | Out-String | Write-Host
}

if ($Strict -and ($missingDocumentation.Count -gt 0 -or $overridesWithoutInheritdoc.Count -gt 0)) {
    $failed = $true
}

if ($failed) {
    throw "API documentation check failed. Add XML documentation or use <inheritdoc/> for matching overrides."
}
