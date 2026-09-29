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
    return (($docBlock -join "`n") -match "<inheritdoc(?:\s[^>]*)?\s*/>")
}

function Get-NormalizedXmlDocumentation([System.Collections.Generic.List[string]] $docBlock) {
    return (($docBlock | ForEach-Object { $_ -replace '^\s*''{3}\s?', '' }) -join "`n")
}

function Test-XmlContentEmpty([string] $content) {
    if ($content -match '<(?:see|paramref|typeparamref)\b') {
        return $false
    }

    return ([System.Text.RegularExpressions.Regex]::Replace($content, '<[^>]+>', '')).Trim().Length -eq 0
}

function Get-CompleteDeclaration([string[]] $lines, [int] $declarationIndex, [string] $firstLine) {
    $declaration = $firstLine
    $openParentheses = ([System.Text.RegularExpressions.Regex]::Matches($firstLine, '\(')).Count
    $closeParentheses = ([System.Text.RegularExpressions.Regex]::Matches($firstLine, '\)')).Count
    $index = $declarationIndex + 1

    while ($index -lt $lines.Count -and ($openParentheses -gt $closeParentheses -or $declaration.TrimEnd().EndsWith('_'))) {
        $nextLine = $lines[$index].Trim()
        $declaration += ' ' + $nextLine
        $openParentheses += ([System.Text.RegularExpressions.Regex]::Matches($nextLine, '\(')).Count
        $closeParentheses += ([System.Text.RegularExpressions.Regex]::Matches($nextLine, '\)')).Count
        $index++
    }

    return ($declaration -replace '\s+_\s+', ' ' -replace '\s+', ' ').Trim()
}

function Get-TopLevelParenthesisGroups([string] $declaration) {
    $groups = New-Object System.Collections.Generic.List[string]
    $depth = 0
    $start = -1

    for ($index = 0; $index -lt $declaration.Length; $index++) {
        $character = $declaration[$index]
        if ($character -eq '(') {
            if ($depth -eq 0) {
                $start = $index + 1
            }
            $depth++
        } elseif ($character -eq ')') {
            $depth--
            if ($depth -eq 0 -and $start -ge 0) {
                $groups.Add($declaration.Substring($start, $index - $start))
                $start = -1
            }
        }
    }

    return $groups
}

function Split-TopLevelCommaSeparated([string] $value) {
    $items = New-Object System.Collections.Generic.List[string]
    $depth = 0
    $start = 0

    for ($index = 0; $index -lt $value.Length; $index++) {
        $character = $value[$index]
        if ($character -eq '(') {
            $depth++
        } elseif ($character -eq ')') {
            $depth--
        } elseif ($character -eq ',' -and $depth -eq 0) {
            $items.Add($value.Substring($start, $index - $start).Trim())
            $start = $index + 1
        }
    }

    if ($start -lt $value.Length) {
        $items.Add($value.Substring($start).Trim())
    }

    return $items
}

function Get-DeclarationParameterNames([string] $declaration) {
    $groups = @(Get-TopLevelParenthesisGroups $declaration)
    if ($groups.Count -eq 0) {
        return @()
    }

    $parameterGroupIndex = 0
    if ($groups[0].TrimStart() -match '^Of\s+') {
        $parameterGroupIndex = 1
    }
    if ($parameterGroupIndex -ge $groups.Count -or [string]::IsNullOrWhiteSpace($groups[$parameterGroupIndex])) {
        return @()
    }

    $parameterNames = New-Object System.Collections.Generic.List[string]
    foreach ($parameter in (Split-TopLevelCommaSeparated $groups[$parameterGroupIndex])) {
        $normalizedParameter = $parameter -replace '^\s*<[^>]+>\s*', ''
        if ($normalizedParameter -match '^(?:(?:Optional|ByVal|ByRef|ParamArray)\s+)*(\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_]*)(?=\s|As\b|=|,|$)') {
            $parameterNames.Add($matches[1].Trim('[', ']'))
        }
    }

    return $parameterNames
}

function Get-DeclarationTypeParameterNames([string] $declaration) {
    if ($declaration -notmatch '^(?:Public|Protected(?:\s+Friend)?)\s+(?:(?:Shared|Overrides|Overridable|MustOverride|MustInherit|NotInheritable|ReadOnly|WriteOnly|Partial|Default|Shadows|Overloads|Widening|Narrowing|Custom|Async|Iterator|Declare|Auto|Ansi|Unicode)\s+)*(?:Class|Structure|Interface|Delegate|Function|Sub|Property)\s+(?:\[[^\]]+\]|[A-Za-z_][A-Za-z0-9_]*)\s*\(Of\s+') {
        return @()
    }

    $groups = @(Get-TopLevelParenthesisGroups $declaration)
    if ($groups.Count -eq 0 -or $groups[0].TrimStart() -notmatch '^Of\s+') {
        return @()
    }

    $typeParameterNames = New-Object System.Collections.Generic.List[string]
    $typeParameterList = $groups[0].Trim() -replace '^Of\s+', ''
    foreach ($typeParameter in (Split-TopLevelCommaSeparated $typeParameterList)) {
        if ($typeParameter -match '^([A-Za-z_][A-Za-z0-9_]*)\b') {
            $typeParameterNames.Add($matches[1])
        }
    }

    return $typeParameterNames
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
$incompleteDocumentation = New-Object System.Collections.Generic.List[object]

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

        if (-not (Test-Inheritdoc $docBlock)) {
            $completeDeclaration = Get-CompleteDeclaration $lines $i $declaration
            $xmlDocumentation = Get-NormalizedXmlDocumentation $docBlock
            $declarationParameters = @(Get-DeclarationParameterNames $completeDeclaration)
            $documentedParameters = @([System.Text.RegularExpressions.Regex]::Matches($xmlDocumentation, '<param\s+name="([^"]+)"[^>]*>(.*?)</param>', 'IgnoreCase, Singleline'))
            $documentedParameterNames = @($documentedParameters | ForEach-Object { $_.Groups[1].Value.Trim('[', ']') })

            foreach ($parameterName in $declarationParameters) {
                if ($documentedParameterNames -notcontains $parameterName) {
                    $incompleteDocumentation.Add([pscustomobject]@{ File = $relativeFile; Line = $i + 1; Issue = "Missing <param> for '$parameterName'" })
                }
            }
            foreach ($parameterDocumentation in $documentedParameters) {
                $parameterName = $parameterDocumentation.Groups[1].Value.Trim('[', ']')
                if ($declarationParameters -notcontains $parameterName) {
                    $incompleteDocumentation.Add([pscustomobject]@{ File = $relativeFile; Line = $i + 1; Issue = "Unknown <param> '$parameterName'" })
                } elseif (Test-XmlContentEmpty $parameterDocumentation.Groups[2].Value) {
                    $incompleteDocumentation.Add([pscustomobject]@{ File = $relativeFile; Line = $i + 1; Issue = "Empty <param> for '$parameterName'" })
                }
            }

            $declarationTypeParameters = @(Get-DeclarationTypeParameterNames $completeDeclaration)
            $documentedTypeParameters = @([System.Text.RegularExpressions.Regex]::Matches($xmlDocumentation, '<typeparam\s+name="([^"]+)"[^>]*>(.*?)</typeparam>', 'IgnoreCase, Singleline'))
            $documentedTypeParameterNames = @($documentedTypeParameters | ForEach-Object { $_.Groups[1].Value })
            foreach ($typeParameterName in $declarationTypeParameters) {
                if ($documentedTypeParameterNames -notcontains $typeParameterName) {
                    $incompleteDocumentation.Add([pscustomobject]@{ File = $relativeFile; Line = $i + 1; Issue = "Missing <typeparam> for '$typeParameterName'" })
                }
            }
            foreach ($typeParameterDocumentation in $documentedTypeParameters) {
                if (Test-XmlContentEmpty $typeParameterDocumentation.Groups[2].Value) {
                    $incompleteDocumentation.Add([pscustomobject]@{ File = $relativeFile; Line = $i + 1; Issue = "Empty <typeparam> for '$($typeParameterDocumentation.Groups[1].Value)'" })
                }
            }

            foreach ($tagName in @('summary', 'returns', 'value', 'remarks')) {
                foreach ($tag in [System.Text.RegularExpressions.Regex]::Matches($xmlDocumentation, "<$tagName(?:\s[^>]*)?>(.*?)</$tagName>", 'IgnoreCase, Singleline')) {
                    if (Test-XmlContentEmpty $tag.Groups[1].Value) {
                        $incompleteDocumentation.Add([pscustomobject]@{ File = $relativeFile; Line = $i + 1; Issue = "Empty <$tagName>" })
                    }
                }
            }

            if ($completeDeclaration -match '^(?:Public|Protected(?:\s+Friend)?)\s+(?:(?:' + $memberModifierPattern + ')\s+)*Function\b' -and
                $xmlDocumentation -notmatch '<returns(?:\s[^>]*)?>') {
                $incompleteDocumentation.Add([pscustomobject]@{ File = $relativeFile; Line = $i + 1; Issue = 'Missing <returns>' })
            }
        }

        if ($declaration -match "\bEnum\b") {
            $insidePublicOrProtectedEnum = $true
        }
    }
}

Write-Host "API documentation check"
Write-Host "Missing XML documentation: $($missingDocumentation.Count) (allowed: $maxMissingDocumentation)"
Write-Host "Overrides without inheritdoc: $($overridesWithoutInheritdoc.Count) (allowed: $maxOverridesWithoutInheritdoc)"
Write-Host "Incomplete XML documentation: $($incompleteDocumentation.Count) (allowed: 0)"

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

if ($incompleteDocumentation.Count -gt 0) {
    $failed = $true
    Write-Host ""
    Write-Host "Public/protected API members with incomplete XML documentation:"
    $incompleteDocumentation | Select-Object -First 100 | Format-Table -AutoSize | Out-String | Write-Host
}

if ($Strict -and ($missingDocumentation.Count -gt 0 -or $overridesWithoutInheritdoc.Count -gt 0 -or $incompleteDocumentation.Count -gt 0)) {
    $failed = $true
}

if ($failed) {
    throw "API documentation check failed. Add XML documentation or use <inheritdoc/> for matching overrides."
}
