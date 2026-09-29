function Write-SourceCodeTargetFileWithReadonlyFlag {
    param (
        [Parameter(Mandatory)]
        [string] $TargetFile,

        [Parameter(Mandatory)]
        [string[]] $SourceCode,

        [switch] $SetReadOnly
    )

    # Falls Zieldatei existiert und ReadOnly ist → temporär entfernen
    if (Test-Path -LiteralPath $TargetFile) {
        $item = Get-Item -LiteralPath $TargetFile
        if ($item.IsReadOnly) {
            $item.IsReadOnly = $false
        }
    }

    # Normalize generated source files to the repository format without adding a second trailing newline.
    $NormalizedSourceCode = ($SourceCode -join "`r`n") -replace "`r`n|`r|`n", "`r`n"
    $NormalizedSourceCode = $NormalizedSourceCode.TrimEnd("`r", "`n") + "`r`n"
    $Utf8WithBom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Join-Path $PSScriptRoot $TargetFile), $NormalizedSourceCode, $Utf8WithBom)

    # Optional wieder ReadOnly setzen
    if ($SetReadOnly) {
        (Get-Item -LiteralPath $TargetFile).IsReadOnly = $true
    }
}

"TASK 1: RE-USE FILES test-data + TestEnvironment FROM MAIN-TEST TO UNIT-TEST-DEPENDENCY-PROJECTS"

# clone files test-data => FreeSpireXls
remove-item -recurse ExcelOpsTest-FreeSpireXls/test_data
copy-item -recurse ExcelOpsTest/test_data ExcelOpsTest-FreeSpireXls/test_data/

# clone files test-data => SpireXls
remove-item -recurse ExcelOpsTest-SpireXls/test_data
copy-item -recurse ExcelOpsTest/test_data ExcelOpsTest-SpireXls/test_data/

# clone files TestEnvironment => FreeSpireXls
copy-item ExcelOpsTest/TestFiles.vb ExcelOpsTest-SpireXls/
copy-item ExcelOpsTest/TestTools.vb ExcelOpsTest-SpireXls/
copy-item ExcelOpsTest/TestEnvironment.vb ExcelOpsTest-SpireXls/
copy-item ExcelOpsTest/Console.vb ExcelOpsTest-SpireXls/

# clone files TestEnvironment => SpireXls
copy-item ExcelOpsTest/TestFiles.vb ExcelOpsTest-SpireXls/
copy-item ExcelOpsTest/TestTools.vb ExcelOpsTest-SpireXls/
copy-item ExcelOpsTest/TestEnvironment.vb ExcelOpsTest-SpireXls/
copy-item ExcelOpsTest/Console.vb ExcelOpsTest-SpireXls/


"TASK 2: INCLUDE LATEST LOGIC FROM FreeSpireXls edition into SpireXls edition"

# clone files FreeSpireXls => SpireXls
$SourceCode = (gc -Raw ExcelOps-FreeSpireXls/FreeSpireXlsDataOperations.SharedCode.vb) -replace 'Partial Public Class FreeSpireXlsDataOperations', 'Partial Public Class SpireXlsDataOperations' 
$TargetFile = 'ExcelOps-SpireXls/SpireXlsDataOperations.SharedCode.vb'
Write-SourceCodeTargetFileWithReadonlyFlag -TargetFile $TargetFile -SourceCode $SourceCode -SetReadOnly

"TASK 3: INCLUDE LATEST LOGIC FROM EpplusFreeFixCalcsEdition edition into SpireXls edition"

# clone files EpplusFreeFixCalcsEdition => EpplusPolyform
$SourceCode = (gc -Raw ExcelOps-EpplusFreeFixCalcsEdition/EpplusFreeExcelDataOperations.SharedCode.vb) -replace 'Partial Public Class EpplusFreeExcelDataOperations', 'Partial Public Class EpplusPolyformExcelDataOperations' -replace 'CompuMaster.Epplus4', 'OfficeOpenXml' -replace 'EpplusFreeExcelDataOperations', 'EpplusPolyformExcelDataOperations'
$TargetFile = 'ExcelOps-EpplusPolyform/EpplusPolyformExcelDataOperations.SharedCode.vb'
Write-SourceCodeTargetFileWithReadonlyFlag -TargetFile $TargetFile -SourceCode $SourceCode -SetReadOnly

"TASK 4: INCLUDE LATEST LOGIC FROM XlsEpplusFixCalcsEdition edition into XlsEpplusPolyformEdition edition"

# clone files XlsEpplusFixCalcsEdition => XlsEpplusPolyformEdition
$SourceCode = (gc -Raw CM.Data.EpplusFixCalcsEdition/XlsEpplusFixCalcsEdition.vb) -replace 'Public NotInheritable Class XlsEpplusFixCalcsEdition', 'Public NotInheritable Class XlsEpplusPolyformEdition' -replace 'CompuMaster.Epplus4', 'OfficeOpenXml'
$TargetFile = 'CM.Data.EpplusPolyformEdition/XlsEpplusPolyformEdition.vb'
Write-SourceCodeTargetFileWithReadonlyFlag -TargetFile $TargetFile -SourceCode $SourceCode -SetReadOnly

"TASK 5: INCLUDE LATEST UNIT TEST LOGIC FROM CmDataXlsEpplusFixCalcsEditionTest edition into CmDataXlsEpplusPolyformEditionTest edition"

# clong unit test files
$SourceCode = (gc -Raw ExcelOpsTest/Data/CmDataXlsEpplusFixCalcsEditionTest.vb) -replace 'Public Class CmDataXlsEpplusFixCalcsEditionTest', 'Public Class CmDataXlsEpplusPolyformEditionTest' -replace 'XlsEpplusFixCalcsEdition', 'XlsEpplusPolyformEdition'
$TargetFile = 'ExcelOpsTest/Data/CmDataXlsEpplusPolyformEditionTest.vb'
Write-SourceCodeTargetFileWithReadonlyFlag -TargetFile $TargetFile -SourceCode $SourceCode -SetReadOnly

"TASKS COMPLETED."
