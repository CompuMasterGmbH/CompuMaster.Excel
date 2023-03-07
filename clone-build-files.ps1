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
$SourceCode = (gc ExcelOps-FreeSpireXls/FreeSpireXlsDataOperations.SharedCode.vb) -replace 'Partial Public Class FreeSpireXlsDataOperations', 'Partial Public Class SpireXlsDataOperations' 
if ($PSVersionTable.PSVersion.Major -ge 7) 
{ 
    # "pwsh" -> force encoding UTF8BOM
    $SourceCode | Out-File -encoding UTF8BOM ExcelOps-SpireXls/SpireXlsDataOperations.SharedCode.vb
} 
else 
{ 
    # "powershell" -> UTF8BOM not supported, but BOM added by default for UTF8
    $SourceCode | Out-File -encoding UTF8 ExcelOps-SpireXls/SpireXlsDataOperations.SharedCode.vb
}

"TASK 3: INCLUDE LATEST LOGIC FROM EpplusFreeFixCalcsEdition edition into SpireXls edition"

# clone files EpplusFreeFixCalcsEdition => EpplusPolyform
$SourceCode = (gc ExcelOps-EpplusFreeFixCalcsEdition/EpplusFreeExcelDataOperations.SharedCode.vb) -replace 'Partial Public Class EpplusFreeExcelDataOperations', 'Partial Public Class EpplusPolyformExcelDataOperations' -replace 'CompuMaster.Epplus4', 'OfficeOpenXml' -replace 'EpplusFreeExcelDataOperations', 'EpplusPolyformExcelDataOperations'
if ($PSVersionTable.PSVersion.Major -ge 7) 
{ 
    # "pwsh" -> force encoding UTF8BOM
    $SourceCode | Out-File -encoding UTF8BOM ExcelOps-EpplusPolyform/EpplusPolyformExcelDataOperations.SharedCode.vb
} 
else 
{ 
    # "powershell" -> UTF8BOM not supported, but BOM added by default for UTF8
    $SourceCode | Out-File -encoding UTF8 ExcelOps-EpplusPolyform/EpplusPolyformExcelDataOperations.SharedCode.vb
}

"TASK 4: INCLUDE LATEST LOGIC FROM XlsEpplusFixCalcsEdition edition into XlsEpplusPolyformEdition edition"

# clone files XlsEpplusFixCalcsEdition => XlsEpplusPolyformEdition
$SourceCode = (gc CM.Data.EpplusFixCalcsEdition/XlsEpplusFixCalcsEdition.vb) -replace 'Public Class XlsEpplusFixCalcsEdition', 'Public Class XlsEpplusPolyformEdition' -replace 'CompuMaster.Epplus4', 'OfficeOpenXml'
if ($PSVersionTable.PSVersion.Major -ge 7) 
{ 
    # "pwsh" -> force encoding UTF8BOM
    $SourceCode | Out-File -encoding UTF8BOM CM.Data.EpplusPolyformEdition/XlsEpplusPolyformEdition.vb
} 
else 
{ 
    # "powershell" -> UTF8BOM not supported, but BOM added by default for UTF8
    $SourceCode | Out-File -encoding UTF8 CM.Data.EpplusPolyformEdition/XlsEpplusPolyformEdition.vb
}

"TASK 5: INCLUDE LATEST UNIT TEST LOGIC FROM CmDataXlsEpplusFixCalcsEditionTest edition into CmDataXlsEpplusPolyformEditionTest edition"

# clong unit test files
$SourceCode = (gc ExcelOpsTest/Data/CmDataXlsEpplusFixCalcsEditionTest.vb) -replace 'Public Class CmDataXlsEpplusFixCalcsEditionTest', 'Public Class CmDataXlsEpplusPolyformEditionTest' -replace 'XlsEpplusFixCalcsEdition', 'XlsEpplusPolyformEdition'
if ($PSVersionTable.PSVersion.Major -ge 7) 
{ 
    # "pwsh" -> force encoding UTF8BOM
    $SourceCode | Out-File -encoding UTF8BOM ExcelOpsTest/Data/CmDataXlsEpplusPolyformEditionTest.vb
} 
else 
{ 
    # "powershell" -> UTF8BOM not supported, but BOM added by default for UTF8
    $SourceCode | Out-File -encoding UTF8 ExcelOpsTest/Data/CmDataXlsEpplusPolyformEditionTest.vb
}


"TASKS COMPLETED."