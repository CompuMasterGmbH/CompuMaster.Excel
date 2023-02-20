"TASK 1: RE-USE FILES test-data + TestEnvironment FORM MAIN-TEST TO DEPENDENCIES"

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

(gc ExcelOps-FreeSpireXls/FreeSpireXlsDataOperations.SharedCode.vb) -replace 'Partial Public Class FreeSpireXlsDataOperations', 'Partial Public Class SpireXlsDataOperations' | Out-File -encoding UTF8 ExcelOps-SpireXls/SpireXlsDataOperations.SharedCode.vb

"TASKS COMPLETED."