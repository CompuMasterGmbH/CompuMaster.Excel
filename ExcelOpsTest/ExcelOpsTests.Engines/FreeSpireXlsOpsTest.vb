Public Class FreeSpireXlsOpsTest
    Inherits ExcelOpsTestBase(Of ExcelOps.FreeSpireXlsDataOperations)

    Protected Overrides Function CreateInstance(file As String, mode As ExcelOps.ExcelDataOperationsBase.OpenMode, [readOnly] As Boolean, passwordForOpening As String) As ExcelOps.FreeSpireXlsDataOperations
        Return New ExcelOps.FreeSpireXlsDataOperations(file, mode, [readOnly], passwordForOpening)
    End Function

    Protected Overrides Function CreateInstance() As ExcelOps.FreeSpireXlsDataOperations
        Return New ExcelOps.FreeSpireXlsDataOperations()
    End Function

End Class