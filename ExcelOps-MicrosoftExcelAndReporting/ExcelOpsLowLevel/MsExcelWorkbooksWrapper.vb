Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.ExcelOps

    Public Class MsExcelWorkbooksWrapper
        Inherits CompuMaster.ComInterop.ComChildObject(Of MsExcelApplicationWrapper, MsExcel.Workbooks)

        Public Sub New(parent As MsExcelApplicationWrapper, obj As MsExcel.Workbooks)
            MyBase.New(parent, obj)
        End Sub

    End Class

End Namespace