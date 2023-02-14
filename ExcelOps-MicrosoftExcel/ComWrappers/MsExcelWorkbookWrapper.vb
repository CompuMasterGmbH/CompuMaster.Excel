Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.MsExcelCom

    Public Class MsExcelWorkbookWrapper
        Inherits CompuMaster.ComInterop.ComChildObject(Of MsExcelWorkbooksWrapper, MsExcel.Workbook)

        Public Sub New(parent As MsExcelWorkbooksWrapper, obj As MsExcel.Workbook)
            MyBase.New(parent, obj)
        End Sub

    End Class

End Namespace