Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.MsExcelCom

    Public Class MsExcelWorkbooksWrapper
        Inherits CompuMaster.ComInterop.ComChildObject(Of MsExcelApplicationWrapper, MsExcel.Workbooks)

        Public Sub New(parent As MsExcelApplicationWrapper, obj As MsExcel.Workbooks)
            MyBase.New(parent, obj)
        End Sub

        Public Function Add() As MsExcelWorkbookWrapper
            Return New MsExcelWorkbookWrapper(Me, Me.ComObjectStronglyTyped.Add())
        End Function

        Public Function Open(path As String, [readOnly] As Boolean, passwordForOpening As String) As MsExcelWorkbookWrapper
            Return New MsExcelWorkbookWrapper(Me, Me.ComObjectStronglyTyped.Open(path, False, [readOnly], Nothing, If(passwordForOpening = Nothing, Nothing, passwordForOpening)))
        End Function

        Public Function Workbook(index As Integer) As MsExcelWorkbookWrapper
            Return New MsExcelWorkbookWrapper(Me, Me.ComObjectStronglyTyped.Item(index))
        End Function

        Public Function Workbook(name As String) As MsExcelWorkbookWrapper
            Return New MsExcelWorkbookWrapper(Me, Me.ComObjectStronglyTyped.Item(name))
        End Function

    End Class

End Namespace