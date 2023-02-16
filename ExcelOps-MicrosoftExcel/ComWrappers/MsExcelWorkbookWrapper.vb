Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.MsExcelCom

    ''' <summary>
    ''' A COM wrapper class for a MS Excel workbook
    ''' </summary>
    Public Class MsExcelWorkbookWrapper
        Inherits CompuMaster.ComInterop.ComChildObject(Of MsExcelWorkbooksWrapper, MsExcel.Workbook)

        Public Sub New(parent As MsExcelWorkbooksWrapper, obj As MsExcel.Workbook)
            MyBase.New(parent, obj)
        End Sub

        ''' <summary>
        ''' Close workbook
        ''' </summary>
        Public Sub CloseAndDispose()
            Me.Close()
        End Sub

        ''' <summary>
        ''' Close workbook
        ''' </summary>
        Public Overrides Sub Close()
            If Me.ComObject IsNot Nothing Then Me.ComObjectStronglyTyped.Close(False)
            Me.Parent.RemoveWorkbookWrapper(Me)
            MyBase.Close()
        End Sub

        ''' <summary>
        ''' Workbook name
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Name As String = Me.ComObjectStronglyTyped.Name

    End Class

End Namespace