Imports MsExcel = Microsoft.Office.Interop.Excel

Namespace Global.CompuMaster.Excel.MsExcelCom

    ''' <summary>
    ''' COM Wrapper class for MS Excel workbooks collection
    ''' </summary>
    Public Class MsExcelWorkbooksWrapper
        Inherits CompuMaster.ComInterop.ComChildObject(Of MsExcelApplicationWrapper, MsExcel.Workbooks)

        Public Sub New(parent As MsExcelApplicationWrapper, obj As MsExcel.Workbooks)
            MyBase.New(parent, obj)
        End Sub

        ''' <summary>
        ''' Create a new workbook
        ''' </summary>
        ''' <returns></returns>
        Public Function Add() As MsExcelWorkbookWrapper
            Return Me.GetWorkbookWrapper(Me.ComObjectStronglyTyped.Add())
        End Function

        ''' <summary>
        ''' Open an existing workbook
        ''' </summary>
        ''' <param name="path"></param>
        ''' <param name="[readOnly]"></param>
        ''' <param name="passwordForOpening"></param>
        ''' <returns></returns>
        Public Function Open(path As String, [readOnly] As Boolean, passwordForOpening As String) As MsExcelWorkbookWrapper
            Return Me.GetWorkbookWrapper(Me.ComObjectStronglyTyped.Open(path, False, [readOnly], Nothing, If(passwordForOpening = Nothing, Nothing, passwordForOpening)))
        End Function

        Private _WorkbookWrappers As New Dictionary(Of MsExcel.Workbook, MsExcelWorkbookWrapper)
        Private Function GetWorkbookWrapper(comObject As MsExcel.Workbook) As MsExcelWorkbookWrapper
            If comObject Is Nothing Then Throw New ArgumentNullException(NameOf(comObject))

            Dim WorkbookForIsClosedCheck As MsExcelWorkbookWrapper = Nothing
            If _WorkbookWrappers.TryGetValue(comObject, WorkbookForIsClosedCheck) AndAlso WorkbookForIsClosedCheck.IsClosed Then
                Me._WorkbookWrappers.Remove(comObject)
            End If

            Dim Result As MsExcelWorkbookWrapper = Nothing
            If Not _WorkbookWrappers.TryGetValue(comObject, Result) Then
                Result = New MsExcelWorkbookWrapper(Me, comObject)
                Me._WorkbookWrappers.Add(comObject, Result)
            End If
            Return Result
        End Function
        Friend Sub RemoveWorkbookWrapper(item As MsExcelWorkbookWrapper)
            Me._WorkbookWrappers.Remove(item.ComObjectStronglyTyped)
        End Sub

        ''' <summary>
        ''' COM wrapper for Workbook
        ''' </summary>
        ''' <param name="index1Based">1-based index</param>
        ''' <returns></returns>
        Public Function Workbook(index1Based As Integer) As MsExcelWorkbookWrapper
            Return Me.GetWorkbookWrapper(Me.ComObjectStronglyTyped.Item(index1Based))
        End Function

        Public Function Workbook(name As String) As MsExcelWorkbookWrapper
            Return Me.GetWorkbookWrapper(Me.ComObjectStronglyTyped.Item(name))
        End Function

        ''' <summary>
        ''' Count of opened workbooks
        ''' </summary>
        ''' <returns></returns>
        Public ReadOnly Property Count() As Integer
            Get
                If Me.ComObject Is Nothing Then
                    Return 0
                Else
                    Return Me.ComObjectStronglyTyped.Count
                End If
            End Get
        End Property

        ''' <summary>
        ''' Close all opened workbooks
        ''' </summary>
        Public Sub CloseAllWorkbooks()
            For MyCounter As Integer = Me.Count - 1 To 0 Step -1
                Me.Workbook(MyCounter + 1).Close()
            Next
        End Sub

    End Class

End Namespace