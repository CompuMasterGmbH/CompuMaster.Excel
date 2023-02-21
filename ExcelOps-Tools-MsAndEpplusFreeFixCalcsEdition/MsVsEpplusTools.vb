Option Explicit On
Option Strict On

Namespace ExcelOps

    ''' <summary>
    ''' Tools for fixing issues of workbooks between Epplus and Microsoft Excel
    ''' </summary>
    ''' <remarks>
    ''' PLEASE NOTE: Considerations for server-side Automation of Office https://support.microsoft.com/en-us/topic/considerations-for-server-side-automation-of-office-48bcfe93-8a89-47f1-0bce-017433ad79e2
    ''' </remarks>
    Public NotInheritable Class MsVsEpplusTools

        ''' <summary>
        ''' Due to a bug in EPPlus, the Excel workbook file contains calculated caches which are used by MS Excel but never reset by MS Excel, for this it might be required to reset all cached calucations
        ''' </summary>
        ''' <param name="path"></param>
        Public Shared Sub OpenAndClearCalculatedValuesToForceRecalculationOnNextOpeningWithMsExcelAndCloseExcelWorkbookWithEpplus(path As String)
            OpenAndClearCalculatedValuesToForceRecalculationOnNextOpeningWithMsExcelAndCloseExcelWorkbookWithEpplus(path, Nothing)
        End Sub

        ''' <summary>
        ''' Due to a bug in EPPlus, the Excel workbook file contains calculated caches which are used by MS Excel but never reset by MS Excel, for this it might be required to reset all cached calucations
        ''' </summary>
        ''' <param name="path"></param>
        Public Shared Sub OpenAndClearCalculatedValuesToForceRecalculationOnNextOpeningWithMsExcelAndCloseExcelWorkbookWithEpplus(path As String, passwordForOpening As String)
            Dim Wb As New ExcelOps.EpplusFreeExcelDataOperations(path, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, passwordForOpening) With {
                .RecalculationRequired = True
            }
            Wb.SaveAs(Wb.FilePath, ExcelOps.ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.AlwaysResetCalculatedValuesForForcedCellRecalculation)
            Wb.Close()
            Wb.CloseExcelAppInstance()
        End Sub

        ''' <summary>
        ''' Execute a full recalculation
        ''' </summary>
        ''' <param name="path"></param>
        Public Shared Sub OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(path As String)
            OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(path, Nothing)
        End Sub

        ''' <summary>
        ''' Execute a full recalculation
        ''' </summary>
        ''' <param name="path"></param>
        Public Shared Sub OpenAndClearCalculationCachesAndRecalculateAndCloseExcelWorkbookWithMsExcel(path As String, passwordForOpening As String)
            OpenAndClearCalculatedValuesToForceRecalculationOnNextOpeningWithMsExcelAndCloseExcelWorkbookWithEpplus(path)
            Dim MSExcel As MsExcelCom.MsExcelApplicationWrapper = Nothing
            Dim MsExcelWb As ExcelOps.MsExcelDataOperations = Nothing
            Try
                MSExcel = New MsExcelCom.MsExcelApplicationWrapper
                MsExcelWb = New ExcelOps.MsExcelDataOperations(path, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, False, passwordForOpening)
                MsExcelWb.RecalculateAll()
                MsExcelWb.Save()
                MsExcelWb.Close()
                MsExcelWb.CloseExcelAppInstance()
                MsExcelWb = Nothing
            Finally
                If MsExcelWb IsNot Nothing Then
                    Try
                        MsExcelWb.Close()
                        MsExcelWb.CloseExcelAppInstance()
                    Catch
                        'Ignore
                    End Try
                End If
                If MSExcel IsNot Nothing Then
                    MSExcel.Dispose()
                End If
            End Try
        End Sub

    End Class

End Namespace