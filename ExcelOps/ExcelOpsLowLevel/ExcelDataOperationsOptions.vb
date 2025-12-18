Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text

Namespace ExcelOps

    ''' <summary>
    ''' Base implementation for common API for the several Excel engines
    ''' </summary>
    Public Class ExcelDataOperationsOptions

        ''' <summary>
        ''' Create a new options instance
        ''' </summary>
        Public Sub New()
        End Sub

        Public Sub New(fileProtection As WriteProtectionMode)
            Me.FileWriteProtection = fileProtection
        End Sub

        Public Sub New(fileProtection As WriteProtectionMode, passwordForOpening As String)
            Me.FileWriteProtection = fileProtection
            Me.PasswordForOpening = passwordForOpening
        End Sub

        ''' <summary>
        ''' Create a new options instance
        ''' </summary>
        ''' <param name="passwordForOpening">Password for opening protected Excel files</param>
        Public Sub New(passwordForOpening As String)
            Me.PasswordForOpening = passwordForOpening
        End Sub

        ''' <summary>
        ''' Create a new options instance
        ''' </summary>
        ''' <param name="passwordForOpening">Password for opening protected Excel files</param>
        ''' <param name="disableInitialCalculation">If set to true, no initial calculation of formulas is performed when opening/loading an Excel file</param>
        ''' <param name="disableCalculationEngine">If set to true, the calculation engine is disabled and no formula calculations are performed</param>
        Public Sub New(passwordForOpening As String, disableInitialCalculation As Boolean?, disableAutoCalculation As Boolean?, disableCalculationEngine As Boolean?)
            Me.PasswordForOpening = passwordForOpening
            Me.DisableInitialCalculation = disableInitialCalculation
            Me.DisableAutoCalculationInWorkbook = disableAutoCalculation
            Me.DisableCalculationEngine = disableCalculationEngine
        End Sub

        ''' <summary>
        ''' Create a new options instance
        ''' </summary>
        ''' <param name="passwordForOpening">Password for opening protected Excel files</param>
        ''' <param name="disableInitialCalculation">If set to true, no initial calculation of formulas is performed when opening/loading an Excel file</param>
        ''' <param name="disableCalculationEngine">If set to true, the calculation engine is disabled and no formula calculations are performed</param>
        Public Sub New(fileProtection As WriteProtectionMode, passwordForOpening As String, disableInitialCalculation As Boolean?, disableAutoCalculation As Boolean?, disableCalculationEngine As Boolean?)
            Me.FileWriteProtection = fileProtection
            Me.PasswordForOpening = passwordForOpening
            Me.DisableInitialCalculation = disableInitialCalculation
            Me.DisableAutoCalculationInWorkbook = disableAutoCalculation
            Me.DisableCalculationEngine = disableCalculationEngine
        End Sub

        Public Enum WriteProtectionMode As Byte
            ''' <summary>
            ''' File can't be saved (saving with same file name is forbidden), but SaveAs with another file name is allowed
            ''' </summary>
            [ReadOnly] = 0
            ''' <summary>
            ''' No limitation
            ''' </summary>
            ReadWrite = 1
        End Enum

        ''' <summary>
        ''' Write protection for this filename prevents Save, but still allows SaveAs
        ''' </summary>
        ''' <returns></returns>
        Public Property FileWriteProtection As WriteProtectionMode = WriteProtectionMode.ReadOnly

        ''' <summary>
        ''' If set to true, the calculation engine is disabled and no formula calculations are performed, if set to false, the calculation engine is enabled, if null/not set, the engine default is used
        ''' </summary>
        ''' <remarks>Feature belongs to Excel engine</remarks>
        Public Property DisableCalculationEngine As Boolean?

        ''' <summary>
        ''' If set to true, no initial calculation of formulas is performed when opening/loading an Excel file, if set to false, the calculation engine is enabled, if null/not set, the engine default is used
        ''' </summary>
        ''' <remarks>Feature belongs to Excel engine</remarks>
        Public Property DisableInitialCalculation As Boolean?

        ''' <summary>
        ''' If set to true, automatic calculation mode is disabled in workbook (only manual calculation mode is used), if set to false, the calculation engine is enabled, if null/not set, the engine default is used
        ''' </summary>
        ''' <remarks>Feature belongs to workbook and changes permanently the workbook's behaviour when saved</remarks>
        Public Property DisableAutoCalculationInWorkbook As Boolean?

        ''' <summary>
        ''' Password for opening protected Excel files
        ''' </summary>
        ''' <returns></returns>
        Public Property PasswordForOpening As String

        ''' <summary>
        ''' Create a clone of this options instance
        ''' </summary>
        ''' <returns></returns>
        Private Function Clone() As ExcelDataOperationsOptions
            Return New ExcelDataOperationsOptions(PasswordForOpening, DisableInitialCalculation, DisableAutoCalculationInWorkbook, DisableCalculationEngine) With {
                .FileWriteProtection = Me.FileWriteProtection
            }
        End Function

        ''' <summary>
        ''' Apply default options from engine if not set and validate the resulting combination
        ''' </summary>
        ''' <param name="calculationDefaultOptions">Default calculation options from engine</param>
        ''' <returns>Validated options instance</returns>
        Public Function ApplyDefaultsFromEngineAndValidate(calculationDefaultOptions As ExcelEngineDefaultOptions) As ExcelDataOperationsOptions
            Dim Result = Me.Clone
            If Result.DisableCalculationEngine.HasValue = False Then
                Result.DisableCalculationEngine = calculationDefaultOptions.DisableCalculationEngine
            End If
            If Result.DisableInitialCalculation.HasValue = False Then
                Result.DisableInitialCalculation = calculationDefaultOptions.DisableInitialCalculation
            End If
            Return Result
        End Function

    End Class

End Namespace
