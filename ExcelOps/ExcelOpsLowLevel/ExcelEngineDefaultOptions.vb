Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text

Namespace ExcelOps

    ''' <summary>
    ''' Default options for an Excel engine
    ''' </summary>
    Public Class ExcelEngineDefaultOptions

        ''' <summary>
        ''' Create a new options instance
        ''' </summary>
        ''' <param name="disableInitialCalculation">If set to true, no initial calculation of formulas is performed when opening/loading an Excel file, if set to false, the engine runs a full recalculation after loading a workbook</param>
        ''' <param name="disableCalculationEngine">If set to true, the calculation engine is disabled and no formula calculations can be performed, if set to false the calculation engine is available</param>
        Public Sub New(disableInitialCalculation As Boolean, disableCalculationEngine As Boolean)
            Me.DisableInitialCalculation = disableInitialCalculation
            Me.DisableCalculationEngine = disableCalculationEngine
        End Sub

        ''' <summary>
        ''' If set to true, the calculation engine is disabled and no formula calculations are performed, if set to false, the calculation engine is enabled, if null/not set, the engine default is used
        ''' </summary>
        Public ReadOnly Property DisableCalculationEngine As Boolean

        ''' <summary>
        ''' If set to true, no initial calculation of formulas is performed when opening/loading an Excel file, if set to false, the calculation engine is enabled, if null/not set, the engine default is used
        ''' </summary>
        Public ReadOnly Property DisableInitialCalculation As Boolean

    End Class

End Namespace
