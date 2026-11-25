Namespace ExcelOps

    ''' <summary>Reports a known difference from Microsoft Excel during a structural reference update.</summary>
    ''' <remarks>
    ''' Set <see cref="ExcelDataOperationsOptions.AllowKnownStructuralReferenceDifferences"/> on the engine's
    ''' constructor options to allow the existing engine implementation for these edits.
    ''' </remarks>
    Public NotInheritable Class StructuralReferenceCompatibilityException
        Inherits NotSupportedException

        ''' <summary>Creates an exception for a detected compatibility limitation.</summary>
        ''' <param name="engineName">The name of the engine that rejected the edit.</param>
        ''' <param name="reason">The known limitation that applies to the edit.</param>
        Public Sub New(engineName As String, reason As String)
            MyBase.New(engineName & " detected a known Excel compatibility issue with this structural edit: " & reason &
                       ". Set ExcelDataOperationsOptions." & NameOf(ExcelDataOperationsOptions.AllowKnownStructuralReferenceDifferences) &
                       " to True to run the existing engine implementation.")
        End Sub
    End Class

End Namespace
