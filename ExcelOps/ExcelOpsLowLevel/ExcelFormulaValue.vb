Namespace ExcelOps

    ''' <summary>Identifies the type of a scalar formula value.</summary>
    Public Enum ExcelFormulaValueKind
        ''' <summary>The cell or argument is empty.</summary>
        Blank
        ''' <summary>The value is numeric.</summary>
        Number
        ''' <summary>The value is text.</summary>
        Text
        ''' <summary>The value is logical.</summary>
        Logical
        ''' <summary>The value is an Excel error.</summary>
        [Error]
    End Enum

    ''' <summary>Represents a typed scalar argument or result of a custom formula function.</summary>
    Public Structure ExcelFormulaValue
        Private ReadOnly _kind As ExcelFormulaValueKind
        Private ReadOnly _value As Object

        Private Sub New(kind As ExcelFormulaValueKind, value As Object)
            _kind = kind
            _value = value
        End Sub

        ''' <summary>Gets the type of the value.</summary>
        ''' <value>The scalar value kind. A default value represents a blank.</value>
        Public ReadOnly Property Kind As ExcelFormulaValueKind
            Get
                Return _kind
            End Get
        End Property

        ''' <summary>Gets the scalar value.</summary>
        ''' <value>A Double, String, Boolean, Excel error code, or Nothing for a blank.</value>
        Public ReadOnly Property Value As Object
            Get
                Return _value
            End Get
        End Property

        ''' <summary>Creates a numeric formula value.</summary>
        ''' <param name="value">The number.</param>
        ''' <returns>The numeric value, or the Excel error #NUM! for a non-finite number.</returns>
        Public Shared Function FromNumber(value As Double) As ExcelFormulaValue
            If Double.IsNaN(value) OrElse Double.IsInfinity(value) Then Return FromError("#NUM!")
            Return New ExcelFormulaValue(ExcelFormulaValueKind.Number, value)
        End Function

        ''' <summary>Creates a text formula value.</summary>
        ''' <param name="value">The text, including an empty string.</param>
        ''' <returns>The text value.</returns>
        Public Shared Function FromText(value As String) As ExcelFormulaValue
            If value Is Nothing Then Throw New ArgumentNullException(NameOf(value))
            Return New ExcelFormulaValue(ExcelFormulaValueKind.Text, value)
        End Function

        ''' <summary>Creates a logical formula value.</summary>
        ''' <param name="value">The logical value.</param>
        ''' <returns>The logical value.</returns>
        Public Shared Function FromBoolean(value As Boolean) As ExcelFormulaValue
            Return New ExcelFormulaValue(ExcelFormulaValueKind.Logical, value)
        End Function

        ''' <summary>Creates an Excel error value.</summary>
        ''' <param name="code">An Excel error code such as #VALUE! or #N/A.</param>
        ''' <returns>The error value.</returns>
        ''' <exception cref="ArgumentException">The code is outside the supported error set.</exception>
        Public Shared Function FromError(code As String) As ExcelFormulaValue
            Select Case code
                Case "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!"
                    Return New ExcelFormulaValue(ExcelFormulaValueKind.Error, code)
                Case Else
                    Throw New ArgumentException("Unsupported Excel error code.", NameOf(code))
            End Select
        End Function
    End Structure

End Namespace
