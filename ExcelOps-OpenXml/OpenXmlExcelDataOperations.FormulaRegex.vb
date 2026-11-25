Imports PCRE

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private NotInheritable Partial Class FormulaCalculation

            Private Shared ReadOnly LookupMatchLimits As New PcreMatchSettings With {
                .MatchLimit = 100000UI,
                .DepthLimit = 1000UI,
                .HeapLimit = 8192UI
            }

            Private Function SelectRegexLookupValue(needle As Object, keys As Object(), reverse As Boolean,
                                                    selectValue As Func(Of Integer, Object)) As Object
                Dim index = -1
                Try
                    Dim regex As New PcreRegex(TextValue(needle), PcreOptions.Utf Or PcreOptions.Ucp)
                    For offset = 0 To keys.Length - 1
                        Dim candidate = If(reverse, keys.Length - 1 - offset, offset)
                        If TypeOf keys(candidate) Is FormulaError Then Continue For
                        Dim subject = TextValue(keys(candidate))
                        Dim match = regex.Match(subject, 0, PcreMatchOptions.None,
                            DirectCast(Nothing, Func(Of PcreCallout, PcreCalloutResult)), LookupMatchLimits)
                        If Not match.Success Then Continue For
                        index = candidate
                        Exit For
                    Next
                Catch ex As PcrePatternException
                    Return New FormulaError("#VALUE!")
                Catch ex As PcreException When ex.ErrorCode = PcreErrorCode.MatchLimit OrElse ex.ErrorCode = PcreErrorCode.DepthLimit OrElse
                                               ex.ErrorCode = PcreErrorCode.HeapLimit
                    Throw New NotSupportedException("The regular expression exceeds the supported calculation resource limits.", ex)
                Catch ex As PcreException
                    Return New FormulaError("#VALUE!")
                Catch ex As DllNotFoundException
                    Throw New NotSupportedException("PCRE2 is unavailable on this platform.", ex)
                Catch ex As BadImageFormatException
                    Throw New NotSupportedException("The installed PCRE2 library is incompatible with this process.", ex)
                End Try
                Return selectValue(index)
            End Function

        End Class

    End Class

End Namespace
