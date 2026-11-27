Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Spreadsheet
Imports OfficeExcel = DocumentFormat.OpenXml.Office.Excel
Imports Excel2010 = DocumentFormat.OpenXml.Office2010.Excel

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations
        Private Shared Function IsSupportedWorksheetConditionalFormatting(element As OpenXmlElement) As Boolean
            If TypeOf element Is ConditionalFormatting Then
                Return element.ChildElements.All(Function(rule) TypeOf rule Is ConditionalFormattingRule AndAlso
                    rule.ChildElements.All(Function(child) TypeOf child Is Formula))
            End If
            If Not TypeOf element Is WorksheetExtensionList Then Return False
            Return element.ChildElements.All(Function(extension) TypeOf extension Is WorksheetExtension AndAlso
                extension.ChildElements.Count = 1 AndAlso TypeOf extension.FirstChild Is Excel2010.ConditionalFormattings AndAlso
                extension.FirstChild.ChildElements.All(AddressOf IsSupportedExtendedConditionalFormatting))
        End Function

        Private Shared Function IsSupportedExtendedConditionalFormatting(element As OpenXmlElement) As Boolean
            If Not TypeOf element Is Excel2010.ConditionalFormatting Then Return False
            Return element.ChildElements.All(Function(child) TypeOf child Is OfficeExcel.ReferenceSequence OrElse
                (TypeOf child Is Excel2010.ConditionalFormattingRule AndAlso
                    child.ChildElements.All(Function(value) TypeOf value Is OfficeExcel.Formula OrElse TypeOf value Is Excel2010.DifferentialType)))
        End Function

        Private Shared Iterator Function WorksheetConditionalFormulas(worksheet As Worksheet) As IEnumerable(Of OpenXmlLeafTextElement)
            For Each format In worksheet.Elements(Of ConditionalFormatting)()
                For Each formula In format.Descendants(Of Formula)()
                    Yield formula
                Next
            Next
            For Each extension In worksheet.Elements(Of WorksheetExtensionList)()
                For Each format In extension.Descendants(Of Excel2010.ConditionalFormattings)()
                    For Each formula In format.Descendants(Of OfficeExcel.Formula)()
                        Yield formula
                    Next
                Next
            Next
        End Function
    End Class

End Namespace
