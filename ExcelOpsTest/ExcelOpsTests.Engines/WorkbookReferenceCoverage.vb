Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOpsTests.Engines

    Friend NotInheritable Class WorkbookReferenceCoverage
        Friend Const SourceName As String = "WorkbookReferenceSource.xlsx"
        Friend Const TargetName As String = "WorkbookReferenceTarget.xlsx"
        Friend Const LocalSheet As String = "Local Data"
        Friend Const RemoteSheet As String = "Remote Data"
        Friend Const ChecksSheet As String = "Checks"

        Friend Shared Sub CopyPair(directory As String)
            For Each name In {SourceName, TargetName}
                File.Copy(TestEnvironment.FullPathOfExistingTestFile("test_data", name), Path.Combine(directory, name))
            Next
        End Sub

        Friend Shared Sub EditTarget(workbook As ExcelDataOperationsBase)
            workbook.AddRow(LocalSheet, 1, 1, True)
            workbook.AddColumn(LocalSheet, 1, 1, True)
            workbook.AddRow(ChecksSheet, 0, 1, True)
        End Sub

        Friend Shared Function CompareEditedTargetWithExcel(output As String) As List(Of String)
            Dim differences As New List(Of String)
            Using document = SpreadsheetDocument.Open(output, False)
                Dim part = document.WorkbookPart
                Dim checks = StructuralReferenceCoverage.WorksheetFor(part, ChecksSheet)
                StructuralReferenceCoverage.Compare(differences, "only the formula cell moves",
                    StructuralReferenceCoverage.FormulaAt(checks, "A2"), "'[1]Remote Data'!$B$2")
                StructuralReferenceCoverage.Compare(differences, "only the local range follows the edit",
                    StructuralReferenceCoverage.FormulaAt(checks, "A3"), "SUM('Local Data'!$C3:C$4,'[1]Remote Data'!$B$2:$B$3)")
                StructuralReferenceCoverage.Compare(differences, "local absolute and external relative references",
                    StructuralReferenceCoverage.FormulaAt(checks, "A4"), "'Local Data'!$C$3+'[1]Remote Data'!B3")
                Dim link = part.ExternalWorkbookParts.SingleOrDefault()
                Dim relationshipId = link?.ExternalLink?.GetFirstChild(Of ExternalBook)()?.Id?.Value
                Dim relationship = link?.ExternalRelationships.SingleOrDefault(Function(item) item.Id = relationshipId)
                StructuralReferenceCoverage.Compare(differences, "external link target", If(relationship?.Uri.OriginalString, ""),
                                                    SourceName)
            End Using
            Return differences
        End Function
    End Class

End Namespace
