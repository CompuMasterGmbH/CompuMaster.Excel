Imports System.IO
Imports CompuMaster.Excel.ExcelOps
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports NUnit.Framework
Imports NativeExcel = Microsoft.Office.Interop.Excel

Namespace ExcelOpsTests.Engines

    Partial Public Class MsExcelOpsTest

        ' Relative coordinates in a defined name are offsets from each formula using the name.
        <TestCase("InsertRowBeforeFormulas")>
        <TestCase("InsertRowWithinRange")>
        <TestCase("RemoveFirstRangeRow")>
        <TestCase("InsertColumnBeforeFormulas")>
        <TestCase("InsertColumnWithinRange")>
        <TestCase("RemoveFirstRangeColumn")>
        <TestCase("InsertRowOnOtherSheet")>
        <TestCase("ShiftSingleCellDown")>
        <TestCase("ShiftRangeWidthDown")>
        <TestCase("ShiftRangeHeightRight")>
        <TestCase("DeleteSingleCellShiftUp")>
        Public Sub StructuralReference_OpenXmlMatchesExcelRelativeDefinedNames(operation As String)
            Dim directoryPath = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ExcelOps-NameReferences-" & Guid.NewGuid().ToString("N"))).FullName
            Dim input = Path.Combine(directoryPath, "before.xlsx")
            Dim nativeOutput = Path.Combine(directoryPath, "excel.xlsx")
            Dim openXmlOutput = Path.Combine(directoryPath, "openxml.xlsx")
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadOnly)
            Dim workbook As MsExcelDataOperations = Nothing
            Try
                workbook = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile,
                    New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile))
                workbook.RenameWorksheet("Sheet1", "Data", True)
                workbook.AddSheet("Other", Nothing)
                Dim data = DirectCast(workbook.Workbook.Worksheets.Item("Data"), NativeExcel.Worksheet)
                Dim other = DirectCast(workbook.Workbook.Worksheets.Item("Other"), NativeExcel.Worksheet)
                data.Activate()
                data.Range("A1").Select()
                For Each definition In {
                    (Name:="RelativeInput", Reference:="D4"),
                    (Name:="RelativeRange", Reference:="D4:F5"),
                    (Name:="AbsoluteRows", Reference:="D$4:F$5"),
                    (Name:="AbsoluteColumns", Reference:="$D4:$F5"),
                    (Name:="MixedRangeStart", Reference:="$D$4:F5"),
                    (Name:="MixedRangeEnd", Reference:="D4:$F$5"),
                    (Name:="MixedFirstRow", Reference:="D$4:F5"),
                    (Name:="MixedLastColumn", Reference:="D4:$F5"),
                    (Name:="FixedInput", Reference:="$D$4")}
                    workbook.Workbook.Names.Add(definition.Name, RefersTo:="=Data!" & definition.Reference)
                Next
                data.Names.Add("Data!LocalRelative", RefersTo:="=Data!D4")
                data.Names.Add("Data!LocalMixed", RefersTo:="=Data!D$4")
                other.Names.Add("Other!OtherLocalRelative", RefersTo:="=Data!D4")
                For row = 4 To 8
                    For column = 4 To 9
                        DirectCast(data.Cells(row, column), NativeExcel.Range).Value2 = row * 10 + column
                    Next
                Next
                data.Range("A1").Formula = "=RelativeInput"
                data.Range("A2").Formula = "=SUM(MixedRangeStart)"
                data.Range("A3").Formula = "=SUM(LocalMixed)"
                data.Range("A4").Formula = "=LocalRelative"
                other.Range("A1").Formula = "=OtherLocalRelative"
                other.Range("A2").Formula = "=SUM(MixedRangeEnd)"
                workbook.SaveAs(input, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                ApplyDefinedNameReferenceEdit(workbook, operation)
                workbook.SaveAs(nativeOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                workbook.Close()

                Using editor As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    ApplyDefinedNameReferenceEdit(editor, operation)
                    editor.SaveCopyAs(openXmlOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using
                Assert.That(ReadDefinedNameReferences(openXmlOutput), [Is].EqualTo(ReadDefinedNameReferences(nativeOutput)))

                workbook = Me.CreateInstance(nativeOutput, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                workbook.RecalculateAll()
                Dim expected = ReadDefinedNameFormulaResults(workbook, "Data", "Other")
                workbook.Close()
                workbook = Me.CreateInstance(openXmlOutput, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                workbook.RecalculateAll()
                Assert.That(ReadDefinedNameFormulaResults(workbook, "Data", "Other"), [Is].EqualTo(expected))
            Finally
                If workbook IsNot Nothing AndAlso Not workbook.IsClosed Then workbook.Close()
                Directory.Delete(directoryPath, True)
            End Try
        End Sub

        <TestCase("InsertRowWithinRange")>
        <TestCase("RemoveFirstRangeColumn")>
        <TestCase("ShiftSingleCellDown")>
        Public Sub StructuralReference_OpenXmlMatchesExcelRelativeNamesWithExternalSources(operation As String)
            Dim directoryPath = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "ExcelOps-LinkedNames-" & Guid.NewGuid().ToString("N"))).FullName
            Dim sourcePath = Path.Combine(directoryPath, "Linked source.xlsx")
            Dim input = Path.Combine(directoryPath, "before.xlsx")
            Dim nativeOutput = Path.Combine(directoryPath, "excel.xlsx")
            Dim openXmlOutput = Path.Combine(directoryPath, "openxml.xlsx")
            Dim options As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.ReadWrite)
            Dim createOptions As New ExcelDataOperationsOptions(ExcelDataOperationsOptions.WriteProtectionMode.DefaultBehaviourOnCreateFile)
            Dim source As MsExcelDataOperations = Nothing
            Dim target As MsExcelDataOperations = Nothing
            AlwaysCloseAllWorkbooksInNewEngineInstances = False
            Try
                source = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, createOptions)
                source.RenameWorksheet("Sheet1", "Data", True)
                For row = 3 To 6
                    For column = 3 To 5
                        source.WriteCellValue("Data", row, column, row * 10 + column)
                    Next
                Next
                source.SaveAs(sourcePath, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                target = Me.CreateInstance(Nothing, ExcelDataOperationsBase.OpenMode.CreateFile, createOptions)
                target.RenameWorksheet("Sheet1", "Check", True)
                Dim check = DirectCast(target.Workbook.Worksheets.Item("Check"), NativeExcel.Worksheet)
                target.Workbook.Activate()
                check.Activate()
                check.Range("A1").Select()
                Dim qualifier = "='[" & source.Workbook.Name & "]Data'!"
                target.Workbook.Names.Add("RemoteRelative", RefersTo:=qualifier & "D4")
                target.Workbook.Names.Add("RemoteMixedRow", RefersTo:=qualifier & "D$4")
                target.Workbook.Names.Add("RemoteMixedColumn", RefersTo:=qualifier & "$D4")
                target.Workbook.Names.Add("RemoteMixedRange", RefersTo:=qualifier & "$D$4:F5")
                target.Workbook.Names.Add("TargetLocal", RefersTo:="=Check!D4")
                check.Names.Add("Check!RemoteLocal", RefersTo:=qualifier & "D4")
                check.Range("A1").Formula = "=RemoteRelative"
                check.Range("A2").Formula = "=SUM(RemoteMixedRange)"
                check.Range("A3").Formula = "=RemoteLocal"
                target.SaveAs(input, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                ApplyDefinedNameReferenceEdit(source, operation)
                target.RecalculateAll()
                Dim expected = ReadDefinedNameFormulaResults(target, "Check")
                target.SaveAs(nativeOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                target.Close()
                source.Close()

                Using session As New OpenXmlWorkbookSession(),
                      linkedSource As New OpenXmlExcelDataOperations(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options),
                      linkedTarget As New OpenXmlExcelDataOperations(input, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                    session.Attach(linkedSource)
                    session.Attach(linkedTarget)
                    ApplyDefinedNameReferenceEdit(linkedSource, operation)
                    linkedSource.Save()
                    linkedTarget.SaveCopyAs(openXmlOutput, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                End Using
                Assert.That(ReadDefinedNameReferences(openXmlOutput), [Is].EqualTo(ReadDefinedNameReferences(nativeOutput)))
                source = Me.CreateInstance(sourcePath, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target = Me.CreateInstance(openXmlOutput, ExcelDataOperationsBase.OpenMode.OpenExistingFile, options)
                target.RecalculateAll()
                Assert.That(ReadDefinedNameFormulaResults(target, "Check"), [Is].EqualTo(expected))
            Finally
                If target IsNot Nothing AndAlso Not target.IsClosed Then target.Close()
                If source IsNot Nothing AndAlso Not source.IsClosed Then source.Close()
                Directory.Delete(directoryPath, True)
            End Try
        End Sub

        Private Shared Sub ApplyDefinedNameReferenceEdit(workbook As ExcelDataOperationsBase, operation As String)
            Select Case operation
                Case "InsertRowBeforeFormulas"
                    workbook.AddRow("Data", 0, 1, True)
                Case "InsertRowWithinRange"
                    workbook.AddRow("Data", 4, 2, True)
                Case "RemoveFirstRangeRow"
                    workbook.RemoveRows("Data", 3, 1, True)
                Case "InsertColumnBeforeFormulas"
                    workbook.AddColumn("Data", 0, 1, True)
                Case "InsertColumnWithinRange"
                    workbook.AddColumn("Data", 4, 2, True)
                Case "RemoveFirstRangeColumn"
                    workbook.RemoveColumns("Data", 3, 1, True)
                Case "InsertRowOnOtherSheet"
                    workbook.AddRow("Other", 0, 1, True)
                Case "ShiftSingleCellDown"
                    workbook.AddCells("Data", 3, 3, 1, 1, ExcelDataOperationsBase.CellInsertShiftDirection.ShiftCellsDown, True)
                Case "ShiftRangeWidthDown"
                    workbook.AddCells("Data", 3, 3, 1, 3, ExcelDataOperationsBase.CellInsertShiftDirection.ShiftCellsDown, True)
                Case "ShiftRangeHeightRight"
                    workbook.AddCells("Data", 3, 3, 2, 1, ExcelDataOperationsBase.CellInsertShiftDirection.ShiftCellsRight, True)
                Case "DeleteSingleCellShiftUp"
                    workbook.RemoveCells("Data", 3, 3, 1, 1, ExcelDataOperationsBase.CellRemoveShiftDirection.ShiftCellsUp, True)
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(operation))
            End Select
        End Sub

        Private Shared Function ReadDefinedNameReferences(path As String) As (Name As String, Scope As UInteger?, Text As String)()
            Using document = SpreadsheetDocument.Open(path, False)
                Return document.WorkbookPart.Workbook.DefinedNames.Elements(Of DefinedName)().
                    OrderBy(Function(definition) definition.Name.Value, StringComparer.Ordinal).
                    Select(Function(definition) (definition.Name.Value, definition.LocalSheetId?.Value, definition.Text)).ToArray()
            End Using
        End Function

        Private Shared Function ReadDefinedNameFormulaResults(workbook As MsExcelDataOperations, ParamArray sheetNames As String()) As Object()
            Dim formulas As New List(Of Object)
            For Each sheetName In sheetNames
                Dim sheet = DirectCast(workbook.Workbook.Worksheets.Item(sheetName), NativeExcel.Worksheet)
                For row = 1 To 5
                    For column = 1 To 2
                        formulas.Add(DirectCast(sheet.Cells(row, column), NativeExcel.Range).Value2)
                    Next
                Next
            Next
            Return formulas.ToArray()
        End Function

    End Class

End Namespace
