Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports MergeCollection = DocumentFormat.OpenXml.Spreadsheet.MergeCells

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        ''' <remarks>Copies formatted cell content and preserves formula text. Use whole-sheet transfer to adapt workbook references.</remarks>
        Public Overrides Sub CopySheetContentInternal(sheetName As String, targetWorkbook As ExcelDataOperationsBase, targetSheetName As String)
            TryCopySheetContentCore(sheetName, targetWorkbook, targetSheetName, CopySheetOption.TargetSheetMustExist)
        End Sub

        ''' <inheritdoc/>
        Protected Overrides Function TryCopySheetContentCore(sheetName As String, targetWorkbook As ExcelDataOperationsBase,
                                                            targetSheetName As String, copyOption As CopySheetOption) As Boolean
            Dim destination = TryCast(targetWorkbook, OpenXmlExcelDataOperations)
            If destination Is Nothing Then Throw UnsupportedOperation("Copying worksheet content between different engines")
            If destination Is Me Then Throw New ArgumentException("The destination must be a different workbook.", NameOf(targetWorkbook))
            ValidateNewWorksheetName(targetSheetName, NameOf(targetSheetName))
            SyncLock WorkbookSessionLock
                Return WithLockedWorkbooks({Me, destination},
                    Function()
                        CopyOpenXmlSheetContent(sheetName, destination, targetSheetName, copyOption)
                        Return True
                    End Function)
            End SyncLock
        End Function

        Private Sub CopyOpenXmlSheetContent(sheetName As String, destination As OpenXmlExcelDataOperations,
                                            targetSheetName As String, copyOption As CopySheetOption)
            Dim source = GetWorksheetPart(sheetName)
            Dim existing = destination.WorkbookSheets().Find(Function(sheet) String.Equals(sheet.Name.Value, targetSheetName, StringComparison.OrdinalIgnoreCase))
            Select Case copyOption
                Case CopySheetOption.TargetSheetMustNotExist
                    If existing IsNot Nothing Then Throw New InvalidOperationException("The destination worksheet already exists.")
                Case CopySheetOption.TargetSheetMustExist
                    If existing Is Nothing Then Throw New InvalidOperationException("The destination worksheet is missing.")
                Case CopySheetOption.TargetSheetMightExist
                Case Else
                    Throw New ArgumentOutOfRangeException(NameOf(copyOption))
            End Select

            If existing Is Nothing Then
                destination.ValidateSheetLifecycle()
            Else
                destination.ValidateCellLayoutEdit(destination.GetWorksheetPart(existing.Name.Value), Nothing, "Replacing worksheet content")
            End If
            ValidateCopyWorkbookStyles(destination)
            Dim worksheet = PrepareWorksheetCopy(source, sheetName, CopyWorksheetFormulaMode.PreserveFormulas, Nothing, move:=False, observeName:=Nothing)

            Using changes As New StagedWorkbookChanges()
                changes.Add(destination,
                    Sub(book)
                        TransferWorksheetStyles(book, {worksheet})
                        If WorkbookPart.ThemePart IsNot Nothing AndAlso book.ThemePart Is Nothing Then book.AddPart(WorkbookPart.ThemePart)
                        Dim part As WorksheetPart
                        If existing Is Nothing Then
                            part = book.AddNewPart(Of WorksheetPart)()
                            part.Worksheet = New Worksheet(New SheetData())
                            Dim sheets = book.Workbook.Sheets
                            Dim id = sheets.Elements(Of Sheet)().Select(Function(item) item.SheetId.Value).DefaultIfEmpty(0UI).Max() + 1UI
                            sheets.AppendChild(New Sheet With {.Id = book.GetIdOfPart(part), .Name = targetSheetName, .SheetId = id})
                        Else
                            part = DirectCast(book.GetPartById(existing.Id.Value), WorksheetPart)
                        End If
                        Dim target = part.Worksheet
                        target.ReplaceChild(worksheet.GetFirstChild(Of SheetData)().CloneNode(True), GetSheetData(part))
                        target.SheetDimension = If(worksheet.SheetDimension Is Nothing, Nothing, DirectCast(worksheet.SheetDimension.CloneNode(True), SheetDimension))
                        target.RemoveAllChildren(Of Columns)()
                        For Each columns In worksheet.Elements(Of Columns)()
                            target.InsertBefore(columns.CloneNode(True), GetSheetData(part))
                        Next
                        target.RemoveAllChildren(Of MergeCollection)()
                        Dim merges = worksheet.GetFirstChild(Of MergeCollection)()
                        If merges IsNot Nothing Then target.AddChild(merges.CloneNode(True))
                        SaveWorksheet(target, part)
                    End Sub)
                changes.Apply()
            End Using
        End Sub

    End Class

End Namespace
