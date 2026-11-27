Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        ''' <inheritdoc/>
        ''' <remarks>Inserts an empty worksheet and selects it in the first workbook view.</remarks>
        Public Overrides Sub AddSheet(sheetName As String, beforeSheetName As String)
            ValidateNewWorksheetName(sheetName, NameOf(sheetName))
            SyncLock WorkbookSaveLock
                ValidateSheetLifecycle()
                Dim sheets = WorkbookSheets()
                If sheets.Any(Function(sheet) String.Equals(sheet.Name.Value, sheetName, StringComparison.OrdinalIgnoreCase)) Then
                    Throw New ArgumentException("A sheet with the requested name already exists.", NameOf(sheetName))
                End If
                Dim index = If(String.IsNullOrEmpty(beforeSheetName), sheets.Count, sheets.IndexOf(FindSheet(beforeSheetName)))
                Dim sheetId = sheets.Select(Function(sheet) sheet.SheetId.Value).DefaultIfEmpty(0UI).Max() + 1UI
                Using changes As New StagedWorkbookChanges()
                    changes.Add(Me,
                        Sub(book)
                            Dim workbook = book.Workbook
                            Dim part = book.AddNewPart(Of WorksheetPart)()
                            part.Worksheet = New Worksheet(New SheetData())
                            SaveWorksheet(part.Worksheet, part)
                            ShiftDestinationSheetIndexes(workbook, CUInt(index), 1UI)
                            workbook.Sheets.InsertAt(New Sheet With {.Name = sheetName, .SheetId = sheetId, .Id = book.GetIdOfPart(part)}, index)
                        End Sub)
                    changes.Apply()
                End Using
                SelectSheet(sheetName)
            End SyncLock
        End Sub

        ''' <inheritdoc/>
        ''' <remarks>Rewrites supported references to the removed sheet. The remaining workbook must satisfy the Open XML sheet transfer limitations.</remarks>
        Public Overrides Sub RemoveSheet(sheetName As String)
            SyncLock WorkbookSaveLock
                ValidateSheetLifecycle()
                Dim sheet = FindSheet(sheetName)
                ValidateMoveWorkbook()
                Dim removedName = sheet.Name.Value
                Dim removed As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {{removedName, removedName}}
                Dim wasActive = String.Equals(SelectedSheetName(), removedName, StringComparison.OrdinalIgnoreCase)
                Using changes As New StagedWorkbookChanges()
                    changes.Add(Me,
                        Sub(book)
                            RewriteMoveDependents(book, removed,
                                Function(referencedName, qualifier) If(removed.ContainsKey(referencedName), "#REF!", qualifier))
                            If book.Workbook.BookViews Is Nothing Then book.Workbook.AddChild(New BookViews(New WorkbookView()))
                            RemoveTransferredSheets(book, removed)
                        End Sub)
                    changes.Apply()
                End Using
                Dim selectedName = SelectedSheetName()
                If wasActive AndAlso selectedName IsNot Nothing Then
                    Dim selected = FindSheet(selectedName)
                    For Each remaining In WorkbookSheets()
                        SetTabSelected(remaining, 0UI, remaining Is selected)
                    Next
                End If
            End SyncLock
        End Sub

        Private Sub ValidateSheetLifecycle()
            If If(WorkbookPart.Workbook.WorkbookProtection?.LockStructure?.Value, False) Then
                Throw New InvalidOperationException("The workbook structure is protected.")
            End If
        End Sub

    End Class

End Namespace
