Imports System.Globalization
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData
Imports RichTypes = DocumentFormat.OpenXml.Office2019.Excel.RichData2
Imports Features = DocumentFormat.OpenXml.Office2019.Excel.CalcFeatures

Namespace ExcelOps

    Partial Public Class OpenXmlExcelDataOperations

        Private Const CalcErrorCode As String = "#CALC!"
        Private Const RichValueMetadataName As String = "XLRICHVALUE"
        Private Const DynamicArrayMetadataName As String = "XLDAPR"
        Private Const DynamicArrayNamespace As String = "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"
        Private Const RichDataNamespace As String = "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"

        Private Function IsSingleCellDynamicFormula(cell As Cell) As Boolean
            If cell?.CellFormula?.FormulaType Is Nothing OrElse cell.CellFormula.FormulaType.Value <> CellFormulaValues.Array OrElse
                cell.CellMetaIndex Is Nothing OrElse cell.CellReference Is Nothing Then Return False
            Dim reference = cell.CellFormula.Reference?.Value
            Dim address = cell.CellReference.Value
            If Not String.Equals(reference, address, StringComparison.OrdinalIgnoreCase) AndAlso
                Not String.Equals(reference, address & ":" & address, StringComparison.OrdinalIgnoreCase) Then Return False

            Dim metadata = WorkbookPart.CellMetadataPart?.Metadata
            Dim blocks = metadata?.GetFirstChild(Of CellMetadata)()
            Dim block = FindFutureMetadataBlock(metadata, blocks, cell.CellMetaIndex.Value, DynamicArrayMetadataName)
            Dim isDynamic = MetadataExtensionAttribute(block, DynamicArrayNamespace, "dynamicArrayProperties", "fDynamic")
            Return isDynamic = "1" OrElse isDynamic = "true"
        End Function

        Private Function ReadStoredFormulaError(cell As Cell) As String
            Dim fallback = If(cell.CellValue?.Text, "#VALUE!")
            If cell.ValueMetaIndex Is Nothing OrElse fallback <> "#VALUE!" Then Return fallback
            Dim metadata = WorkbookPart.CellMetadataPart?.Metadata
            Dim block = FindFutureMetadataBlock(metadata, metadata?.GetFirstChild(Of ValueMetadata)(),
                cell.ValueMetaIndex.Value, RichValueMetadataName)
            Dim richIndex As Integer
            If Not Integer.TryParse(MetadataExtensionAttribute(block, RichDataNamespace, "rvb", "i"),
                NumberStyles.None, CultureInfo.InvariantCulture, richIndex) Then Return fallback
            Dim value = WorkbookPart.GetPartsOfType(Of RdRichValuePart)().FirstOrDefault()?.RichValueData?.
                Elements(Of Rich.RichValue)().ElementAtOrDefault(richIndex)
            If value?.S Is Nothing OrElse value.S.Value > Integer.MaxValue Then Return fallback
            Dim errorStructure = WorkbookPart.GetPartsOfType(Of RdRichValueStructurePart)().FirstOrDefault()?.RichValueStructures?.
                Elements(Of Rich.RichValueStructure)().ElementAtOrDefault(CInt(value.S.Value))
            If errorStructure?.T Is Nothing OrElse errorStructure.T.Value <> "_error" Then Return fallback

            Dim keys = errorStructure.Elements(Of Rich.Key)().ToList()
            Dim index = keys.FindIndex(Function(key) key.N?.Value = "errorType")
            If index >= 0 AndAlso value.Elements(Of Rich.Value)().ElementAtOrDefault(index)?.Text = "13" Then Return CalcErrorCode
            Return fallback
        End Function

        Private Shared Function MetadataExtensionAttribute(block As FutureMetadataBlock, namespaceUri As String, elementName As String,
                                                           attributeName As String) As String
            ' The SDK loads the contents of a future-metadata extension as unknown elements.
            Dim extensionElement = block?.Descendants().FirstOrDefault(
                Function(item) item.LocalName = elementName AndAlso item.NamespaceUri = namespaceUri)
            If extensionElement Is Nothing Then Return Nothing
            Return extensionElement.GetAttributes().FirstOrDefault(
                Function(attribute) attribute.LocalName = attributeName AndAlso String.IsNullOrEmpty(attribute.NamespaceUri)).Value
        End Function

        Private Shared Function FindFutureMetadataBlock(metadata As Metadata, blocks As MetadataBlocksType, index As UInteger,
                                                        typeName As String) As FutureMetadataBlock
            If blocks Is Nothing OrElse index = 0UI OrElse index > Integer.MaxValue Then Return Nothing
            Dim block = blocks.Elements(Of MetadataBlock)().ElementAtOrDefault(CInt(index - 1UI))
            If block Is Nothing Then Return Nothing
            Dim types = metadata.MetadataTypes?.Elements(Of MetadataType)().ToList()
            If types Is Nothing Then Return Nothing
            Dim typeIndex = types.FindIndex(Function(item) item.Name?.Value = typeName) + 1
            If typeIndex = 0 Then Return Nothing
            Dim record = block.Elements(Of MetadataRecord)().FirstOrDefault(Function(item) If(item.TypeIndex?.Value, 0UI) = CUInt(typeIndex))
            If record?.Val Is Nothing OrElse record.Val.Value > Integer.MaxValue Then Return Nothing
            Return metadata.Elements(Of FutureMetadata)().FirstOrDefault(Function(item) item.Name?.Value = typeName)?.
                Elements(Of FutureMetadataBlock)().ElementAtOrDefault(CInt(record.Val.Value))
        End Function

        Private Shared Function AddCalcErrorValueMetadata(workbook As WorkbookPart) As UInteger
            Dim metadataPart = If(workbook.CellMetadataPart, workbook.AddNewPart(Of CellMetadataPart)())
            Dim metadata = metadataPart.Metadata
            If metadata Is Nothing Then
                metadata = New Metadata()
                metadataPart.Metadata = metadata
            End If

            Dim typeList = metadata.MetadataTypes
            If typeList Is Nothing Then
                typeList = New MetadataTypes()
                metadata.MetadataTypes = typeList
            End If
            Dim types = typeList.Elements(Of MetadataType)().ToList()
            Dim typeIndex = types.FindIndex(Function(item) item.Name?.Value = RichValueMetadataName)
            If typeIndex < 0 Then
                typeIndex = types.Count
                typeList.AppendChild(New MetadataType With {
                    .Name = RichValueMetadataName, .MinSupportedVersion = 120000UI, .Copy = True, .PasteAll = True,
                    .PasteValues = True, .Merge = True, .SplitFirst = True, .RowColumnShift = True,
                    .ClearFormats = True, .ClearComments = True, .Assign = True, .Coerce = True
                })
                typeList.Count = CUInt(types.Count + 1)
            End If

            If Not workbook.RdRichValueTypesParts.Any() Then
                Dim typesPart = workbook.AddNewPart(Of RdRichValueTypesPart)()
                typesPart.RichValueTypesInfo = New RichTypes.RichValueTypesInfo()
                typesPart.RichValueTypesInfo.Save()
            End If
            Dim valuesPart = If(workbook.GetPartsOfType(Of RdRichValuePart)().FirstOrDefault(), workbook.AddNewPart(Of RdRichValuePart)())
            Dim structuresPart = If(workbook.GetPartsOfType(Of RdRichValueStructurePart)().FirstOrDefault(), workbook.AddNewPart(Of RdRichValueStructurePart)())
            Dim structures = structuresPart.RichValueStructures
            If structures Is Nothing Then
                structures = New Rich.RichValueStructures()
                structuresPart.RichValueStructures = structures
            End If
            Dim errorKeys = {"errorType", "subType"}
            Dim existingStructures = structures.Elements(Of Rich.RichValueStructure)().ToList()
            Dim structureIndex = existingStructures.FindIndex(
                Function(item) item.T?.Value = "_error" AndAlso item.Elements(Of Rich.Key)().
                    Select(Function(key) key.N?.Value).SequenceEqual(errorKeys))
            If structureIndex < 0 Then
                structureIndex = existingStructures.Count
                Dim errorStructure As New Rich.RichValueStructure With {.T = "_error"}
                For Each name In errorKeys
                    Dim key As New Rich.Key With {.N = name}
                    key.SetAttribute(New OpenXmlAttribute("t", String.Empty, "i"))
                    errorStructure.AppendChild(key)
                Next
                structures.AppendChild(errorStructure)
                structures.Count = CUInt(structureIndex + 1)
                structures.Save()
            End If

            Dim values = valuesPart.RichValueData
            If values Is Nothing Then
                values = New Rich.RichValueData()
                valuesPart.RichValueData = values
            End If
            Dim valueIndex = CUInt(values.Elements(Of Rich.RichValue)().Count())
            ' Excel uses error type 13 and subtype 3 for an empty-array calculation error.
            values.AppendChild(New Rich.RichValue(New Rich.Value("13"), New Rich.Value("3")) With {.S = CUInt(structureIndex)})
            values.Count = valueIndex + 1UI
            values.Save()

            Dim future = metadata.Elements(Of FutureMetadata)().FirstOrDefault(Function(item) item.Name?.Value = RichValueMetadataName)
            If future Is Nothing Then
                future = New FutureMetadata With {.Name = RichValueMetadataName}
                metadata.AddChild(future)
            End If
            Dim futureIndex = CUInt(future.Elements(Of FutureMetadataBlock)().Count())
            Dim extension As New Extension With {.Uri = "{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"}
            extension.AppendChild(New Rich.RichValueBlock With {.I = valueIndex})
            future.AppendChild(New FutureMetadataBlock(New ExtensionList(extension)))
            future.Count = futureIndex + 1UI

            Dim blocks = metadata.GetFirstChild(Of ValueMetadata)()
            If blocks Is Nothing Then
                blocks = New ValueMetadata()
                metadata.AddChild(blocks)
            End If
            blocks.AppendChild(New MetadataBlock(New MetadataRecord With {.TypeIndex = CUInt(typeIndex + 1), .Val = futureIndex}))
            blocks.Count = CUInt(blocks.Elements(Of MetadataBlock)().Count())
            metadata.Save()
            EnsureRichDataCalculationFeature(workbook.Workbook)
            Return blocks.Count.Value
        End Function

        Private Shared Sub EnsureRichDataCalculationFeature(workbook As Workbook)
            Const extensionUri As String = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}"
            Const featureName As String = "microsoft.com:RD"
            Const featuresNamespace As String = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures"
            Const richErrorCalculationVersion As UInteger = 181029UI

            ' Excel loads saved rich errors only with compatible workbook and calculation version metadata.
            If workbook.FileVersion Is Nothing Then
                workbook.FileVersion = New FileVersion With {.ApplicationName = "xl", .LastEdited = "7", .LowestEdited = "7"}
            End If
            Dim calculation = workbook.CalculationProperties
            If calculation Is Nothing Then
                calculation = New CalculationProperties()
                workbook.CalculationProperties = calculation
            End If
            If calculation.CalculationId Is Nothing Then
                calculation.CalculationId = richErrorCalculationVersion
            ElseIf calculation.CalculationId.Value < richErrorCalculationVersion Then
                calculation.CalculationId = richErrorCalculationVersion
                calculation.FullCalculationOnLoad = True
            End If

            Dim extensions = workbook.GetFirstChild(Of WorkbookExtensionList)()
            If extensions Is Nothing Then
                extensions = New WorkbookExtensionList()
                workbook.AddChild(extensions)
            End If
            Dim extension = extensions.Elements(Of WorkbookExtension)().FirstOrDefault(
                Function(item) String.Equals(item.Uri?.Value, extensionUri, StringComparison.OrdinalIgnoreCase))
            If extension Is Nothing Then
                extension = New WorkbookExtension With {.Uri = extensionUri}
                extensions.AppendChild(extension)
            End If
            Dim featureList = extension.GetFirstChild(Of Features.CalcFeatures)()
            If featureList Is Nothing Then
                ' Some SDK compatibility modes expose existing extension children as unknown elements.
                Dim existing = extension.ChildElements.FirstOrDefault(
                    Function(item) item.LocalName = "calcFeatures" AndAlso item.NamespaceUri = featuresNamespace)
                featureList = If(existing Is Nothing, New Features.CalcFeatures(), New Features.CalcFeatures(existing.OuterXml))
                If existing IsNot Nothing Then existing.Remove()
                extension.AppendChild(featureList)
            End If
            If Not featureList.Elements(Of Features.CalcFeature)().Any(Function(item) item.Name?.Value = featureName) Then
                featureList.AppendChild(New Features.CalcFeature With {.Name = featureName})
            End If
            workbook.Save()
        End Sub

        Private Shared Function IsUnencodedCalcError(cell As Cell) As Boolean
            Return cell.CellFormula IsNot Nothing AndAlso If(cell.DataType?.Value, CellValues.Number) = CellValues.Error AndAlso
                cell.ValueMetaIndex Is Nothing AndAlso cell.CellValue?.Text = CalcErrorCode
        End Function

        Private Shared Sub SaveRichFormulaErrors(workbook As WorkbookPart, worksheet As WorksheetPart, ByRef calcErrorMetadataIndex As UInteger?)
            Dim calcErrors = GetSheetData(worksheet).Descendants(Of Cell)().Where(AddressOf IsUnencodedCalcError).ToList()
            If calcErrors.Count = 0 Then Return
            If Not calcErrorMetadataIndex.HasValue Then calcErrorMetadataIndex = AddCalcErrorValueMetadata(workbook)
            For Each cell In calcErrors
                cell.CellValue.Text = "#VALUE!"
                cell.ValueMetaIndex = calcErrorMetadataIndex.Value
            Next
            worksheet.Worksheet.Save()
        End Sub

    End Class

End Namespace
