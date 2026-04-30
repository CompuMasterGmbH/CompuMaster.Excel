''' <summary>
''' Contains Excel COM enumeration values used by the lightweight interop wrapper.
''' </summary>
Public Class Enumerations

    ''' <summary>
    ''' Defines worksheet visibility states used by Excel COM automation.
    ''' </summary>
    Public Enum XlSheetVisibility As Integer
        ''' <summary>
        ''' Represents a worksheet that is hidden but can be shown by the user.
        ''' </summary>
        xlSheetHidden = 0
        ''' <summary>
        ''' Represents a worksheet that is hidden and cannot be shown through the Excel user interface.
        ''' </summary>
        xlSheetVeryHidden = 2
        ''' <summary>
        ''' Represents a visible worksheet.
        ''' </summary>
        xlSheetVisible = -1
    End Enum

    ''' <summary>
    ''' Defines fixed output formats supported by Excel export.
    ''' </summary>
    Public Enum XlFixedFormatType As Integer
        ''' <summary>
        ''' Represents PDF fixed-format output.
        ''' </summary>
        xlTypePDF = 0
        ''' <summary>
        ''' Represents XPS fixed-format output.
        ''' </summary>
        xlTypeXPS = 1
    End Enum

    ''' <summary>
    ''' Defines quality levels for Excel fixed-format export.
    ''' </summary>
    Public Enum XlFixedFormatQuality As Integer
        ''' <summary>
        ''' Represents standard fixed-format output quality.
        ''' </summary>
        xlQualityStandard = 0
        ''' <summary>
        ''' Represents minimum fixed-format output quality.
        ''' </summary>
        xlQualityMinimum = 1
    End Enum

    ''' <summary>
    ''' Defines Excel built-in dialogs that can be shown through COM automation.
    ''' </summary>
    Public Enum XlBuiltInDialog As Integer
        ''' <summary>
        ''' Represents the Excel built-in Activate dialog.
        ''' </summary>
        xlDialogActivate = 103
        ''' <summary>
        ''' Represents the Excel built-in Active Cell Font dialog.
        ''' </summary>
        xlDialogActiveCellFont = 476
        ''' <summary>
        ''' Represents the Excel built-in Add Chart Autoformat dialog.
        ''' </summary>
        xlDialogAddChartAutoformat = 390
        ''' <summary>
        ''' Represents the Excel built-in Addin Manager dialog.
        ''' </summary>
        xlDialogAddinManager = 321
        ''' <summary>
        ''' Represents the Excel built-in Alignment dialog.
        ''' </summary>
        xlDialogAlignment = 43
        ''' <summary>
        ''' Represents the Excel built-in Apply Names dialog.
        ''' </summary>
        xlDialogApplyNames = 133
        ''' <summary>
        ''' Represents the Excel built-in Apply Style dialog.
        ''' </summary>
        xlDialogApplyStyle = 212
        ''' <summary>
        ''' Represents the Excel built-in App Move dialog.
        ''' </summary>
        xlDialogAppMove = 170
        ''' <summary>
        ''' Represents the Excel built-in App Size dialog.
        ''' </summary>
        xlDialogAppSize = 171
        ''' <summary>
        ''' Represents the Excel built-in Arrange All dialog.
        ''' </summary>
        xlDialogArrangeAll = 12
        ''' <summary>
        ''' Represents the Excel built-in Assign To Object dialog.
        ''' </summary>
        xlDialogAssignToObject = 213
        ''' <summary>
        ''' Represents the Excel built-in Assign To Tool dialog.
        ''' </summary>
        xlDialogAssignToTool = 293
        ''' <summary>
        ''' Represents the Excel built-in Attach Text dialog.
        ''' </summary>
        xlDialogAttachText = 80
        ''' <summary>
        ''' Represents the Excel built-in Attach Toolbars dialog.
        ''' </summary>
        xlDialogAttachToolbars = 323
        ''' <summary>
        ''' Represents the Excel built-in Auto Correct dialog.
        ''' </summary>
        xlDialogAutoCorrect = 485
        ''' <summary>
        ''' Represents the Excel built-in Axes dialog.
        ''' </summary>
        xlDialogAxes = 78
        ''' <summary>
        ''' Represents the Excel built-in Border dialog.
        ''' </summary>
        xlDialogBorder = 45
        ''' <summary>
        ''' Represents the Excel built-in Calculation dialog.
        ''' </summary>
        xlDialogCalculation = 32
        ''' <summary>
        ''' Represents the Excel built-in Cell Protection dialog.
        ''' </summary>
        xlDialogCellProtection = 46
        ''' <summary>
        ''' Represents the Excel built-in Change Link dialog.
        ''' </summary>
        xlDialogChangeLink = 166
        ''' <summary>
        ''' Represents the Excel built-in Chart Add Data dialog.
        ''' </summary>
        xlDialogChartAddData = 392
        ''' <summary>
        ''' Represents the Excel built-in Chart Location dialog.
        ''' </summary>
        xlDialogChartLocation = 527
        ''' <summary>
        ''' Represents the Excel built-in Chart Options Data Label Multiple dialog.
        ''' </summary>
        xlDialogChartOptionsDataLabelMultiple = 724
        ''' <summary>
        ''' Represents the Excel built-in Chart Options Data Labels dialog.
        ''' </summary>
        xlDialogChartOptionsDataLabels = 505
        ''' <summary>
        ''' Represents the Excel built-in Chart Options Data Table dialog.
        ''' </summary>
        xlDialogChartOptionsDataTable = 506
        ''' <summary>
        ''' Represents the Excel built-in Chart Source Data dialog.
        ''' </summary>
        xlDialogChartSourceData = 540
        ''' <summary>
        ''' Represents the Excel built-in Chart Trend dialog.
        ''' </summary>
        xlDialogChartTrend = 350
        ''' <summary>
        ''' Represents the Excel built-in Chart Type dialog.
        ''' </summary>
        xlDialogChartType = 526
        ''' <summary>
        ''' Represents the Excel built-in Chart Wizard dialog.
        ''' </summary>
        xlDialogChartWizard = 288
        ''' <summary>
        ''' Represents the Excel built-in Checkbox Properties dialog.
        ''' </summary>
        xlDialogCheckboxProperties = 435
        ''' <summary>
        ''' Represents the Excel built-in Clear dialog.
        ''' </summary>
        xlDialogClear = 52
        ''' <summary>
        ''' Represents the Excel built-in Color Palette dialog.
        ''' </summary>
        xlDialogColorPalette = 161
        ''' <summary>
        ''' Represents the Excel built-in Column Width dialog.
        ''' </summary>
        xlDialogColumnWidth = 47
        ''' <summary>
        ''' Represents the Excel built-in Combination dialog.
        ''' </summary>
        xlDialogCombination = 73
        ''' <summary>
        ''' Represents the Excel built-in Conditional Formatting dialog.
        ''' </summary>
        xlDialogConditionalFormatting = 583
        ''' <summary>
        ''' Represents the Excel built-in Consolidate dialog.
        ''' </summary>
        xlDialogConsolidate = 191
        ''' <summary>
        ''' Represents the Excel built-in Copy Chart dialog.
        ''' </summary>
        xlDialogCopyChart = 147
        ''' <summary>
        ''' Represents the Excel built-in Copy Picture dialog.
        ''' </summary>
        xlDialogCopyPicture = 108
        ''' <summary>
        ''' Represents the Excel built-in Create List dialog.
        ''' </summary>
        xlDialogCreateList = 796
        ''' <summary>
        ''' Represents the Excel built-in Create Names dialog.
        ''' </summary>
        xlDialogCreateNames = 62
        ''' <summary>
        ''' Represents the Excel built-in Create Publisher dialog.
        ''' </summary>
        xlDialogCreatePublisher = 217
        ''' <summary>
        ''' Represents the Excel built-in Create Relationship dialog.
        ''' </summary>
        xlDialogCreateRelationship = 1272
        ''' <summary>
        ''' Represents the Excel built-in Customize Toolbar dialog.
        ''' </summary>
        xlDialogCustomizeToolbar = 276
        ''' <summary>
        ''' Represents the Excel built-in Custom Views dialog.
        ''' </summary>
        xlDialogCustomViews = 493
        ''' <summary>
        ''' Represents the Excel built-in Data Delete dialog.
        ''' </summary>
        xlDialogDataDelete = 36
        ''' <summary>
        ''' Represents the Excel built-in Data Label dialog.
        ''' </summary>
        xlDialogDataLabel = 379
        ''' <summary>
        ''' Represents the Excel built-in Data Label Multiple dialog.
        ''' </summary>
        xlDialogDataLabelMultiple = 723
        ''' <summary>
        ''' Represents the Excel built-in Data Series dialog.
        ''' </summary>
        xlDialogDataSeries = 40
        ''' <summary>
        ''' Represents the Excel built-in Data Validation dialog.
        ''' </summary>
        xlDialogDataValidation = 525
        ''' <summary>
        ''' Represents the Excel built-in Define Name dialog.
        ''' </summary>
        xlDialogDefineName = 61
        ''' <summary>
        ''' Represents the Excel built-in Define Style dialog.
        ''' </summary>
        xlDialogDefineStyle = 229
        ''' <summary>
        ''' Represents the Excel built-in Delete Format dialog.
        ''' </summary>
        xlDialogDeleteFormat = 111
        ''' <summary>
        ''' Represents the Excel built-in Delete Name dialog.
        ''' </summary>
        xlDialogDeleteName = 110
        ''' <summary>
        ''' Represents the Excel built-in Demote dialog.
        ''' </summary>
        xlDialogDemote = 203
        ''' <summary>
        ''' Represents the Excel built-in Display dialog.
        ''' </summary>
        xlDialogDisplay = 27
        ''' <summary>
        ''' Represents the Excel built-in Document Inspector dialog.
        ''' </summary>
        xlDialogDocumentInspector = 862
        ''' <summary>
        ''' Represents the Excel built-in Editbox Properties dialog.
        ''' </summary>
        xlDialogEditboxProperties = 438
        ''' <summary>
        ''' Represents the Excel built-in Edit Color dialog.
        ''' </summary>
        xlDialogEditColor = 223
        ''' <summary>
        ''' Represents the Excel built-in Edit Delete dialog.
        ''' </summary>
        xlDialogEditDelete = 54
        ''' <summary>
        ''' Represents the Excel built-in Edition Options dialog.
        ''' </summary>
        xlDialogEditionOptions = 251
        ''' <summary>
        ''' Represents the Excel built-in Edit Series dialog.
        ''' </summary>
        xlDialogEditSeries = 228
        ''' <summary>
        ''' Represents the Excel built-in Errorbar X dialog.
        ''' </summary>
        xlDialogErrorbarX = 463
        ''' <summary>
        ''' Represents the Excel built-in Errorbar Y dialog.
        ''' </summary>
        xlDialogErrorbarY = 464
        ''' <summary>
        ''' Represents the Excel built-in Error Checking dialog.
        ''' </summary>
        xlDialogErrorChecking = 732
        ''' <summary>
        ''' Represents the Excel built-in Evaluate Formula dialog.
        ''' </summary>
        xlDialogEvaluateFormula = 709
        ''' <summary>
        ''' Represents the Excel built-in External Data Properties dialog.
        ''' </summary>
        xlDialogExternalDataProperties = 530
        ''' <summary>
        ''' Represents the Excel built-in Extract dialog.
        ''' </summary>
        xlDialogExtract = 35
        ''' <summary>
        ''' Represents the Excel built-in File Delete dialog.
        ''' </summary>
        xlDialogFileDelete = 6
        ''' <summary>
        ''' Represents the Excel built-in File Sharing dialog.
        ''' </summary>
        xlDialogFileSharing = 481
        ''' <summary>
        ''' Represents the Excel built-in Fill Group dialog.
        ''' </summary>
        xlDialogFillGroup = 200
        ''' <summary>
        ''' Represents the Excel built-in Fill Workgroup dialog.
        ''' </summary>
        xlDialogFillWorkgroup = 301
        ''' <summary>
        ''' Represents the Excel built-in Filter dialog.
        ''' </summary>
        xlDialogFilter = 447
        ''' <summary>
        ''' Represents the Excel built-in Filter Advanced dialog.
        ''' </summary>
        xlDialogFilterAdvanced = 370
        ''' <summary>
        ''' Represents the Excel built-in Find File dialog.
        ''' </summary>
        xlDialogFindFile = 475
        ''' <summary>
        ''' Represents the Excel built-in Font dialog.
        ''' </summary>
        xlDialogFont = 26
        ''' <summary>
        ''' Represents the Excel built-in Font Properties dialog.
        ''' </summary>
        xlDialogFontProperties = 381
        ''' <summary>
        ''' Represents the Excel built-in Format Auto dialog.
        ''' </summary>
        xlDialogFormatAuto = 269
        ''' <summary>
        ''' Represents the Excel built-in Format Chart dialog.
        ''' </summary>
        xlDialogFormatChart = 465
        ''' <summary>
        ''' Represents the Excel built-in Format Chart Type dialog.
        ''' </summary>
        xlDialogFormatCharttype = 423
        ''' <summary>
        ''' Represents the Excel built-in Format Font dialog.
        ''' </summary>
        xlDialogFormatFont = 150
        ''' <summary>
        ''' Represents the Excel built-in Format Legend dialog.
        ''' </summary>
        xlDialogFormatLegend = 88
        ''' <summary>
        ''' Represents the Excel built-in Format Main dialog.
        ''' </summary>
        xlDialogFormatMain = 225
        ''' <summary>
        ''' Represents the Excel built-in Format Move dialog.
        ''' </summary>
        xlDialogFormatMove = 128
        ''' <summary>
        ''' Represents the Excel built-in Format Number dialog.
        ''' </summary>
        xlDialogFormatNumber = 42
        ''' <summary>
        ''' Represents the Excel built-in Format Overlay dialog.
        ''' </summary>
        xlDialogFormatOverlay = 226
        ''' <summary>
        ''' Represents the Excel built-in Format Size dialog.
        ''' </summary>
        xlDialogFormatSize = 129
        ''' <summary>
        ''' Represents the Excel built-in Format Text dialog.
        ''' </summary>
        xlDialogFormatText = 89
        ''' <summary>
        ''' Represents the Excel built-in Formula Find dialog.
        ''' </summary>
        xlDialogFormulaFind = 64
        ''' <summary>
        ''' Represents the Excel built-in Formula Goto dialog.
        ''' </summary>
        xlDialogFormulaGoto = 63
        ''' <summary>
        ''' Represents the Excel built-in Formula Replace dialog.
        ''' </summary>
        xlDialogFormulaReplace = 130
        ''' <summary>
        ''' Represents the Excel built-in Function Wizard dialog.
        ''' </summary>
        xlDialogFunctionWizard = 450
        ''' <summary>
        ''' Represents the Excel built-in Gallery3d Area dialog.
        ''' </summary>
        xlDialogGallery3dArea = 193
        ''' <summary>
        ''' Represents the Excel built-in Gallery3d Bar dialog.
        ''' </summary>
        xlDialogGallery3dBar = 272
        ''' <summary>
        ''' Represents the Excel built-in Gallery3d Column dialog.
        ''' </summary>
        xlDialogGallery3dColumn = 194
        ''' <summary>
        ''' Represents the Excel built-in Gallery3d Line dialog.
        ''' </summary>
        xlDialogGallery3dLine = 195
        ''' <summary>
        ''' Represents the Excel built-in Gallery3d Pie dialog.
        ''' </summary>
        xlDialogGallery3dPie = 196
        ''' <summary>
        ''' Represents the Excel built-in Gallery3d Surface dialog.
        ''' </summary>
        xlDialogGallery3dSurface = 273
        ''' <summary>
        ''' Represents the Excel built-in Gallery Area dialog.
        ''' </summary>
        xlDialogGalleryArea = 67
        ''' <summary>
        ''' Represents the Excel built-in Gallery Bar dialog.
        ''' </summary>
        xlDialogGalleryBar = 68
        ''' <summary>
        ''' Represents the Excel built-in Gallery Column dialog.
        ''' </summary>
        xlDialogGalleryColumn = 69
        ''' <summary>
        ''' Represents the Excel built-in Gallery Custom dialog.
        ''' </summary>
        xlDialogGalleryCustom = 388
        ''' <summary>
        ''' Represents the Excel built-in Gallery Doughnut dialog.
        ''' </summary>
        xlDialogGalleryDoughnut = 344
        ''' <summary>
        ''' Represents the Excel built-in Gallery Line dialog.
        ''' </summary>
        xlDialogGalleryLine = 70
        ''' <summary>
        ''' Represents the Excel built-in Gallery Pie dialog.
        ''' </summary>
        xlDialogGalleryPie = 71
        ''' <summary>
        ''' Represents the Excel built-in Gallery Radar dialog.
        ''' </summary>
        xlDialogGalleryRadar = 249
        ''' <summary>
        ''' Represents the Excel built-in Gallery Scatter dialog.
        ''' </summary>
        xlDialogGalleryScatter = 72
        ''' <summary>
        ''' Represents the Excel built-in Goal Seek dialog.
        ''' </summary>
        xlDialogGoalSeek = 198
        ''' <summary>
        ''' Represents the Excel built-in Gridlines dialog.
        ''' </summary>
        xlDialogGridlines = 76
        ''' <summary>
        ''' Represents the Excel built-in Import Text File dialog.
        ''' </summary>
        xlDialogImportTextFile = 666
        ''' <summary>
        ''' Represents the Excel built-in Insert dialog.
        ''' </summary>
        xlDialogInsert = 55
        ''' <summary>
        ''' Represents the Excel built-in Insert Hyperlink dialog.
        ''' </summary>
        xlDialogInsertHyperlink = 596
        ''' <summary>
        ''' Represents the Excel built-in Insert Object dialog.
        ''' </summary>
        xlDialogInsertObject = 259
        ''' <summary>
        ''' Represents the Excel built-in Insert Picture dialog.
        ''' </summary>
        xlDialogInsertPicture = 342
        ''' <summary>
        ''' Represents the Excel built-in Insert Title dialog.
        ''' </summary>
        xlDialogInsertTitle = 380
        ''' <summary>
        ''' Represents the Excel built-in Label Properties dialog.
        ''' </summary>
        xlDialogLabelProperties = 436
        ''' <summary>
        ''' Represents the Excel built-in Listbox Properties dialog.
        ''' </summary>
        xlDialogListboxProperties = 437
        ''' <summary>
        ''' Represents the Excel built-in Macro Options dialog.
        ''' </summary>
        xlDialogMacroOptions = 382
        ''' <summary>
        ''' Represents the Excel built-in Mail Edit Mailer dialog.
        ''' </summary>
        xlDialogMailEditMailer = 470
        ''' <summary>
        ''' Represents the Excel built-in Mail Logon dialog.
        ''' </summary>
        xlDialogMailLogon = 339
        ''' <summary>
        ''' Represents the Excel built-in Mail Next Letter dialog.
        ''' </summary>
        xlDialogMailNextLetter = 378
        ''' <summary>
        ''' Represents the Excel built-in Main Chart dialog.
        ''' </summary>
        xlDialogMainChart = 85
        ''' <summary>
        ''' Represents the Excel built-in Main Chart Type dialog.
        ''' </summary>
        xlDialogMainChartType = 185
        ''' <summary>
        ''' Represents the Excel built-in Manage Relationships dialog.
        ''' </summary>
        xlDialogManageRelationships = 1271
        ''' <summary>
        ''' Represents the Excel built-in MEnu Editor dialog.
        ''' </summary>
        xlDialogMenuEditor = 322
        ''' <summary>
        ''' Represents the Excel built-in Move dialog.
        ''' </summary>
        xlDialogMove = 262
        ''' <summary>
        ''' Represents the Excel built-in My Permission dialog.
        ''' </summary>
        xlDialogMyPermission = 834
        ''' <summary>
        ''' Represents the Excel built-in Name Manager dialog.
        ''' </summary>
        xlDialogNameManager = 977
        ''' <summary>
        ''' Represents the Excel built-in New dialog.
        ''' </summary>
        xlDialogNew = 119
        ''' <summary>
        ''' Represents the Excel built-in New Name dialog.
        ''' </summary>
        xlDialogNewName = 978
        ''' <summary>
        ''' Represents the Excel built-in New Web Query dialog.
        ''' </summary>
        xlDialogNewWebQuery = 667
        ''' <summary>
        ''' Represents the Excel built-in Note dialog.
        ''' </summary>
        xlDialogNote = 154
        ''' <summary>
        ''' Represents the Excel built-in Object Properties dialog.
        ''' </summary>
        xlDialogObjectProperties = 207
        ''' <summary>
        ''' Represents the Excel built-in Object Protection dialog.
        ''' </summary>
        xlDialogObjectProtection = 214
        ''' <summary>
        ''' Represents the Excel built-in Open dialog.
        ''' </summary>
        xlDialogOpen = 1
        ''' <summary>
        ''' Represents the Excel built-in Open Links dialog.
        ''' </summary>
        xlDialogOpenLinks = 2
        ''' <summary>
        ''' Represents the Excel built-in Open Mail dialog.
        ''' </summary>
        xlDialogOpenMail = 188
        ''' <summary>
        ''' Represents the Excel built-in Open Text dialog.
        ''' </summary>
        xlDialogOpenText = 441
        ''' <summary>
        ''' Represents the Excel built-in Options Calculation dialog.
        ''' </summary>
        xlDialogOptionsCalculation = 318
        ''' <summary>
        ''' Represents the Excel built-in Options Chart dialog.
        ''' </summary>
        xlDialogOptionsChart = 325
        ''' <summary>
        ''' Represents the Excel built-in Options Edit dialog.
        ''' </summary>
        xlDialogOptionsEdit = 319
        ''' <summary>
        ''' Represents the Excel built-in Options General dialog.
        ''' </summary>
        xlDialogOptionsGeneral = 356
        ''' <summary>
        ''' Represents the Excel built-in Options Lists Add dialog.
        ''' </summary>
        xlDialogOptionsListsAdd = 458
        ''' <summary>
        ''' Represents the Excel built-in Options ME dialog.
        ''' </summary>
        xlDialogOptionsME = 647
        ''' <summary>
        ''' Represents the Excel built-in Options Transition dialog.
        ''' </summary>
        xlDialogOptionsTransition = 355
        ''' <summary>
        ''' Represents the Excel built-in Options View dialog.
        ''' </summary>
        xlDialogOptionsView = 320
        ''' <summary>
        ''' Represents the Excel built-in Outline dialog.
        ''' </summary>
        xlDialogOutline = 142
        ''' <summary>
        ''' Represents the Excel built-in Overlay dialog.
        ''' </summary>
        xlDialogOverlay = 86
        ''' <summary>
        ''' Represents the Excel built-in Overlay Chart Type dialog.
        ''' </summary>
        xlDialogOverlayChartType = 186
        ''' <summary>
        ''' Represents the Excel built-in Page Setup dialog.
        ''' </summary>
        xlDialogPageSetup = 7
        ''' <summary>
        ''' Represents the Excel built-in Parse dialog.
        ''' </summary>
        xlDialogParse = 91
        ''' <summary>
        ''' Represents the Excel built-in Paste Names dialog.
        ''' </summary>
        xlDialogPasteNames = 58
        ''' <summary>
        ''' Represents the Excel built-in Paste Special dialog.
        ''' </summary>
        xlDialogPasteSpecial = 53
        ''' <summary>
        ''' Represents the Excel built-in Patterns dialog.
        ''' </summary>
        xlDialogPatterns = 84
        ''' <summary>
        ''' Represents the Excel built-in Permission dialog.
        ''' </summary>
        xlDialogPermission = 832
        ''' <summary>
        ''' Represents the Excel built-in Phonetic dialog.
        ''' </summary>
        xlDialogPhonetic = 656
        ''' <summary>
        ''' Represents the Excel built-in Pivot Calculated Field dialog.
        ''' </summary>
        xlDialogPivotCalculatedField = 570
        ''' <summary>
        ''' Represents the Excel built-in Pivot Calculated Item dialog.
        ''' </summary>
        xlDialogPivotCalculatedItem = 572
        ''' <summary>
        ''' Represents the Excel built-in Pivot Client Server Set dialog.
        ''' </summary>
        xlDialogPivotClientServerSet = 689
        ''' <summary>
        ''' Represents the Excel built-in Pivot Field Group dialog.
        ''' </summary>
        xlDialogPivotFieldGroup = 433
        ''' <summary>
        ''' Represents the Excel built-in Pivot Field Properties dialog.
        ''' </summary>
        xlDialogPivotFieldProperties = 313
        ''' <summary>
        ''' Represents the Excel built-in Pivot Field Ungroup dialog.
        ''' </summary>
        xlDialogPivotFieldUngroup = 434
        ''' <summary>
        ''' Represents the Excel built-in Pivot Show Pages dialog.
        ''' </summary>
        xlDialogPivotShowPages = 421
        ''' <summary>
        ''' Represents the Excel built-in Pivot Solve Order dialog.
        ''' </summary>
        xlDialogPivotSolveOrder = 568
        ''' <summary>
        ''' Represents the Excel built-in Pivot Table Options dialog.
        ''' </summary>
        xlDialogPivotTableOptions = 567
        ''' <summary>
        ''' Represents the Excel built-in Pivot Table Slicer Connections dialog.
        ''' </summary>
        xlDialogPivotTableSlicerConnections = 1183
        ''' <summary>
        ''' Represents the Excel built-in Pivot Table What If Analysis Settings dialog.
        ''' </summary>
        xlDialogPivotTableWhatIfAnalysisSettings = 1153
        ''' <summary>
        ''' Represents the Excel built-in Pivot Table Wizard dialog.
        ''' </summary>
        xlDialogPivotTableWizard = 312
        ''' <summary>
        ''' Represents the Excel built-in Placement dialog.
        ''' </summary>
        xlDialogPlacement = 300
        ''' <summary>
        ''' Represents the Excel built-in Print dialog.
        ''' </summary>
        xlDialogPrint = 8
        ''' <summary>
        ''' Represents the Excel built-in Printer Setup dialog.
        ''' </summary>
        xlDialogPrinterSetup = 9
        ''' <summary>
        ''' Represents the Excel built-in Print Preview dialog.
        ''' </summary>
        xlDialogPrintPreview = 222
        ''' <summary>
        ''' Represents the Excel built-in Promote dialog.
        ''' </summary>
        xlDialogPromote = 202
        ''' <summary>
        ''' Represents the Excel built-in Properties dialog.
        ''' </summary>
        xlDialogProperties = 474
        ''' <summary>
        ''' Represents the Excel built-in Property Fields dialog.
        ''' </summary>
        xlDialogPropertyFields = 754
        ''' <summary>
        ''' Represents the Excel built-in Protect Document dialog.
        ''' </summary>
        xlDialogProtectDocument = 28
        ''' <summary>
        ''' Represents the Excel built-in Protect Sharing dialog.
        ''' </summary>
        xlDialogProtectSharing = 620
        ''' <summary>
        ''' Represents the Excel built-in Publish As Web Page dialog.
        ''' </summary>
        xlDialogPublishAsWebPage = 653
        ''' <summary>
        ''' Represents the Excel built-in Pushbutton Properties dialog.
        ''' </summary>
        xlDialogPushbuttonProperties = 445
        ''' <summary>
        ''' Represents the Excel built-in Recommended Pivot Tables dialog.
        ''' </summary>
        xlDialogRecommendedPivotTables = 1258
        ''' <summary>
        ''' Represents the Excel built-in Replace Font dialog.
        ''' </summary>
        xlDialogReplaceFont = 134
        ''' <summary>
        ''' Represents the Excel built-in Routing Slip dialog.
        ''' </summary>
        xlDialogRoutingSlip = 336
        ''' <summary>
        ''' Represents the Excel built-in Row Height dialog.
        ''' </summary>
        xlDialogRowHeight = 127
        ''' <summary>
        ''' Represents the Excel built-in Run dialog.
        ''' </summary>
        xlDialogRun = 17
        ''' <summary>
        ''' Represents the Excel built-in Save As dialog.
        ''' </summary>
        xlDialogSaveAs = 5
        ''' <summary>
        ''' Represents the Excel built-in Save Copy As dialog.
        ''' </summary>
        xlDialogSaveCopyAs = 456
        ''' <summary>
        ''' Represents the Excel built-in Save New Object dialog.
        ''' </summary>
        xlDialogSaveNewObject = 208
        ''' <summary>
        ''' Represents the Excel built-in Save Workbook dialog.
        ''' </summary>
        xlDialogSaveWorkbook = 145
        ''' <summary>
        ''' Represents the Excel built-in Save Workspace dialog.
        ''' </summary>
        xlDialogSaveWorkspace = 285
        ''' <summary>
        ''' Represents the Excel built-in Scale dialog.
        ''' </summary>
        xlDialogScale = 87
        ''' <summary>
        ''' Represents the Excel built-in Scenario Add dialog.
        ''' </summary>
        xlDialogScenarioAdd = 307
        ''' <summary>
        ''' Represents the Excel built-in Scenario Cells dialog.
        ''' </summary>
        xlDialogScenarioCells = 305
        ''' <summary>
        ''' Represents the Excel built-in Scenario Edit dialog.
        ''' </summary>
        xlDialogScenarioEdit = 308
        ''' <summary>
        ''' Represents the Excel built-in Scenario MErge dialog.
        ''' </summary>
        xlDialogScenarioMerge = 473
        ''' <summary>
        ''' Represents the Excel built-in Scenario Summary dialog.
        ''' </summary>
        xlDialogScenarioSummary = 311
        ''' <summary>
        ''' Represents the Excel built-in Scrollbar Properties dialog.
        ''' </summary>
        xlDialogScrollbarProperties = 420
        ''' <summary>
        ''' Represents the Excel built-in Search dialog.
        ''' </summary>
        xlDialogSearch = 731
        ''' <summary>
        ''' Represents the Excel built-in Select Special dialog.
        ''' </summary>
        xlDialogSelectSpecial = 132
        ''' <summary>
        ''' Represents the Excel built-in Send Mail dialog.
        ''' </summary>
        xlDialogSendMail = 189
        ''' <summary>
        ''' Represents the Excel built-in Series Axes dialog.
        ''' </summary>
        xlDialogSeriesAxes = 460
        ''' <summary>
        ''' Represents the Excel built-in Series Options dialog.
        ''' </summary>
        xlDialogSeriesOptions = 557
        ''' <summary>
        ''' Represents the Excel built-in Series Order dialog.
        ''' </summary>
        xlDialogSeriesOrder = 466
        ''' <summary>
        ''' Represents the Excel built-in Series Shape dialog.
        ''' </summary>
        xlDialogSeriesShape = 504
        ''' <summary>
        ''' Represents the Excel built-in Series X dialog.
        ''' </summary>
        xlDialogSeriesX = 461
        ''' <summary>
        ''' Represents the Excel built-in Series Y dialog.
        ''' </summary>
        xlDialogSeriesY = 462
        ''' <summary>
        ''' Represents the Excel built-in Set Background Picture dialog.
        ''' </summary>
        xlDialogSetBackgroundPicture = 509
        ''' <summary>
        ''' Represents the Excel built-in Set Manager dialog.
        ''' </summary>
        xlDialogSetManager = 1109
        ''' <summary>
        ''' Represents the Excel built-in Set MDXEditor dialog.
        ''' </summary>
        xlDialogSetMDXEditor = 1208
        ''' <summary>
        ''' Represents the Excel built-in Set Print Titles dialog.
        ''' </summary>
        xlDialogSetPrintTitles = 23
        ''' <summary>
        ''' Represents the Excel built-in Set Tuple Editor On Columns dialog.
        ''' </summary>
        xlDialogSetTupleEditorOnColumns = 1108
        ''' <summary>
        ''' Represents the Excel built-in Set Tuple Editor On Rows dialog.
        ''' </summary>
        xlDialogSetTupleEditorOnRows = 1107
        ''' <summary>
        ''' Represents the Excel built-in Set Update Status dialog.
        ''' </summary>
        xlDialogSetUpdateStatus = 159
        ''' <summary>
        ''' Represents the Excel built-in Show Detail dialog.
        ''' </summary>
        xlDialogShowDetail = 204
        ''' <summary>
        ''' Represents the Excel built-in Show Toolbar dialog.
        ''' </summary>
        xlDialogShowToolbar = 220
        ''' <summary>
        ''' Represents the Excel built-in Size dialog.
        ''' </summary>
        xlDialogSize = 261
        ''' <summary>
        ''' Represents the Excel built-in Slicer Creation dialog.
        ''' </summary>
        xlDialogSlicerCreation = 1182
        ''' <summary>
        ''' Represents the Excel built-in Slicer Pivot Table Connections dialog.
        ''' </summary>
        xlDialogSlicerPivotTableConnections = 1184
        ''' <summary>
        ''' Represents the Excel built-in Slicer Settings dialog.
        ''' </summary>
        xlDialogSlicerSettings = 1179
        ''' <summary>
        ''' Represents the Excel built-in Sort dialog.
        ''' </summary>
        xlDialogSort = 39
        ''' <summary>
        ''' Represents the Excel built-in Sort Special dialog.
        ''' </summary>
        xlDialogSortSpecial = 192
        ''' <summary>
        ''' Represents the Excel built-in Sparkline Insert Column dialog.
        ''' </summary>
        xlDialogSparklineInsertColumn = 1134
        ''' <summary>
        ''' Represents the Excel built-in Sparkline Insert Line dialog.
        ''' </summary>
        xlDialogSparklineInsertLine = 1133
        ''' <summary>
        ''' Represents the Excel built-in Sparkline Insert Win Loss dialog.
        ''' </summary>
        xlDialogSparklineInsertWinLoss = 1135
        ''' <summary>
        ''' Represents the Excel built-in Split dialog.
        ''' </summary>
        xlDialogSplit = 137
        ''' <summary>
        ''' Represents the Excel built-in Standard Font dialog.
        ''' </summary>
        xlDialogStandardFont = 190
        ''' <summary>
        ''' Represents the Excel built-in Standard Width dialog.
        ''' </summary>
        xlDialogStandardWidth = 472
        ''' <summary>
        ''' Represents the Excel built-in Style dialog.
        ''' </summary>
        xlDialogStyle = 44
        ''' <summary>
        ''' Represents the Excel built-in Subscribe To dialog.
        ''' </summary>
        xlDialogSubscribeTo = 218
        ''' <summary>
        ''' Represents the Excel built-in Subtotal Create dialog.
        ''' </summary>
        xlDialogSubtotalCreate = 398
#Disable Warning CA1069 ' Enums values should not be duplicated
        ''' <summary>
        ''' Represents the Excel built-in Summary Info dialog.
        ''' </summary>
        xlDialogSummaryInfo = 474
#Enable Warning CA1069 ' Enums values should not be duplicated
        ''' <summary>
        ''' Represents the Excel built-in Table dialog.
        ''' </summary>
        xlDialogTable = 41
        ''' <summary>
        ''' Represents the Excel built-in Tab Order dialog.
        ''' </summary>
        xlDialogTabOrder = 394
        ''' <summary>
        ''' Represents the Excel built-in Text To Columns dialog.
        ''' </summary>
        xlDialogTextToColumns = 422
        ''' <summary>
        ''' Represents the Excel built-in Unhide dialog.
        ''' </summary>
        xlDialogUnhide = 94
        ''' <summary>
        ''' Represents the Excel built-in Update Link dialog.
        ''' </summary>
        xlDialogUpdateLink = 201
        ''' <summary>
        ''' Represents the Excel built-in VBA Insert File dialog.
        ''' </summary>
        xlDialogVbaInsertFile = 328
        ''' <summary>
        ''' Represents the Excel built-in VBA Make Addin dialog.
        ''' </summary>
        xlDialogVbaMakeAddin = 478
        ''' <summary>
        ''' Represents the Excel built-in VBA Procedure Definition dialog.
        ''' </summary>
        xlDialogVbaProcedureDefinition = 330
        ''' <summary>
        ''' Represents the Excel built-in View3d dialog.
        ''' </summary>
        xlDialogView3d = 197
        ''' <summary>
        ''' Represents the Excel built-in Web Options Browsers dialog.
        ''' </summary>
        xlDialogWebOptionsBrowsers = 773
        ''' <summary>
        ''' Represents the Excel built-in Web Options Encoding dialog.
        ''' </summary>
        xlDialogWebOptionsEncoding = 686
        ''' <summary>
        ''' Represents the Excel built-in Web Options Files dialog.
        ''' </summary>
        xlDialogWebOptionsFiles = 684
        ''' <summary>
        ''' Represents the Excel built-in Web Options Fonts dialog.
        ''' </summary>
        xlDialogWebOptionsFonts = 687
        ''' <summary>
        ''' Represents the Excel built-in Web Options General dialog.
        ''' </summary>
        xlDialogWebOptionsGeneral = 683
        ''' <summary>
        ''' Represents the Excel built-in Web Options Pictures dialog.
        ''' </summary>
        xlDialogWebOptionsPictures = 685
        ''' <summary>
        ''' Represents the Excel built-in Window Move dialog.
        ''' </summary>
        xlDialogWindowMove = 14
        ''' <summary>
        ''' Represents the Excel built-in Window Size dialog.
        ''' </summary>
        xlDialogWindowSize = 13
        ''' <summary>
        ''' Represents the Excel built-in Workbook Add dialog.
        ''' </summary>
        xlDialogWorkbookAdd = 281
        ''' <summary>
        ''' Represents the Excel built-in Workbook Copy dialog.
        ''' </summary>
        xlDialogWorkbookCopy = 283
        ''' <summary>
        ''' Represents the Excel built-in Workbook Insert dialog.
        ''' </summary>
        xlDialogWorkbookInsert = 354
        ''' <summary>
        ''' Represents the Excel built-in Workbook Move dialog.
        ''' </summary>
        xlDialogWorkbookMove = 282
        ''' <summary>
        ''' Represents the Excel built-in Workbook Name dialog.
        ''' </summary>
        xlDialogWorkbookName = 386
        ''' <summary>
        ''' Represents the Excel built-in Workbook New dialog.
        ''' </summary>
        xlDialogWorkbookNew = 302
        ''' <summary>
        ''' Represents the Excel built-in Workbook Options dialog.
        ''' </summary>
        xlDialogWorkbookOptions = 284
        ''' <summary>
        ''' Represents the Excel built-in Workbook Protect dialog.
        ''' </summary>
        xlDialogWorkbookProtect = 417
        ''' <summary>
        ''' Represents the Excel built-in Workbook Tab Split dialog.
        ''' </summary>
        xlDialogWorkbookTabSplit = 415
        ''' <summary>
        ''' Represents the Excel built-in Workbook Unhide dialog.
        ''' </summary>
        xlDialogWorkbookUnhide = 384
        ''' <summary>
        ''' Represents the Excel built-in Workgroup dialog.
        ''' </summary>
        xlDialogWorkgroup = 199
        ''' <summary>
        ''' Represents the Excel built-in Workspace dialog.
        ''' </summary>
        xlDialogWorkspace = 95
        ''' <summary>
        ''' Represents the Excel built-in Zoom dialog.
        ''' </summary>
        xlDialogZoom = 256
    End Enum

End Class
