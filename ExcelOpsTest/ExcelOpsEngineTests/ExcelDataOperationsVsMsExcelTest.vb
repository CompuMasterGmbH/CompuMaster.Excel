Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps

Namespace ExcelOpsEngineTests
    <TestFixture> Public Class ExcelDataOperationsTest

        <OneTimeSetUp()> Public Sub OneTimeInitConfig()
        End Sub

        <SetUp> Public Sub Setup()
            Test.Console.ResetConsoleForTestOutput()
        End Sub

        <Test> Public Sub SheetContentMatrix()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim MsExcelApp As New MsExcelDataOperations.MsAppInstance
            Try
                Dim ExpectedMatrix As String
                Dim TestControllingToolFileName As String = TestEnvironment.TestFiles.TestFileV0SRH.FullName
                Dim TestSheet As String = "Grunddaten"

                eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
                mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, True)

                ExpectedMatrix =
                         "# |A                           |B        |C  |D     |E    " & vbNewLine &
                         "--+----------------------------+---------+---+------+-----" & vbNewLine &
                         "1 |Jahr                        |2019     |   |      |False" & vbNewLine &
                         "2 |Geschäftsjahr von           |         |bis|      |     " & vbNewLine &
                         "3 |Aktueller Monat             |1        |   |Januar|     " & vbNewLine &
                         "4 |                            |         |   |      |     " & vbNewLine &
                         "5 |Name Betrieb                |Test     |   |      |     " & vbNewLine &
                         "6 |                            |         |   |      |     " & vbNewLine &
                         "7 |Arbeitgeberanteile in %     |         |   |      |     " & vbNewLine &
                         "8 |Chef: 14,09                 |         |   |      |     " & vbNewLine &
                         "9 |Büroangestellte: 20,00      |         |   |      |     " & vbNewLine &
                         "10|Produktivkraft: 25,00       |         |   |      |     " & vbNewLine &
                         "11|Azubi / Aushilfen: 33,00    |         |   |      |     " & vbNewLine &
                         "12|                            |         |   |      |     " & vbNewLine &
                         "13|Berechnung Jahresarbeitszeit|         |   |      |     " & vbNewLine &
                         "14|Tage / Jahr:                |365      |   |      |     " & vbNewLine &
                         "15|Wochenendtage               |104      |   |      |     " & vbNewLine &
                         "16|=Zahltage:                  |261      |   |      |     " & vbNewLine &
                         "17|Wochenarbeitszeit           |40       |   |      |     " & vbNewLine &
                         "18|Tagesarbeitszeit:           |8        |   |      |     " & vbNewLine &
                         "19|Normallohnstunden / Jahr:   |2.088,00 |   |      |     " & vbNewLine &
                         "20|                            |         |   |      |     " & vbNewLine &
                         "21|                            |         |   |      |     " & vbNewLine &
                         "22|                            |         |   |      |     " & vbNewLine &
                         "23|1                           |Januar   |   |      |     " & vbNewLine &
                         "24|2                           |Februar  |   |      |     " & vbNewLine &
                         "25|3                           |März     |   |      |     " & vbNewLine &
                         "26|4                           |April    |   |      |     " & vbNewLine &
                         "27|5                           |Mai      |   |      |     " & vbNewLine &
                         "28|6                           |Juni     |   |      |     " & vbNewLine &
                         "29|7                           |Juli     |   |      |     " & vbNewLine &
                         "30|8                           |August   |   |      |     " & vbNewLine &
                         "31|9                           |September|   |      |     " & vbNewLine &
                         "32|10                          |Oktober  |   |      |     " & vbNewLine &
                         "33|11                          |November |   |      |     " & vbNewLine &
                         "34|12                          |Dezember |   |      |     " & vbNewLine &
                         "35|Zusammensetzung AG Anteile  |         |   |      |     " & vbNewLine &
                         "36|Krankenkasse                |2,8      |   |      |     " & vbNewLine &
                         "37|Rentenkasse                 |8        |   |      |     " & vbNewLine &
                         "38|Pflegekasse                 |1,4      |   |      |     " & vbNewLine &
                         "39|Krankengeld                 |0,25     |   |      |     " & vbNewLine &
                         "40|                            |12,45    |   |      |     " & vbNewLine
                Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues, ExpectedMatrix)
                Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues, ExpectedMatrix)

                ExpectedMatrix =
                         "# |A                           |B        |C  |D |E    " & vbNewLine &
                         "--+----------------------------+---------+---+--+-----" & vbNewLine &
                         "1 |Jahr                        |2019     |   |  |False" & vbNewLine &
                         "2 |Geschäftsjahr von           |         |bis|  |     " & vbNewLine &
                         "3 |Aktueller Monat             |1        |   |  |     " & vbNewLine &
                         "4 |                            |         |   |  |     " & vbNewLine &
                         "5 |Name Betrieb                |Test     |   |  |     " & vbNewLine &
                         "6 |                            |         |   |  |     " & vbNewLine &
                         "7 |Arbeitgeberanteile in %     |         |   |  |     " & vbNewLine &
                         "8 |Chef: 14,09                 |         |   |  |     " & vbNewLine &
                         "9 |Büroangestellte: 20,00      |         |   |  |     " & vbNewLine &
                         "10|Produktivkraft: 25,00       |         |   |  |     " & vbNewLine &
                         "11|Azubi / Aushilfen: 33,00    |         |   |  |     " & vbNewLine &
                         "12|                            |         |   |  |     " & vbNewLine &
                         "13|Berechnung Jahresarbeitszeit|         |   |  |     " & vbNewLine &
                         "14|Tage / Jahr:                |365      |   |  |     " & vbNewLine &
                         "15|Wochenendtage               |         |   |  |     " & vbNewLine &
                         "16|=Zahltage:                  |         |   |  |     " & vbNewLine &
                         "17|Wochenarbeitszeit           |40       |   |  |     " & vbNewLine &
                         "18|Tagesarbeitszeit:           |         |   |  |     " & vbNewLine &
                         "19|Normallohnstunden / Jahr:   |         |   |  |     " & vbNewLine &
                         "20|                            |         |   |  |     " & vbNewLine &
                         "21|                            |         |   |  |     " & vbNewLine &
                         "22|                            |         |   |  |     " & vbNewLine &
                         "23|1                           |Januar   |   |  |     " & vbNewLine &
                         "24|2                           |Februar  |   |  |     " & vbNewLine &
                         "25|3                           |März     |   |  |     " & vbNewLine &
                         "26|4                           |April    |   |  |     " & vbNewLine &
                         "27|5                           |Mai      |   |  |     " & vbNewLine &
                         "28|6                           |Juni     |   |  |     " & vbNewLine &
                         "29|7                           |Juli     |   |  |     " & vbNewLine &
                         "30|8                           |August   |   |  |     " & vbNewLine &
                         "31|9                           |September|   |  |     " & vbNewLine &
                         "32|10                          |Oktober  |   |  |     " & vbNewLine &
                         "33|11                          |November |   |  |     " & vbNewLine &
                         "34|12                          |Dezember |   |  |     " & vbNewLine &
                         "35|Zusammensetzung AG Anteile  |         |   |  |     " & vbNewLine &
                         "36|Krankenkasse                |2,8      |   |  |     " & vbNewLine &
                         "37|Rentenkasse                 |8        |   |  |     " & vbNewLine &
                         "38|Pflegekasse                 |1,4      |   |  |     " & vbNewLine &
                         "39|Krankengeld                 |0,25     |   |  |     " & vbNewLine
                Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticValues, ExpectedMatrix)
                Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticValues, ExpectedMatrix)

                ExpectedMatrix =
                         "# |A |B            |C |D                                  " & vbNewLine &
                         "--+--+-------------+--+-----------------------------------" & vbNewLine &
                         "1 |  |             |  |                                   " & vbNewLine &
                         "2 |  |             |  |                                   " & vbNewLine &
                         "3 |  |             |  |=INDEX(B23:B34,MATCH(B3,A23:A34,0))" & vbNewLine &
                         "4 |  |             |  |                                   " & vbNewLine &
                         "5 |  |             |  |                                   " & vbNewLine &
                         "6 |  |             |  |                                   " & vbNewLine &
                         "7 |  |             |  |                                   " & vbNewLine &
                         "8 |  |             |  |                                   " & vbNewLine &
                         "9 |  |             |  |                                   " & vbNewLine &
                         "10|  |             |  |                                   " & vbNewLine &
                         "11|  |             |  |                                   " & vbNewLine &
                         "12|  |             |  |                                   " & vbNewLine &
                         "13|  |             |  |                                   " & vbNewLine &
                         "14|  |             |  |                                   " & vbNewLine &
                         "15|  |=2*52        |  |                                   " & vbNewLine &
                         "16|  |=B14-B15     |  |                                   " & vbNewLine &
                         "17|  |             |  |                                   " & vbNewLine &
                         "18|  |=B17/5       |  |                                   " & vbNewLine &
                         "19|  |=B18*B16     |  |                                   " & vbNewLine &
                         "20|  |             |  |                                   " & vbNewLine &
                         "21|  |             |  |                                   " & vbNewLine &
                         "22|  |             |  |                                   " & vbNewLine &
                         "23|  |             |  |                                   " & vbNewLine &
                         "24|  |             |  |                                   " & vbNewLine &
                         "25|  |             |  |                                   " & vbNewLine &
                         "26|  |             |  |                                   " & vbNewLine &
                         "27|  |             |  |                                   " & vbNewLine &
                         "28|  |             |  |                                   " & vbNewLine &
                         "29|  |             |  |                                   " & vbNewLine &
                         "30|  |             |  |                                   " & vbNewLine &
                         "31|  |             |  |                                   " & vbNewLine &
                         "32|  |             |  |                                   " & vbNewLine &
                         "33|  |             |  |                                   " & vbNewLine &
                         "34|  |             |  |                                   " & vbNewLine &
                         "35|  |             |  |                                   " & vbNewLine &
                         "36|  |             |  |                                   " & vbNewLine &
                         "37|  |             |  |                                   " & vbNewLine &
                         "38|  |             |  |                                   " & vbNewLine &
                         "39|  |             |  |                                   " & vbNewLine &
                         "40|  |=SUM(B36:B39)|  |                                   " & vbNewLine
                Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.Formulas, ExpectedMatrix)
                Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.Formulas, ExpectedMatrix)

                ExpectedMatrix =
                         "# |A                           |B        |C  |D     |E    " & vbNewLine &
                         "--+----------------------------+---------+---+------+-----" & vbNewLine &
                         "1 |Jahr                        |2019     |   |      |False" & vbNewLine &
                         "2 |Geschäftsjahr von           |         |bis|      |     " & vbNewLine &
                         "3 |Aktueller Monat             |1        |   |Januar|     " & vbNewLine &
                         "4 |                            |         |   |      |     " & vbNewLine &
                         "5 |Name Betrieb                |Test     |   |      |     " & vbNewLine &
                         "6 |                            |         |   |      |     " & vbNewLine &
                         "7 |Arbeitgeberanteile in %     |         |   |      |     " & vbNewLine &
                         "8 |Chef: 14,09                 |         |   |      |     " & vbNewLine &
                         "9 |Büroangestellte: 20,00      |         |   |      |     " & vbNewLine &
                         "10|Produktivkraft: 25,00       |         |   |      |     " & vbNewLine &
                         "11|Azubi / Aushilfen: 33,00    |         |   |      |     " & vbNewLine &
                         "12|                            |         |   |      |     " & vbNewLine &
                         "13|Berechnung Jahresarbeitszeit|         |   |      |     " & vbNewLine &
                         "14|Tage / Jahr:                |365      |   |      |     " & vbNewLine &
                         "15|Wochenendtage               |104      |   |      |     " & vbNewLine &
                         "16|=Zahltage:                  |261      |   |      |     " & vbNewLine &
                         "17|Wochenarbeitszeit           |40       |   |      |     " & vbNewLine &
                         "18|Tagesarbeitszeit:           |8        |   |      |     " & vbNewLine &
                         "19|Normallohnstunden / Jahr:   |2.088,00 |   |      |     " & vbNewLine &
                         "20|                            |         |   |      |     " & vbNewLine &
                         "21|                            |         |   |      |     " & vbNewLine &
                         "22|                            |         |   |      |     " & vbNewLine &
                         "23|1                           |Januar   |   |      |     " & vbNewLine &
                         "24|2                           |Februar  |   |      |     " & vbNewLine &
                         "25|3                           |März     |   |      |     " & vbNewLine &
                         "26|4                           |April    |   |      |     " & vbNewLine &
                         "27|5                           |Mai      |   |      |     " & vbNewLine &
                         "28|6                           |Juni     |   |      |     " & vbNewLine &
                         "29|7                           |Juli     |   |      |     " & vbNewLine &
                         "30|8                           |August   |   |      |     " & vbNewLine &
                         "31|9                           |September|   |      |     " & vbNewLine &
                         "32|10                          |Oktober  |   |      |     " & vbNewLine &
                         "33|11                          |November |   |      |     " & vbNewLine &
                         "34|12                          |Dezember |   |      |     " & vbNewLine &
                         "35|Zusammensetzung AG Anteile  |         |   |      |     " & vbNewLine &
                         "36|Krankenkasse                |2,8      |   |      |     " & vbNewLine &
                         "37|Rentenkasse                 |8        |   |      |     " & vbNewLine &
                         "38|Pflegekasse                 |1,4      |   |      |     " & vbNewLine &
                         "39|Krankengeld                 |0,25     |   |      |     " & vbNewLine &
                         "40|                            |12,45    |   |      |     " & vbNewLine
                Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormattedText, ExpectedMatrix)
                Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormattedText, ExpectedMatrix)

                ExpectedMatrix =
                         "# |A                           |B            |C  |D                                  |E    " & vbNewLine &
                         "--+----------------------------+-------------+---+-----------------------------------+-----" & vbNewLine &
                         "1 |Jahr                        |2019         |   |                                   |False" & vbNewLine &
                         "2 |Geschäftsjahr von           |             |bis|                                   |     " & vbNewLine &
                         "3 |Aktueller Monat             |1            |   |=INDEX(B23:B34,MATCH(B3,A23:A34,0))|     " & vbNewLine &
                         "4 |                            |             |   |                                   |     " & vbNewLine &
                         "5 |Name Betrieb                |Test         |   |                                   |     " & vbNewLine &
                         "6 |                            |             |   |                                   |     " & vbNewLine &
                         "7 |Arbeitgeberanteile in %     |             |   |                                   |     " & vbNewLine &
                         "8 |Chef: 14,09                 |             |   |                                   |     " & vbNewLine &
                         "9 |Büroangestellte: 20,00      |             |   |                                   |     " & vbNewLine &
                         "10|Produktivkraft: 25,00       |             |   |                                   |     " & vbNewLine &
                         "11|Azubi / Aushilfen: 33,00    |             |   |                                   |     " & vbNewLine &
                         "12|                            |             |   |                                   |     " & vbNewLine &
                         "13|Berechnung Jahresarbeitszeit|             |   |                                   |     " & vbNewLine &
                         "14|Tage / Jahr:                |365          |   |                                   |     " & vbNewLine &
                         "15|Wochenendtage               |=2*52        |   |                                   |     " & vbNewLine &
                         "16|=Zahltage:                  |=B14-B15     |   |                                   |     " & vbNewLine &
                         "17|Wochenarbeitszeit           |40           |   |                                   |     " & vbNewLine &
                         "18|Tagesarbeitszeit:           |=B17/5       |   |                                   |     " & vbNewLine &
                         "19|Normallohnstunden / Jahr:   |=B18*B16     |   |                                   |     " & vbNewLine &
                         "20|                            |             |   |                                   |     " & vbNewLine &
                         "21|                            |             |   |                                   |     " & vbNewLine &
                         "22|                            |             |   |                                   |     " & vbNewLine &
                         "23|1                           |Januar       |   |                                   |     " & vbNewLine &
                         "24|2                           |Februar      |   |                                   |     " & vbNewLine &
                         "25|3                           |März         |   |                                   |     " & vbNewLine &
                         "26|4                           |April        |   |                                   |     " & vbNewLine &
                         "27|5                           |Mai          |   |                                   |     " & vbNewLine &
                         "28|6                           |Juni         |   |                                   |     " & vbNewLine &
                         "29|7                           |Juli         |   |                                   |     " & vbNewLine &
                         "30|8                           |August       |   |                                   |     " & vbNewLine &
                         "31|9                           |September    |   |                                   |     " & vbNewLine &
                         "32|10                          |Oktober      |   |                                   |     " & vbNewLine &
                         "33|11                          |November     |   |                                   |     " & vbNewLine &
                         "34|12                          |Dezember     |   |                                   |     " & vbNewLine &
                         "35|Zusammensetzung AG Anteile  |             |   |                                   |     " & vbNewLine &
                         "36|Krankenkasse                |2,8          |   |                                   |     " & vbNewLine &
                         "37|Rentenkasse                 |8            |   |                                   |     " & vbNewLine &
                         "38|Pflegekasse                 |1,4          |   |                                   |     " & vbNewLine &
                         "39|Krankengeld                 |0,25         |   |                                   |     " & vbNewLine &
                         "40|                            |=SUM(B36:B39)|   |                                   |     " & vbNewLine
                Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText, ExpectedMatrix)
                Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText, ExpectedMatrix)
            Finally
                MsExcelDataOperations.PrepareCloseExcelAppInstance(MsExcelApp)
                MsExcelDataOperations.SafelyCloseExcelAppInstance(MsExcelApp)
            End Try
        End Sub

        <Test> Public Sub LookupCellValue()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestEnvironment.TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, True)
            Try
                '## Expected matrix like following
                '"# |A                           |B              |C  |D                                  |E     
                '"--+----------------------------+---------------+---+-----------------------------------+----- 
                '"1 |Jahr                        |2019           |   |                                   |False 
                '"2 |Geschäftsjahr von           |               |bis|                                   |      
                '"3 |Aktueller Monat             |1              |   |=INDEX(B23:B34,MATCH(B3,A23:A34,0))|      
                '"4 |                            |               |   |                                   |      
                '"5 |Name Betrieb                |Test           |   |                                   |      
                '"6 |                            |               |   |                                   |      
                '"7 |Arbeitgeberanteile In %     |               |   |                                   |      
                '"8 |Chef: 14,09                 |               |   |                                   |      
                '"9 |Büroangestellte: 20,00      |               |   |                                   |      
                '"10|Produktivkraft: 25,00       |               |   |                                   |      
                '"11|Azubi / Aushilfen: 33,00    |               |   |                                   |      
                '"12|                            |               |   |                                   |      
                '"13|Berechnung Jahresarbeitszeit|               |   |                                   |      
                '"14|Tage / Jahr:                |365            |   |                                   |      
                '"15|Wochenendtage               |=2*52          |   |                                   |      
                '"16|=Zahltage:                  |=B14-B15       |   |                                   |      
                '"17|Wochenarbeitszeit           |40             |   |                                   |      
                '"18|Tagesarbeitszeit:           |=B17/5         |   |                                   |      
                '"19|Normallohnstunden / Jahr:   |=B18*B16       |   |                                   |      
                '"20|                            |               |   |                                   |      
                '"36|Krankenkasse                |2,8            |   |                                   |      
                '"37|Rentenkasse                 |8              |   |                                   |      
                '"38|Pflegekasse                 |1,4            |   |                                   |      
                '"39|Krankengeld                 |0,25           |   |                                   |      
                '"40|                            |=SUMME(B36:B39)|   |                                   |      

                'D3
                Assert.AreEqual("Januar", eppeo.LookupCellValue(Of String)(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual("Januar", mseo.LookupCellValue(Of String)(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual("Januar", eppeo.LookupCellFormattedText(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual("Januar", mseo.LookupCellFormattedText(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual("INDEX(B23:B34,MATCH(B3,A23:A34,0))", eppeo.LookupCellFormula(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual("INDEX(B23:B34,MATCH(B3,A23:A34,0))", mseo.LookupCellFormula(New ExcelOps.ExcelCell(TestSheet, "D3", ExcelOps.ExcelCell.ValueTypes.All)))

                'A8
                Assert.AreEqual(14.09D, eppeo.LookupCellValue(Of Double)(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual(14.09D, mseo.LookupCellValue(Of Double)(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual("Chef: 14,09", eppeo.LookupCellFormattedText(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual("Chef: 14,09", mseo.LookupCellFormattedText(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual(Nothing, eppeo.LookupCellFormula(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))
                Assert.AreEqual(Nothing, mseo.LookupCellFormula(New ExcelOps.ExcelCell(TestSheet, "A8", ExcelOps.ExcelCell.ValueTypes.All)))

                'E1
                Assert.AreEqual(False, eppeo.LookupCellValue(Of Boolean)(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                Assert.AreEqual(False, mseo.LookupCellValue(Of Boolean)(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                Assert.AreEqual("False", eppeo.LookupCellValue(Of String)(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                Assert.AreEqual("False", mseo.LookupCellValue(Of String)(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                Assert.AreEqual("False", eppeo.LookupCellFormattedText(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
                Assert.AreEqual("False", mseo.LookupCellFormattedText(New ExcelCell(TestSheet, "E1", ExcelCell.ValueTypes.All)))
            Finally
                mseo.CloseExcelAppInstance()
            End Try
        End Sub

        <Test> Public Sub LookupLastCellAddress()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestEnvironment.TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, True)
            Try
                Dim LastCellFound As ExcelOps.ExcelCell
                LastCellFound = eppeo.LookupLastContentCell(TestSheet)
                Assert.AreEqual("E40", LastCellFound.Address) 'in file: D40, but E40 after CT-Load with setup of field E1 for tracking required calculations
                Assert.AreEqual(eppeo.LookupLastContentRowIndex(TestSheet), LastCellFound.RowIndex)
                Assert.AreEqual(eppeo.LookupLastContentColumnIndex(TestSheet), LastCellFound.ColumnIndex)
                LastCellFound = mseo.LookupLastContentCell(TestSheet)
                Assert.AreEqual("E40", LastCellFound.Address) 'in file: D40, but E40 after CT-Load with setup of field E1 for tracking required calculations
                Assert.AreEqual(mseo.LookupLastContentRowIndex(TestSheet), LastCellFound.RowIndex)
                Assert.AreEqual(mseo.LookupLastContentColumnIndex(TestSheet), LastCellFound.ColumnIndex)
            Finally
                mseo.CloseExcelAppInstance()
            End Try
        End Sub

        <Test> Public Sub AddSheet()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestEnvironment.TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            Dim BeforeSheet As String = "Grunddaten"
            Dim SheetNameTopPosition As String = "SheetOnTop"
            Dim SheetNameBottomPosition As String = "SheetOnBottom"
            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, True)
            Try
                Dim ExpectedSheetNamesList, NewSheetNamesList As List(Of String)
                ExpectedSheetNamesList = eppeo.SheetNames
                ExpectedSheetNamesList.Add(SheetNameBottomPosition)
                ExpectedSheetNamesList.Insert(0, SheetNameTopPosition)

                eppeo.AddSheet(SheetNameBottomPosition)
                eppeo.AddSheet(SheetNameTopPosition, BeforeSheet)
                NewSheetNamesList = eppeo.SheetNames
                Assert.AreEqual(ExpectedSheetNamesList.ToArray, NewSheetNamesList.ToArray)

                ExpectedSheetNamesList = mseo.SheetNames
                ExpectedSheetNamesList.Add(SheetNameBottomPosition)
                ExpectedSheetNamesList.Insert(0, SheetNameTopPosition)
                System.Console.WriteLine("MS Expected: " & Strings.Join(ExpectedSheetNamesList.ToArray, ","))

                mseo.AddSheet(SheetNameBottomPosition)
                mseo.AddSheet(SheetNameTopPosition, BeforeSheet)
                NewSheetNamesList = mseo.SheetNames
                System.Console.WriteLine("MS NewList : " & Strings.Join(NewSheetNamesList.ToArray, ","))
                Assert.AreEqual(ExpectedSheetNamesList.ToArray, NewSheetNamesList.ToArray)
            Finally
                mseo.CloseExcelAppInstance()
            End Try
        End Sub

        <Test> Public Sub SheetNames()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestEnvironment.TestFiles.TestFileGrund01.FullName
            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, True)
            Try
                Dim EppeoSheetNamesList, MseoSheetNamesList As List(Of String)
                EppeoSheetNamesList = eppeo.SheetNames
                MseoSheetNamesList = mseo.SheetNames
                System.Console.WriteLine("EPP: " & Strings.Join(EppeoSheetNamesList.ToArray, ","))
                System.Console.WriteLine("MS : " & Strings.Join(MseoSheetNamesList.ToArray, ","))
                Assert.AreEqual(EppeoSheetNamesList.ToArray, MseoSheetNamesList.ToArray)
            Finally
                mseo.CloseExcelAppInstance()
            End Try
        End Sub

        <Test> Public Sub CalcTest()
            Dim wb As New OfficeOpenXml.ExcelPackage()
            Dim TestCell As OfficeOpenXml.ExcelRange
            wb.Workbook.Worksheets.Add("test-calcs")
            TestCell = wb.Workbook.Worksheets(1).Cells(1, 1)
            Assert.AreEqual("#NAME?", Me.CalcTestCell(TestCell, "INVALIDFUNCTION(B2)"))
            Assert.AreEqual("5", Me.CalcTestCell(TestCell, "2+3"))
            Assert.AreEqual("6", Me.CalcTestCell(TestCell, "2*3"))
            Assert.AreEqual("0", Me.CalcTestCell(TestCell, "B2"))
            If "#VALUE!" = Me.CalcTestCell(TestCell, "B2+B3") Then
                Assert.Warn("EPPlus calculation engine not working for formula '=B2+B3'")
                Assert.Ignore("EPPlus calculation engine not working for formula '=B2+B3'")
            End If
            Assert.AreEqual("0", Me.CalcTestCell(TestCell, "B2+B3"))
            Assert.AreEqual("5", Me.CalcTestCell(TestCell, "SUM(2,3)"))
        End Sub

        Private Function CalcTestCell(cell As OfficeOpenXml.ExcelRange, formula As String) As String
            cell.Formula = formula
            OfficeOpenXml.CalculationExtension.Calculate(cell)
            Try
                If cell Is Nothing Then
                    Return Nothing
                ElseIf cell.Value.GetType Is GetType(OfficeOpenXml.ExcelErrorValue) Then
                    Return CType(cell.Value, OfficeOpenXml.ExcelErrorValue).ToString
                Else
                    Return CType(cell.Value, String)
                End If
#Disable Warning CA1031 ' Do not catch general exception types
            Catch ex As Exception
                Return "ERROR: " & ex.Message
#Enable Warning CA1031 ' Do not catch general exception types
            End Try
        End Function


        <Test> Public Sub CellWithErrorMsExcel()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim wb As New MsExcelDataOperations(TestEnvironment.TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, True)
            Dim SheetName As String = wb.SheetNames(0)

            Try
                wb.WriteCellFormula(SheetName, 0, 0, "B2", True)
                Assert.AreEqual(Nothing, wb.LookupCellErrorValue(SheetName, 0, 0))

                wb.WriteCellFormula(SheetName, 0, 0, "1/1", True)
                Assert.AreEqual(Nothing, wb.LookupCellErrorValue(SheetName, 0, 0))
                Assert.AreEqual(1, wb.LookupCellValue(Of Integer)(SheetName, 0, 0))

                wb.WriteCellFormula(SheetName, 0, 0, "INVALIDFUNCTION(B2)", True)
                Assert.AreEqual("#NAME?", wb.LookupCellErrorValue(SheetName, 0, 0))

                wb.WriteCellFormula(SheetName, 0, 0, "1/0", True)
                Assert.AreEqual("#DIV/0!", wb.LookupCellErrorValue(SheetName, 0, 0))

                wb.WriteCellFormula(SheetName, 0, 0, "B2/0", True)
                Assert.AreEqual("#DIV/0!", wb.LookupCellErrorValue(SheetName, 0, 0))

                wb.WriteCellFormula(SheetName, 0, 0, "A0", True)
                Assert.AreEqual("#NAME?", wb.LookupCellErrorValue(SheetName, 0, 0))
            Finally
                wb.CloseExcelAppInstance()
            End Try
        End Sub

        <Test> Public Sub CellWithErrorEpplus()
            Dim wb As New EpplusFreeExcelDataOperations(TestEnvironment.TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
            Dim SheetName As String = wb.SheetNames(0)

            wb.WriteCellFormula(SheetName, 0, 0, "B2", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual(Nothing, wb.LookupCellErrorValue(SheetName, 0, 0))

            wb.WriteCellFormula(SheetName, 0, 0, "1/1", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual(Nothing, wb.LookupCellErrorValue(SheetName, 0, 0))
            Assert.AreEqual(1, wb.LookupCellValue(Of Integer)(SheetName, 0, 0))

            wb.WriteCellFormula(SheetName, 0, 0, "INVALIDFUNCTION(B2)", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual("#NAME?", wb.LookupCellErrorValue(SheetName, 0, 0))

            wb.WriteCellFormula(SheetName, 0, 0, "1/0", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual("#DIV/0!", wb.LookupCellErrorValue(SheetName, 0, 0))

            wb.WriteCellFormula(SheetName, 0, 0, "B2/0", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual("#DIV/0!", wb.LookupCellErrorValue(SheetName, 0, 0))

            Assert.Ignore("Bugs in Epplus formula manager engine")

            wb.WriteCellFormula(SheetName, 0, 0, "A0", False)
            wb.RecalculateCell(SheetName, 0, 0, False)
            Assert.AreEqual("#NAME?", wb.LookupCellErrorValue(SheetName, 0, 0)) 'bug in Epplus engine
        End Sub

        <Test> Public Sub AllFormulasOfWorkbook()
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String
            Dim AllFormulas As List(Of ExcelOps.TextTableCell)

            TestControllingToolFileName = TestEnvironment.TestFiles.TestFileGrund01.FullName
            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
            AllFormulas = eppeo.AllFormulasOfWorkbook
            Console.WriteLine("Test file: " & TestControllingToolFileName)
            Assert.NotZero(AllFormulas.Count)

            Dim ReferencesFromTestSheet As List(Of TextTableCell) = ExcelOps.Tools.FormulasWithSheetReferencesFromSheet("Kostenplanung", AllFormulas, eppeo.SheetNames.ToArray)
            System.Console.WriteLine("## Formulas of Kostenplanung targetting other sheets in workbook")
            For MyFormulaCounter As Integer = 0 To ReferencesFromTestSheet.Count - 1
                System.Console.WriteLine(ReferencesFromTestSheet(MyFormulaCounter).ToString)
            Next
            Assert.NotZero(ReferencesFromTestSheet.Count)
            Assert.IsTrue(ExcelOps.Tools.ContainsFormulasWithSheetReferencesFromSheet("Kostenplanung", AllFormulas, eppeo.SheetNames.ToArray))
            Assert.IsFalse(ExcelOps.Tools.ContainsFormulasWithSheetReferencesFromSheet("Grunddaten", AllFormulas, eppeo.SheetNames.ToArray))

            System.Console.WriteLine()
            System.Console.WriteLine("## Formulas of sheets in workbook targetting Grunddaten")
            Dim ReferencesToTestSheet As List(Of TextTableCell) = ExcelOps.Tools.FormulasWithSheetReferencesToSheet(AllFormulas, "Grunddaten", Nothing)
            For MyFormulaCounter As Integer = 0 To ReferencesToTestSheet.Count - 1
                System.Console.WriteLine(ReferencesToTestSheet(MyFormulaCounter).ToString)
            Next
            Assert.NotZero(ReferencesToTestSheet.Count)
            Assert.IsTrue(ExcelOps.Tools.ContainsFormulasWithSheetReferencesToSheet(AllFormulas, "Grunddaten", Nothing))
        End Sub

        <Test> Public Sub CopySheetContentEpplus()
            Dim eppeoIn As ExcelOps.ExcelDataOperationsBase
            Dim eppeoOut As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileNameIn As String
            Dim TestControllingToolFileNameOutTemplate As String
            Dim TestControllingToolFileNameOut As String

            TestControllingToolFileNameIn = TestEnvironment.TestFiles.TestFileGrund01.FullName
            TestControllingToolFileNameOutTemplate = CTTestFiles.TestFileV25.FullName
            TestControllingToolFileNameOut = TestEnvironment.FullPathOfDynTestFile("CopySheetContentEpplus.xlsx")
            eppeoIn = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileNameIn, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
            eppeoOut = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileNameOutTemplate, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)

            Console.WriteLine("Test file in: " & TestControllingToolFileNameIn)
            Console.WriteLine("Test file output template: " & TestControllingToolFileNameOutTemplate)
            Console.WriteLine("Test file output: " & TestControllingToolFileNameOut)

            Try
                eppeoIn.CopySheetContent("Unternehmerlohn", eppeoOut, CopySheetOption.TargetSheetMightExist)
                eppeoOut.SelectSheet("Unternehmerlohn")
                eppeoOut.SaveAs(TestControllingToolFileNameOut, False, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.AreEqual(eppeoIn.SheetContentMatrix("Unternehmerlohn", ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText), eppeoOut.SheetContentMatrix("Unternehmerlohn", ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText))
                Assert.Pass("Required manual, optical review for comparison to check for formattings")
            Catch ex As NotSupportedException
                Assert.Ignore("Not supported by Epplus engine, currently: copy sheet content incl. data+formats+locks")
            End Try
        End Sub

        <Test> Public Sub CopySheetContentMsExcel()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeoIn As ExcelOps.ExcelDataOperationsBase = Nothing
            Dim eppeoOut As ExcelOps.ExcelDataOperationsBase = Nothing
            Dim TestControllingToolFileNameIn As String
            Dim TestControllingToolFileNameOutTemplate As String
            Dim TestControllingToolFileNameOut As String
            Dim MsExcelApp As New MsExcelDataOperations.MsAppInstance()

            TestControllingToolFileNameIn = TestEnvironment.TestFiles.TestFileGrund01.FullName
            TestControllingToolFileNameOutTemplate = CTTestFiles.TestFileV25.FullName
            TestControllingToolFileNameOut = TestEnvironment.FullPathOfDynTestFile("CopySheetContentMsExcel.xlsx")
            Try
                eppeoIn = New ExcelOps.MsExcelDataOperations(TestControllingToolFileNameIn, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelApp, True)
                eppeoOut = New ExcelOps.MsExcelDataOperations(TestControllingToolFileNameOutTemplate, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelApp, True)

                Console.WriteLine("Test file in: " & TestControllingToolFileNameIn)
                Console.WriteLine("Test file output template: " & TestControllingToolFileNameOutTemplate)
                Console.WriteLine("Test file output: " & TestControllingToolFileNameOut)

                eppeoIn.CopySheetContent("Unternehmerlohn", eppeoOut, CopySheetOption.TargetSheetMightExist)
                eppeoOut.SelectSheet("Unternehmerlohn")
                eppeoOut.SaveAs(TestControllingToolFileNameOut, False, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.AreEqual(eppeoIn.SheetContentMatrix("Unternehmerlohn", ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText), eppeoOut.SheetContentMatrix("Unternehmerlohn", ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText))
                Assert.Pass("Required manual, optical review for comparison to check for formattings")
            Finally
                If eppeoOut IsNot Nothing Then eppeoOut.Close()
                eppeoIn.CloseExcelAppInstance()
            End Try
        End Sub

        <Test> Public Sub ExcelOpsTestCollection_ZahlenUndProzentwerte()
            Dim eppeo As New ExcelOps.EpplusFreeExcelDataOperations(TestEnvironment.TestFiles.TestFileExcelOpsTestCollection.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True)
#If Not CI_CD Then
            Dim mseo As New MsExcelDataOperations(TestEnvironment.TestFiles.TestFileExcelOpsTestCollection.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, False, True)
#End If
            Try
                Dim SheetName As String
                SheetName = "ZahlenUndProzentwerte"
                Assert.AreEqual("0.00", eppeo.LookupCellFormat(SheetName, 0, 1))
                Assert.AreEqual("0.00%", eppeo.LookupCellFormat(SheetName, 1, 1))
                Assert.AreEqual(10.0, System.Math.Round(eppeo.LookupCellValue(Of Double)(SheetName, 0, 1), 2))
                Assert.AreEqual(0.1, System.Math.Round(eppeo.LookupCellValue(Of Double)(SheetName, 1, 1), 2))
                eppeo.WriteCellValue(Of Double)(SheetName, 0, 1, 20.0)
                eppeo.WriteCellValue(Of Double)(SheetName, 1, 1, 0.2)
                Assert.AreEqual(20.0, System.Math.Round(eppeo.LookupCellValue(Of Double)(SheetName, 0, 1), 2))
                Assert.AreEqual(0.2, System.Math.Round(eppeo.LookupCellValue(Of Double)(SheetName, 1, 1), 2))
#If Not CI_CD Then
                Assert.AreEqual("0.00", mseo.LookupCellFormat(SheetName, 0, 1))
                Assert.AreEqual("0.00%", mseo.LookupCellFormat(SheetName, 1, 1))
                Assert.AreEqual(10.0, System.Math.Round(mseo.LookupCellValue(Of Double)(SheetName, 0, 1), 2))
                Assert.AreEqual(0.1, System.Math.Round(mseo.LookupCellValue(Of Double)(SheetName, 1, 1), 2))
                mseo.WriteCellValue(Of Double)(SheetName, 0, 1, 20.0)
                mseo.WriteCellValue(Of Double)(SheetName, 1, 1, 0.2)
                Assert.AreEqual(20.0, System.Math.Round(mseo.LookupCellValue(Of Double)(SheetName, 0, 1), 2))
                Assert.AreEqual(0.2, System.Math.Round(mseo.LookupCellValue(Of Double)(SheetName, 1, 1), 2))
#End If
            Finally
#If Not CI_CD Then
                If mseo IsNot Nothing Then
                    mseo.CloseExcelAppInstance()
                End If
#End If
            End Try
        End Sub

    End Class
End Namespace