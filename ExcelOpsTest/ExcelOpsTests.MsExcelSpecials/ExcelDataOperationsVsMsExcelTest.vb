Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps

Namespace ExcelOpsTests.MsExcelSpecials

    <TestFixture> Public Class ExcelDataOperationsTest

        <OneTimeSetUp()> Public Sub OneTimeInitConfig()
        End Sub

        <SetUp> Public Sub Setup()
            Test.Console.ResetConsoleForTestOutput()
            If _MsExcelAppWrapper IsNot Nothing Then
                _MsExcelAppWrapper.Workbooks.CloseAllWorkbooks()
            End If
        End Sub

        Private _MsExcelAppWrapper As MsExcelCom.MsExcelApplicationWrapper
        Private ReadOnly Property MsExcelAppWrapper As MsExcelCom.MsExcelApplicationWrapper
            Get
                If _MsExcelAppWrapper Is Nothing Then
                    Try
                        _MsExcelAppWrapper = New MsExcelCom.MsExcelApplicationWrapper
                    Catch ex As System.PlatformNotSupportedException
                        Assert.Ignore("Platform not supported or MS Excel app not installed: " & ex.Message)
                    End Try
                End If
                Return _MsExcelAppWrapper
            End Get
        End Property

        <TearDown>
        Public Sub TearDown()
            CompuMaster.ComInterop.ComTools.GarbageCollectAndWaitForPendingFinalizers()
        End Sub

        <OneTimeTearDown>
        Public Sub OneTimeTearDown()
            If _MsExcelAppWrapper IsNot Nothing Then
                _MsExcelAppWrapper.Dispose()
            End If
        End Sub

        <Test> Public Sub SheetContentMatrix()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase

            Dim ExpectedMatrix As String
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, True, True, String.Empty)

            ExpectedMatrix =
                         "# |A                           |B        |C  |D     |E    " & ControlChars.CrLf &
                         "--+----------------------------+---------+---+------+-----" & ControlChars.CrLf &
                         "1 |Jahr                        |2019     |   |      |False" & ControlChars.CrLf &
                         "2 |Geschäftsjahr von           |         |bis|      |     " & ControlChars.CrLf &
                         "3 |Aktueller Monat             |1        |   |Januar|     " & ControlChars.CrLf &
                         "4 |                            |         |   |      |     " & ControlChars.CrLf &
                         "5 |Name Betrieb                |Test     |   |      |     " & ControlChars.CrLf &
                         "6 |                            |         |   |      |     " & ControlChars.CrLf &
                         "7 |Arbeitgeberanteile in %     |         |   |      |     " & ControlChars.CrLf &
                         "8 |Chef: 14,09                 |         |   |      |     " & ControlChars.CrLf &
                         "9 |Büroangestellte: 20,00      |         |   |      |     " & ControlChars.CrLf &
                         "10|Produktivkraft: 25,00       |         |   |      |     " & ControlChars.CrLf &
                         "11|Azubi / Aushilfen: 33,00    |         |   |      |     " & ControlChars.CrLf &
                         "12|                            |         |   |      |     " & ControlChars.CrLf &
                         "13|Berechnung Jahresarbeitszeit|         |   |      |     " & ControlChars.CrLf &
                         "14|Tage / Jahr:                |365      |   |      |     " & ControlChars.CrLf &
                         "15|Wochenendtage               |104      |   |      |     " & ControlChars.CrLf &
                         "16|=Zahltage:                  |261      |   |      |     " & ControlChars.CrLf &
                         "17|Wochenarbeitszeit           |40       |   |      |     " & ControlChars.CrLf &
                         "18|Tagesarbeitszeit:           |8        |   |      |     " & ControlChars.CrLf &
                         "19|Normallohnstunden / Jahr:   |2.088,00 |   |      |     " & ControlChars.CrLf &
                         "20|                            |         |   |      |     " & ControlChars.CrLf &
                         "21|                            |         |   |      |     " & ControlChars.CrLf &
                         "22|                            |         |   |      |     " & ControlChars.CrLf &
                         "23|1                           |Januar   |   |      |     " & ControlChars.CrLf &
                         "24|2                           |Februar  |   |      |     " & ControlChars.CrLf &
                         "25|3                           |März     |   |      |     " & ControlChars.CrLf &
                         "26|4                           |April    |   |      |     " & ControlChars.CrLf &
                         "27|5                           |Mai      |   |      |     " & ControlChars.CrLf &
                         "28|6                           |Juni     |   |      |     " & ControlChars.CrLf &
                         "29|7                           |Juli     |   |      |     " & ControlChars.CrLf &
                         "30|8                           |August   |   |      |     " & ControlChars.CrLf &
                         "31|9                           |September|   |      |     " & ControlChars.CrLf &
                         "32|10                          |Oktober  |   |      |     " & ControlChars.CrLf &
                         "33|11                          |November |   |      |     " & ControlChars.CrLf &
                         "34|12                          |Dezember |   |      |     " & ControlChars.CrLf &
                         "35|Zusammensetzung AG Anteile  |         |   |      |     " & ControlChars.CrLf &
                         "36|Krankenkasse                |2,8      |   |      |     " & ControlChars.CrLf &
                         "37|Rentenkasse                 |8        |   |      |     " & ControlChars.CrLf &
                         "38|Pflegekasse                 |1,4      |   |      |     " & ControlChars.CrLf &
                         "39|Krankengeld                 |0,25     |   |      |     " & ControlChars.CrLf &
                         "40|                            |12,45    |   |      |     " & ControlChars.CrLf
            Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues, ExpectedMatrix)
            Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticOrCalculatedValues, ExpectedMatrix)

            ExpectedMatrix =
                         "# |A                           |B        |C  |D |E    " & ControlChars.CrLf &
                         "--+----------------------------+---------+---+--+-----" & ControlChars.CrLf &
                         "1 |Jahr                        |2019     |   |  |False" & ControlChars.CrLf &
                         "2 |Geschäftsjahr von           |         |bis|  |     " & ControlChars.CrLf &
                         "3 |Aktueller Monat             |1        |   |  |     " & ControlChars.CrLf &
                         "4 |                            |         |   |  |     " & ControlChars.CrLf &
                         "5 |Name Betrieb                |Test     |   |  |     " & ControlChars.CrLf &
                         "6 |                            |         |   |  |     " & ControlChars.CrLf &
                         "7 |Arbeitgeberanteile in %     |         |   |  |     " & ControlChars.CrLf &
                         "8 |Chef: 14,09                 |         |   |  |     " & ControlChars.CrLf &
                         "9 |Büroangestellte: 20,00      |         |   |  |     " & ControlChars.CrLf &
                         "10|Produktivkraft: 25,00       |         |   |  |     " & ControlChars.CrLf &
                         "11|Azubi / Aushilfen: 33,00    |         |   |  |     " & ControlChars.CrLf &
                         "12|                            |         |   |  |     " & ControlChars.CrLf &
                         "13|Berechnung Jahresarbeitszeit|         |   |  |     " & ControlChars.CrLf &
                         "14|Tage / Jahr:                |365      |   |  |     " & ControlChars.CrLf &
                         "15|Wochenendtage               |         |   |  |     " & ControlChars.CrLf &
                         "16|=Zahltage:                  |         |   |  |     " & ControlChars.CrLf &
                         "17|Wochenarbeitszeit           |40       |   |  |     " & ControlChars.CrLf &
                         "18|Tagesarbeitszeit:           |         |   |  |     " & ControlChars.CrLf &
                         "19|Normallohnstunden / Jahr:   |         |   |  |     " & ControlChars.CrLf &
                         "20|                            |         |   |  |     " & ControlChars.CrLf &
                         "21|                            |         |   |  |     " & ControlChars.CrLf &
                         "22|                            |         |   |  |     " & ControlChars.CrLf &
                         "23|1                           |Januar   |   |  |     " & ControlChars.CrLf &
                         "24|2                           |Februar  |   |  |     " & ControlChars.CrLf &
                         "25|3                           |März     |   |  |     " & ControlChars.CrLf &
                         "26|4                           |April    |   |  |     " & ControlChars.CrLf &
                         "27|5                           |Mai      |   |  |     " & ControlChars.CrLf &
                         "28|6                           |Juni     |   |  |     " & ControlChars.CrLf &
                         "29|7                           |Juli     |   |  |     " & ControlChars.CrLf &
                         "30|8                           |August   |   |  |     " & ControlChars.CrLf &
                         "31|9                           |September|   |  |     " & ControlChars.CrLf &
                         "32|10                          |Oktober  |   |  |     " & ControlChars.CrLf &
                         "33|11                          |November |   |  |     " & ControlChars.CrLf &
                         "34|12                          |Dezember |   |  |     " & ControlChars.CrLf &
                         "35|Zusammensetzung AG Anteile  |         |   |  |     " & ControlChars.CrLf &
                         "36|Krankenkasse                |2,8      |   |  |     " & ControlChars.CrLf &
                         "37|Rentenkasse                 |8        |   |  |     " & ControlChars.CrLf &
                         "38|Pflegekasse                 |1,4      |   |  |     " & ControlChars.CrLf &
                         "39|Krankengeld                 |0,25     |   |  |     " & ControlChars.CrLf
            Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticValues, ExpectedMatrix)
            Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.StaticValues, ExpectedMatrix)

            ExpectedMatrix =
                         "# |A |B            |C |D                                  " & ControlChars.CrLf &
                         "--+--+-------------+--+-----------------------------------" & ControlChars.CrLf &
                         "1 |  |             |  |                                   " & ControlChars.CrLf &
                         "2 |  |             |  |                                   " & ControlChars.CrLf &
                         "3 |  |             |  |=INDEX(B23:B34,MATCH(B3,A23:A34,0))" & ControlChars.CrLf &
                         "4 |  |             |  |                                   " & ControlChars.CrLf &
                         "5 |  |             |  |                                   " & ControlChars.CrLf &
                         "6 |  |             |  |                                   " & ControlChars.CrLf &
                         "7 |  |             |  |                                   " & ControlChars.CrLf &
                         "8 |  |             |  |                                   " & ControlChars.CrLf &
                         "9 |  |             |  |                                   " & ControlChars.CrLf &
                         "10|  |             |  |                                   " & ControlChars.CrLf &
                         "11|  |             |  |                                   " & ControlChars.CrLf &
                         "12|  |             |  |                                   " & ControlChars.CrLf &
                         "13|  |             |  |                                   " & ControlChars.CrLf &
                         "14|  |             |  |                                   " & ControlChars.CrLf &
                         "15|  |=2*52        |  |                                   " & ControlChars.CrLf &
                         "16|  |=B14-B15     |  |                                   " & ControlChars.CrLf &
                         "17|  |             |  |                                   " & ControlChars.CrLf &
                         "18|  |=B17/5       |  |                                   " & ControlChars.CrLf &
                         "19|  |=B18*B16     |  |                                   " & ControlChars.CrLf &
                         "20|  |             |  |                                   " & ControlChars.CrLf &
                         "21|  |             |  |                                   " & ControlChars.CrLf &
                         "22|  |             |  |                                   " & ControlChars.CrLf &
                         "23|  |             |  |                                   " & ControlChars.CrLf &
                         "24|  |             |  |                                   " & ControlChars.CrLf &
                         "25|  |             |  |                                   " & ControlChars.CrLf &
                         "26|  |             |  |                                   " & ControlChars.CrLf &
                         "27|  |             |  |                                   " & ControlChars.CrLf &
                         "28|  |             |  |                                   " & ControlChars.CrLf &
                         "29|  |             |  |                                   " & ControlChars.CrLf &
                         "30|  |             |  |                                   " & ControlChars.CrLf &
                         "31|  |             |  |                                   " & ControlChars.CrLf &
                         "32|  |             |  |                                   " & ControlChars.CrLf &
                         "33|  |             |  |                                   " & ControlChars.CrLf &
                         "34|  |             |  |                                   " & ControlChars.CrLf &
                         "35|  |             |  |                                   " & ControlChars.CrLf &
                         "36|  |             |  |                                   " & ControlChars.CrLf &
                         "37|  |             |  |                                   " & ControlChars.CrLf &
                         "38|  |             |  |                                   " & ControlChars.CrLf &
                         "39|  |             |  |                                   " & ControlChars.CrLf &
                         "40|  |=SUM(B36:B39)|  |                                   " & ControlChars.CrLf
            Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.Formulas, ExpectedMatrix)
            Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.Formulas, ExpectedMatrix)

            ExpectedMatrix =
                         "# |A                           |B        |C  |D     |E    " & ControlChars.CrLf &
                         "--+----------------------------+---------+---+------+-----" & ControlChars.CrLf &
                         "1 |Jahr                        |2019     |   |      |False" & ControlChars.CrLf &
                         "2 |Geschäftsjahr von           |         |bis|      |     " & ControlChars.CrLf &
                         "3 |Aktueller Monat             |1        |   |Januar|     " & ControlChars.CrLf &
                         "4 |                            |         |   |      |     " & ControlChars.CrLf &
                         "5 |Name Betrieb                |Test     |   |      |     " & ControlChars.CrLf &
                         "6 |                            |         |   |      |     " & ControlChars.CrLf &
                         "7 |Arbeitgeberanteile in %     |         |   |      |     " & ControlChars.CrLf &
                         "8 |Chef: 14,09                 |         |   |      |     " & ControlChars.CrLf &
                         "9 |Büroangestellte: 20,00      |         |   |      |     " & ControlChars.CrLf &
                         "10|Produktivkraft: 25,00       |         |   |      |     " & ControlChars.CrLf &
                         "11|Azubi / Aushilfen: 33,00    |         |   |      |     " & ControlChars.CrLf &
                         "12|                            |         |   |      |     " & ControlChars.CrLf &
                         "13|Berechnung Jahresarbeitszeit|         |   |      |     " & ControlChars.CrLf &
                         "14|Tage / Jahr:                |365      |   |      |     " & ControlChars.CrLf &
                         "15|Wochenendtage               |104      |   |      |     " & ControlChars.CrLf &
                         "16|=Zahltage:                  |261      |   |      |     " & ControlChars.CrLf &
                         "17|Wochenarbeitszeit           |40       |   |      |     " & ControlChars.CrLf &
                         "18|Tagesarbeitszeit:           |8        |   |      |     " & ControlChars.CrLf &
                         "19|Normallohnstunden / Jahr:   |2.088,00 |   |      |     " & ControlChars.CrLf &
                         "20|                            |         |   |      |     " & ControlChars.CrLf &
                         "21|                            |         |   |      |     " & ControlChars.CrLf &
                         "22|                            |         |   |      |     " & ControlChars.CrLf &
                         "23|1                           |Januar   |   |      |     " & ControlChars.CrLf &
                         "24|2                           |Februar  |   |      |     " & ControlChars.CrLf &
                         "25|3                           |März     |   |      |     " & ControlChars.CrLf &
                         "26|4                           |April    |   |      |     " & ControlChars.CrLf &
                         "27|5                           |Mai      |   |      |     " & ControlChars.CrLf &
                         "28|6                           |Juni     |   |      |     " & ControlChars.CrLf &
                         "29|7                           |Juli     |   |      |     " & ControlChars.CrLf &
                         "30|8                           |August   |   |      |     " & ControlChars.CrLf &
                         "31|9                           |September|   |      |     " & ControlChars.CrLf &
                         "32|10                          |Oktober  |   |      |     " & ControlChars.CrLf &
                         "33|11                          |November |   |      |     " & ControlChars.CrLf &
                         "34|12                          |Dezember |   |      |     " & ControlChars.CrLf &
                         "35|Zusammensetzung AG Anteile  |         |   |      |     " & ControlChars.CrLf &
                         "36|Krankenkasse                |2,8      |   |      |     " & ControlChars.CrLf &
                         "37|Rentenkasse                 |8        |   |      |     " & ControlChars.CrLf &
                         "38|Pflegekasse                 |1,4      |   |      |     " & ControlChars.CrLf &
                         "39|Krankengeld                 |0,25     |   |      |     " & ControlChars.CrLf &
                         "40|                            |12,45    |   |      |     " & ControlChars.CrLf
            Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormattedText, ExpectedMatrix)
            Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormattedText, ExpectedMatrix)

            ExpectedMatrix =
                         "# |A                           |B            |C  |D                                  |E    " & ControlChars.CrLf &
                         "--+----------------------------+-------------+---+-----------------------------------+-----" & ControlChars.CrLf &
                         "1 |Jahr                        |2019         |   |                                   |False" & ControlChars.CrLf &
                         "2 |Geschäftsjahr von           |             |bis|                                   |     " & ControlChars.CrLf &
                         "3 |Aktueller Monat             |1            |   |=INDEX(B23:B34,MATCH(B3,A23:A34,0))|     " & ControlChars.CrLf &
                         "4 |                            |             |   |                                   |     " & ControlChars.CrLf &
                         "5 |Name Betrieb                |Test         |   |                                   |     " & ControlChars.CrLf &
                         "6 |                            |             |   |                                   |     " & ControlChars.CrLf &
                         "7 |Arbeitgeberanteile in %     |             |   |                                   |     " & ControlChars.CrLf &
                         "8 |Chef: 14,09                 |             |   |                                   |     " & ControlChars.CrLf &
                         "9 |Büroangestellte: 20,00      |             |   |                                   |     " & ControlChars.CrLf &
                         "10|Produktivkraft: 25,00       |             |   |                                   |     " & ControlChars.CrLf &
                         "11|Azubi / Aushilfen: 33,00    |             |   |                                   |     " & ControlChars.CrLf &
                         "12|                            |             |   |                                   |     " & ControlChars.CrLf &
                         "13|Berechnung Jahresarbeitszeit|             |   |                                   |     " & ControlChars.CrLf &
                         "14|Tage / Jahr:                |365          |   |                                   |     " & ControlChars.CrLf &
                         "15|Wochenendtage               |=2*52        |   |                                   |     " & ControlChars.CrLf &
                         "16|=Zahltage:                  |=B14-B15     |   |                                   |     " & ControlChars.CrLf &
                         "17|Wochenarbeitszeit           |40           |   |                                   |     " & ControlChars.CrLf &
                         "18|Tagesarbeitszeit:           |=B17/5       |   |                                   |     " & ControlChars.CrLf &
                         "19|Normallohnstunden / Jahr:   |=B18*B16     |   |                                   |     " & ControlChars.CrLf &
                         "20|                            |             |   |                                   |     " & ControlChars.CrLf &
                         "21|                            |             |   |                                   |     " & ControlChars.CrLf &
                         "22|                            |             |   |                                   |     " & ControlChars.CrLf &
                         "23|1                           |Januar       |   |                                   |     " & ControlChars.CrLf &
                         "24|2                           |Februar      |   |                                   |     " & ControlChars.CrLf &
                         "25|3                           |März         |   |                                   |     " & ControlChars.CrLf &
                         "26|4                           |April        |   |                                   |     " & ControlChars.CrLf &
                         "27|5                           |Mai          |   |                                   |     " & ControlChars.CrLf &
                         "28|6                           |Juni         |   |                                   |     " & ControlChars.CrLf &
                         "29|7                           |Juli         |   |                                   |     " & ControlChars.CrLf &
                         "30|8                           |August       |   |                                   |     " & ControlChars.CrLf &
                         "31|9                           |September    |   |                                   |     " & ControlChars.CrLf &
                         "32|10                          |Oktober      |   |                                   |     " & ControlChars.CrLf &
                         "33|11                          |November     |   |                                   |     " & ControlChars.CrLf &
                         "34|12                          |Dezember     |   |                                   |     " & ControlChars.CrLf &
                         "35|Zusammensetzung AG Anteile  |             |   |                                   |     " & ControlChars.CrLf &
                         "36|Krankenkasse                |2,8          |   |                                   |     " & ControlChars.CrLf &
                         "37|Rentenkasse                 |8            |   |                                   |     " & ControlChars.CrLf &
                         "38|Pflegekasse                 |1,4          |   |                                   |     " & ControlChars.CrLf &
                         "39|Krankengeld                 |0,25         |   |                                   |     " & ControlChars.CrLf &
                         "40|                            |=SUM(B36:B39)|   |                                   |     " & ControlChars.CrLf
            Me.SheetContentMatrix(eppeo, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText, ExpectedMatrix)
            Me.SheetContentMatrix(mseo, ExcelOps.ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText, ExpectedMatrix)
        End Sub

        Private Sub SheetContentMatrix(eo As ExcelOps.ExcelDataOperationsBase, matrixContentType As ExcelOps.ExcelDataOperationsBase.MatrixContent, expectedMatrix As String)
            Dim MatrixContentName As String = matrixContentType.ToString
            Dim Grunddaten As TextTable = eo.SheetContentMatrix("Grunddaten", matrixContentType)
            Grunddaten.AutoTrim()
            Select Case eo.GetType
                Case GetType(ExcelOps.EpplusFreeExcelDataOperations), GetType(ExcelOps.EpplusPolyformExcelDataOperations)
                    Console.WriteLine("## Table EPPlus - " & MatrixContentName)
                Case GetType(MsExcelDataOperations)
#If CI_CD Then
                    If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
                    Console.WriteLine("## Table MS Excel - " & MatrixContentName)
                Case Else
                    Throw New NotImplementedException
            End Select
            Console.WriteLine(Grunddaten.ToUIExcelTable)
            Console.WriteLine("## /Table")
            Assert.AreEqual(expectedMatrix, Grunddaten.ToUIExcelTable)
        End Sub

        <Test> Public Sub LookupCellValue()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, True, True, String.Empty)

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
        End Sub

        <Test> Public Sub LookupLastCellAddress()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, True, True, String.Empty)

            Dim LastCellFound As ExcelOps.ExcelCell
            LastCellFound = eppeo.LookupLastContentCell(TestSheet)
            Assert.AreEqual("E40", LastCellFound.Address)
            Assert.AreEqual(eppeo.LookupLastContentRowIndex(TestSheet), LastCellFound.RowIndex)
            Assert.AreEqual(eppeo.LookupLastContentColumnIndex(TestSheet), LastCellFound.ColumnIndex)
            LastCellFound = mseo.LookupLastContentCell(TestSheet)
            Assert.AreEqual("E40", LastCellFound.Address)
            Assert.AreEqual(mseo.LookupLastContentRowIndex(TestSheet), LastCellFound.RowIndex)
            Assert.AreEqual(mseo.LookupLastContentColumnIndex(TestSheet), LastCellFound.ColumnIndex)
        End Sub

        <Test> Public Sub AddSheet()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            Dim TestSheet As String = "Grunddaten"

            Dim BeforeSheet As String = "Grunddaten"
            Dim SheetNameTopPosition As String = "SheetOnTop"
            Dim SheetNameBottomPosition As String = "SheetOnBottom"
            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, True, True, String.Empty)
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
        End Sub

        <Test> Public Sub SheetNames()
#If CI_CD Then
            If System.Environment.OSVersion.Platform <> PlatformID.Win32NT Then Throw New IgnoreException("MS Excel not supported on Non-Windows platforms")
#End If
            Dim eppeo As ExcelOps.ExcelDataOperationsBase
            Dim mseo As ExcelOps.ExcelDataOperationsBase
            Dim TestControllingToolFileName As String = TestFiles.TestFileGrund01.FullName
            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
            mseo = New ExcelOps.MsExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, True, True, String.Empty)

            Dim EppeoSheetNamesList, MseoSheetNamesList As List(Of String)
            EppeoSheetNamesList = eppeo.SheetNames
            MseoSheetNamesList = mseo.SheetNames
            System.Console.WriteLine("EPP: " & Strings.Join(EppeoSheetNamesList.ToArray, ","))
            System.Console.WriteLine("MS : " & Strings.Join(MseoSheetNamesList.ToArray, ","))
            Assert.AreEqual(EppeoSheetNamesList.ToArray, MseoSheetNamesList.ToArray)
        End Sub

        <Test> Public Sub CalcTest_EpplusPolyform()
            ExcelOpsTests.Engines.EpplusPolyformEditionOpsTest.AssignLicenseContext()
            Dim wb As New OfficeOpenXml.ExcelPackage()
            Dim TestCell As OfficeOpenXml.ExcelRange
            wb.Workbook.Worksheets.Add("test-calcs")
            TestCell = wb.Workbook.Worksheets(0).Cells(1, 1)
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

        <Test> Public Sub CalcTest_EpplusFree()
            Dim wb As New CompuMaster.Epplus4.ExcelPackage
            Dim TestCell As CompuMaster.Epplus4.ExcelRange
            wb.Workbook.Worksheets.Add("test-calcs")
            TestCell = wb.Workbook.Worksheets(0).Cells(1, 1)
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

        Private Function CalcTestCell(cell As CompuMaster.Epplus4.ExcelRange, formula As String) As String
            cell.Formula = formula
            CompuMaster.Epplus4.CalculationExtension.Calculate(cell)
            Try
                If cell Is Nothing Then
                    Return Nothing
                ElseIf cell.Value.GetType Is GetType(CompuMaster.Epplus4.ExcelErrorValue) Then
                    Return CType(cell.Value, CompuMaster.Epplus4.ExcelErrorValue).ToString
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
            Dim wb As New MsExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, True, True, String.Empty)
            Dim SheetName As String = wb.SheetNames(0)

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
        End Sub

        <Test> Public Sub CellWithErrorEpplus()
            Dim wb As New EpplusFreeExcelDataOperations(TestFiles.TestFileGrund02.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
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

            TestControllingToolFileName = TestFiles.TestFileGrund01.FullName
            eppeo = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
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

            TestControllingToolFileNameIn = TestFiles.TestFileGrund01.FullName
            TestControllingToolFileNameOutTemplate = TestFiles.TestFileGrund01.FullName
            TestControllingToolFileNameOut = TestEnvironment.FullPathOfDynTestFile("CopySheetContentEpplus.xlsx")
            eppeoIn = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileNameIn, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
            eppeoOut = New ExcelOps.EpplusFreeExcelDataOperations(TestControllingToolFileNameOutTemplate, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)

            Console.WriteLine("Test file in: " & TestControllingToolFileNameIn)
            Console.WriteLine("Test file output template: " & TestControllingToolFileNameOutTemplate)
            Console.WriteLine("Test file output: " & TestControllingToolFileNameOut)

            Const SheetToCopy As String = "Grunddaten"
            Try
                eppeoIn.CopySheetContent(SheetToCopy, eppeoOut, ExcelOps.ExcelDataOperationsBase.CopySheetOption.TargetSheetMightExist)
                eppeoOut.SelectSheet(SheetToCopy)
                eppeoOut.SaveAs(TestControllingToolFileNameOut, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.AreEqual(eppeoIn.SheetContentMatrix(SheetToCopy, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText), eppeoOut.SheetContentMatrix(SheetToCopy, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText))
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

            TestControllingToolFileNameIn = TestFiles.TestFileGrund01.FullName
            TestControllingToolFileNameOutTemplate = TestFiles.TestFileGrund02.FullName
            TestControllingToolFileNameOut = TestEnvironment.FullPathOfDynTestFile("CopySheetContentMsExcel.xlsx")
            Try
                eppeoIn = New ExcelOps.MsExcelDataOperations(TestControllingToolFileNameIn, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, True, True, String.Empty)
                eppeoOut = New ExcelOps.MsExcelDataOperations(TestControllingToolFileNameOutTemplate, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, True, True, String.Empty)

                Console.WriteLine("Test file in: " & TestControllingToolFileNameIn)
                Console.WriteLine("Test file output template: " & TestControllingToolFileNameOutTemplate)
                Console.WriteLine("Test file output: " & TestControllingToolFileNameOut)

                Const SheetToCopy As String = "Grunddaten"
                eppeoIn.CopySheetContent(SheetToCopy, eppeoOut, ExcelOps.ExcelDataOperationsBase.CopySheetOption.TargetSheetMightExist)
                eppeoOut.SelectSheet(SheetToCopy)
                eppeoOut.SaveAs(TestControllingToolFileNameOut, ExcelDataOperationsBase.SaveOptionsForDisabledCalculationEngines.NoReset)
                Assert.AreEqual(eppeoIn.SheetContentMatrix(SheetToCopy, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText), eppeoOut.SheetContentMatrix(SheetToCopy, ExcelDataOperationsBase.MatrixContent.FormulaOrFormattedText))
                Assert.Pass("Required manual, optical review for comparison to check for formattings")
            Finally
                If eppeoOut IsNot Nothing Then eppeoOut.Close()
            End Try
        End Sub

        <Test> Public Sub ExcelOpsTestCollection_ZahlenUndProzentwerte()
            Dim eppeo As New ExcelOps.EpplusFreeExcelDataOperations(TestFiles.TestFileExcelOpsTestCollection.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, True, String.Empty)
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
            Dim mseo As New MsExcelDataOperations(TestFiles.TestFileExcelOpsTestCollection.FullName, ExcelOps.ExcelDataOperationsBase.OpenMode.OpenExistingFile, MsExcelAppWrapper, False, True, String.Empty)
            Assert.AreEqual("0.00", mseo.LookupCellFormat(SheetName, 0, 1))
            Assert.AreEqual("0.00%", mseo.LookupCellFormat(SheetName, 1, 1))
            Assert.AreEqual(10.0, System.Math.Round(mseo.LookupCellValue(Of Double)(SheetName, 0, 1), 2))
            Assert.AreEqual(0.1, System.Math.Round(mseo.LookupCellValue(Of Double)(SheetName, 1, 1), 2))
            mseo.WriteCellValue(Of Double)(SheetName, 0, 1, 20.0)
            mseo.WriteCellValue(Of Double)(SheetName, 1, 1, 0.2)
            Assert.AreEqual(20.0, System.Math.Round(mseo.LookupCellValue(Of Double)(SheetName, 0, 1), 2))
            Assert.AreEqual(0.2, System.Math.Round(mseo.LookupCellValue(Of Double)(SheetName, 1, 1), 2))
#End If
        End Sub

    End Class
End Namespace