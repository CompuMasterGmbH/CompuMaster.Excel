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

    End Class
End Namespace