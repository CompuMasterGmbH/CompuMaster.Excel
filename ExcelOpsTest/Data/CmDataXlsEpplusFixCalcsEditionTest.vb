Option Explicit On
Option Strict On

'NOTE:    THIS FILE IS UPDATED IN FILE CmDataXlsEpplusFixCalcsEditionTest FIRST AND COPIED TO CmDataXlsEpplusPolyformEditionTest AFTERWARDS
'SEE:     clone-build-files.cmd/.sh/.ps1
'WARNING: PLEASE CHANGE THIS FILE ONLY AT REQUIRED LOCATION, OR CHANGES WILL BE LOST!

Imports NUnit.Framework
Imports NUnit.Framework.Legacy
Imports System.Data
Imports System.IO
Imports System.Linq

Namespace Data

    Public Class CmDataXlsEpplusFixCalcsEditionTest

        <Test> Public Sub ReadDataSetFromXlsFile()

            Dim Path As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx")
            Dim t = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataSetFromXlsFile(Path, False).Tables
            ClassicAssert.AreEqual(1, t.Count)
        End Sub

        ''' <summary>
        ''' Verifies that the options overload applies the requested zero-based start row and header behavior.
        ''' </summary>
        <Test>
        Public Sub ReadSingleWorksheetWithOptions()
            Dim inputPath = TestEnvironment.FullPathOfDynTestFile(GetType(CmDataXlsEpplusFixCalcsEditionTest), "OptionsStartRow.xlsx")
            Using workbook As New CompuMaster.Epplus4.ExcelPackage()
                Dim worksheet = workbook.Workbook.Worksheets.Add("Data")
                worksheet.Cells(1, 1).Value = "Introduction"
                worksheet.Cells(2, 1).Value = "Name"
                worksheet.Cells(2, 2).Value = "Value"
                worksheet.Cells(3, 1).Value = "First"
                worksheet.Cells(3, 2).Value = 42
                workbook.SaveAs(New FileInfo(inputPath))
            End Using
            Dim options = New CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadOptions(True, 1)

            Dim result = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFileWithOptions(inputPath, options)

            Assert.That(result.Columns.Cast(Of DataColumn)().Select(Function(column) column.ColumnName), [Is].EqualTo(New String() {"Name", "Value"}))
            Assert.That(result.Rows, Has.Count.EqualTo(1))
            Assert.That(result.Rows(0)("Name"), [Is].EqualTo("First"))
            Assert.That(result.Rows(0)("Value"), [Is].EqualTo(42))
        End Sub

        Private Function SampleTableDyn01() As System.Data.DataTable
            Dim t1 As New System.Data.DataTable("test")
            t1.Columns.Add()
            t1.Columns.Add()
            t1.Columns.Add()
            Dim r = t1.NewRow
            r.ItemArray = New Object() {"1", "R1", "V1"}
            t1.Rows.Add(r)
            r = t1.NewRow
            r.ItemArray = New Object() {"2", "R2", "V2"}
            t1.Rows.Add(r)
            r = t1.NewRow
            r.ItemArray = New Object() {"3", "R3", "V3"}
            t1.Rows.Add(r)
            Return t1
        End Function

        <Test> Public Sub WriteDataTableToXlsFileAndFirstSheet()
            Dim PathIn As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx")
            Dim PathOut As String

            PathOut = TestEnvironment.FullPathOfDynTestFile(GetType(CmDataXlsEpplusFixCalcsEditionTest), "test_data", "SampleTableDyn-5643857.xlsx")
            System.Console.WriteLine("Writing to file: " & PathOut)
            Dim t1 = SampleTableDyn01()
            CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFileAndFirstSheet(PathOut, t1)

            PathOut = TestEnvironment.FullPathOfDynTestFile(GetType(CmDataXlsEpplusFixCalcsEditionTest), "test_data", "SampleTable01-rewritten.xlsx")
            Dim t = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataSetFromXlsFile(PathIn, False).Tables
            CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFileAndFirstSheet(PathOut, t(0))
        End Sub

        <Test> Public Sub WriteDataTableToXlsFileAndCurrentSheet()
            Dim PathIn As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx")
            Dim PathOut As String

            PathOut = TestEnvironment.FullPathOfDynTestFile(GetType(CmDataXlsEpplusFixCalcsEditionTest), "test_data", "SampleTableDyn-65779925.xlsx")
            System.Console.WriteLine("Writing to file: " & PathOut)
            Dim t1 = SampleTableDyn01()
            CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFileAndCurrentSheet(PathOut, t1)

            PathOut = TestEnvironment.FullPathOfDynTestFile(GetType(CmDataXlsEpplusFixCalcsEditionTest), "test_data", "SampleTable01-rewritten.xlsx")
            Dim t = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataSetFromXlsFile(PathIn, False).Tables
            CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFileAndCurrentSheet(PathOut, t(0))
        End Sub

        <Test> Public Sub WriteDataTableToXlsFile()
            Dim PathIn As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx")
            Dim PathOut As String

            PathOut = TestEnvironment.FullPathOfDynTestFile(GetType(CmDataXlsEpplusFixCalcsEditionTest), "test_data", "SampleTableDyn-97662114.xlsx")
            System.Console.WriteLine("Writing to file: " & PathOut)
            Dim t1 = SampleTableDyn01()
            CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFile(PathOut, t1)

            CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFile(PathIn, PathOut, New System.Data.DataTable() {}, New String() {})

            PathOut = TestEnvironment.FullPathOfDynTestFile(GetType(CmDataXlsEpplusFixCalcsEditionTest), "test_data", "SampleTable01-rewritten.xlsx")
            Dim t = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataSetFromXlsFile(PathIn, False).Tables
            CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFileAndFirstSheet(PathOut, t(0))
        End Sub

    End Class

End Namespace
