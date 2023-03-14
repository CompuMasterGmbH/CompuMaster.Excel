Option Explicit On
Option Strict On

'NOTE:    THIS FILE IS UPDATED IN FILE CmDataXlsEpplusFixCalcsEditionTest FIRST AND COPIED TO CmDataXlsEpplusPolyformEditionTest AFTERWARDS
'SEE:     clone-build-files.cmd/.sh/.ps1
'WARNING: PLEASE CHANGE THIS FILE ONLY AT REQUIRED LOCATION, OR CHANGES WILL BE LOST!

Imports NUnit.Framework

Namespace Data

    Public Class CmDataXlsEpplusFixCalcsEditionTest

        <Test> Public Sub ReadDataSetFromXlsFile()

            Dim Path As String = TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx")
            Dim t = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataSetFromXlsFile(Path, False).Tables
            Assert.AreEqual(1, t.Count)
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