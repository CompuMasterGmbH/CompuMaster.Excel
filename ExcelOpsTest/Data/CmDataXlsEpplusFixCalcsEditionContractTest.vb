Option Explicit On
Option Strict On

Imports System.Data
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Reflection
Imports NUnit.Framework

Namespace Data

    Public Class CmDataXlsEpplusFixCalcsEditionContractTest

        ''' <summary>
        ''' Verifies that the utility API cannot accidentally regain instance-only entry points.
        ''' </summary>
        <Test>
        Public Sub PublicApiIsStaticOnly()
            Dim apiType = GetType(CompuMaster.Data.XlsEpplusFixCalcsEdition)
            Dim instanceMethods = apiType.GetMethods(BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.DeclaredOnly)
            Dim httpMethods = apiType.GetMethods(BindingFlags.Public Or BindingFlags.Static Or BindingFlags.DeclaredOnly).
                Where(Function(methodInfo) methodInfo.Name = NameOf(CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsHttpResponse)).
                ToArray()

            Assert.That(apiType.IsSealed, [Is].True)
            Assert.That(apiType.GetConstructors(BindingFlags.Public Or BindingFlags.Instance), [Is].Empty)
            Assert.That(instanceMethods, [Is].Empty)
            Assert.That(httpMethods, Has.Length.EqualTo(3))
            Assert.That(FindHttpOverload(apiType, GetType(DataTable), GetType(String), GetType(System.Net.HttpListenerContext), GetType(CompuMaster.Data.XlsEpplusFixCalcsEdition.FileFormat), GetType(String)), [Is].Not.Null)
            Assert.That(FindHttpOverload(apiType, GetType(String), GetType(DataTable()), GetType(String()), GetType(System.Net.HttpListenerContext), GetType(CompuMaster.Data.XlsEpplusFixCalcsEdition.FileFormat)), [Is].Not.Null)
            Assert.That(FindHttpOverload(apiType, GetType(String), GetType(DataTable()), GetType(String()), GetType(System.Net.HttpListenerContext), GetType(CompuMaster.Data.XlsEpplusFixCalcsEdition.FileFormat), GetType(String)), [Is].Not.Null)
        End Sub

        ''' <summary>
        ''' Verifies that read and write defaults are immutable and validate row indexes.
        ''' </summary>
        <Test>
        Public Sub OptionsAreImmutableAndValidated()
            Dim readOptions = New CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadOptions(False, 2)
            Dim writeOptions = New CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteOptions(1)

            Assert.That(readOptions.FirstRowContainsColumnNames, [Is].False)
            Assert.That(readOptions.StartReadingAtRowIndex, [Is].EqualTo(2))
            Assert.That(writeOptions.ErrorLevel, [Is].EqualTo(1))
            Assert.That(GetType(CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadOptions).GetProperty(NameOf(readOptions.FirstRowContainsColumnNames)).CanWrite, [Is].False)
            Assert.That(GetType(CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadOptions).GetProperty(NameOf(readOptions.StartReadingAtRowIndex)).CanWrite, [Is].False)
            Assert.That(GetType(CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteOptions).GetProperty(NameOf(writeOptions.ErrorLevel)).CanWrite, [Is].False)
            Assert.Throws(Of ArgumentOutOfRangeException)(
                Sub()
                    Dim ignored = New CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadOptions(True, -1)
                End Sub)
        End Sub

        ''' <summary>
        ''' Verifies that the shared stream API creates a readable workbook.
        ''' </summary>
        <Test>
        Public Sub WriteDataTableToXlsStreamCreatesWorkbook()
            Dim table = CreateSampleTable()

            Using output As New MemoryStream()
                CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsStream(
                    Nothing,
                    output,
                    New DataTable() {table},
                    New String() {"StreamSheet"},
                    CompuMaster.Data.XlsEpplusFixCalcsEdition.FileFormat.Excel2007)

                Assert.That(output.Length, [Is].GreaterThan(0))
                output.Position = 0

                Using workbook As New CompuMaster.Epplus4.ExcelPackage(output)
                    Assert.That(workbook.Workbook.Worksheets.Count, [Is].EqualTo(1))
                    Assert.That(workbook.Workbook.Worksheets(0).Name, [Is].EqualTo("StreamSheet"))
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Verifies that the ZIP backend writes a standard XLSX archive that can be reopened.
        ''' </summary>
        <Test>
        Public Sub ZipArchiveRoundTripPreservesWorksheetsAndSharedStrings()
            Using output As New MemoryStream()
                Using workbook As New CompuMaster.Epplus4.ExcelPackage()
                    workbook.Workbook.Worksheets.Add("First").Cells(1, 1).Value = "Repeated text"
                    workbook.Workbook.Worksheets.Add("Second").Cells(1, 1).Value = "Repeated text"
                    workbook.SaveAs(output)
                End Using

                output.Position = 0
                Using archive As New ZipArchive(output, ZipArchiveMode.Read, True)
                    Assert.That(archive.GetEntry("[Content_Types].xml"), [Is].Not.Null)
                    Assert.That(archive.Entries.Any(Function(entry) entry.FullName.TrimStart("/"c) = "xl/workbook.xml"), [Is].True)
                    Assert.That(archive.Entries.Any(Function(entry) entry.FullName.TrimStart("/"c) = "xl/worksheets/sheet1.xml"), [Is].True)
                    Assert.That(archive.Entries.Any(Function(entry) entry.FullName.TrimStart("/"c) = "xl/worksheets/sheet2.xml"), [Is].True)
                    Assert.That(archive.Entries.Any(Function(entry) entry.FullName.TrimStart("/"c) = "xl/sharedStrings.xml"), [Is].True)
                End Using

                output.Position = 0
                Using reopened As New CompuMaster.Epplus4.ExcelPackage(output)
                    Assert.That(reopened.Workbook.Worksheets.Count, [Is].EqualTo(2))
                    Assert.That(reopened.Workbook.Worksheets(0).Cells(1, 1).Text, [Is].EqualTo("Repeated text"))
                    Assert.That(reopened.Workbook.Worksheets(1).Cells(1, 1).Text, [Is].EqualTo("Repeated text"))
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Verifies that the first worksheet is addressed by its zero-based index.
        ''' </summary>
        <Test>
        Public Sub ReadSingleWorksheetWithStartRowOverload()
            Dim inputPath = TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx")

            Dim result = CompuMaster.Data.XlsEpplusFixCalcsEdition.ReadDataTableFromXlsFile(inputPath, 0, False)

            Assert.That(result, [Is].Not.Null)
            Assert.That(result.Rows.Count, [Is].GreaterThan(0))
        End Sub

        ''' <summary>
        ''' Verifies that mismatched table and worksheet arrays are rejected before a file is written.
        ''' </summary>
        <Test>
        Public Sub WriteDataTablesRejectsMismatchedWorksheetNames()
            Dim outputPath = TestEnvironment.FullPathOfDynTestFile(GetType(CmDataXlsEpplusFixCalcsEditionContractTest), "test_data", "MismatchedArrays.xlsx")
            Dim tables = New DataTable() {CreateSampleTable()}

            Dim exception = Assert.Throws(Of ArgumentException)(
                Sub()
                    CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsFile(
                        Nothing,
                        outputPath,
                        tables,
                        Array.Empty(Of String)())
                End Sub)

            Assert.That(exception.ParamName, [Is].EqualTo("sheetnames"))
        End Sub

        Private Shared Function CreateSampleTable() As DataTable
            Dim table As New DataTable("Data")
            table.Columns.Add("Id", GetType(Integer))
            table.Columns.Add("Value", GetType(String))
            table.Rows.Add(1, "One")
            Return table
        End Function

        Private Shared Function FindHttpOverload(ByVal apiType As Type, ParamArray parameterTypes As Type()) As MethodInfo
            Return apiType.GetMethod(
                NameOf(CompuMaster.Data.XlsEpplusFixCalcsEdition.WriteDataTableToXlsHttpResponse),
                BindingFlags.Public Or BindingFlags.Static,
                Nothing,
                parameterTypes,
                Nothing)
        End Function

    End Class

End Namespace
