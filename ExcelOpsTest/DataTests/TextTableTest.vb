Option Strict On
Option Explicit On

Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps

Namespace DataTests
    <TestFixture> Public Class TextTableTest

        <SetUp> Public Sub ResetConsoleForTestOutput()
            Test.Console.ResetConsoleForTestOutput()
        End Sub

        <Test> Public Sub BasicTestAndToCsvTable()
            Dim Table As New TextTable
            Assert.AreEqual("no rows found" & System.Environment.NewLine, Table.ToUITable)
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Console.WriteLine("## CSV")
            Console.WriteLine(Table.ToCsvTable)
            Console.WriteLine("## /CSV")
            Console.WriteLine()
            Assert.AreEqual("", Table.ToCsvTable)

            Table.AddRow("A1", "B1", "C1")
            Assert.AreEqual("""A1"",""B1"",""C1""" & ControlChars.CrLf, Table.ToCsvTable)
            Table.AddRow("A2", "B2", "C2", "D2", "", Nothing)
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Console.WriteLine("## CSV")
            Console.WriteLine(Table.ToCsvTable)
            Console.WriteLine("## /CSV")
            Console.WriteLine()
            Assert.AreEqual(
                        """A1"",""B1"",""C1"",,," & ControlChars.CrLf &
                        """A2"",""B2"",""C2"",""D2"",""""," & ControlChars.CrLf, Table.ToCsvTable)

            Table.AddRows(2)
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Console.WriteLine("## CSV")
            Console.WriteLine(Table.ToCsvTable)
            Console.WriteLine("## /CSV")
            Console.WriteLine()
            Assert.AreEqual(
                        """A1"",""B1"",""C1"",,," & ControlChars.CrLf &
                        """A2"",""B2"",""C2"",""D2"",""""," & ControlChars.CrLf &
                        ",,,,," & ControlChars.CrLf &
                        ",,,,," & ControlChars.CrLf, Table.ToCsvTable)
            Assert.AreEqual(1, Table.LastContentRowIndex)
            Assert.AreEqual(3, Table.LastContentColumnIndex)

            Table.AutoTrim()
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Console.WriteLine("## CSV")
            Console.WriteLine(Table.ToCsvTable)
            Console.WriteLine("## /CSV")
            Console.WriteLine()
            Assert.AreEqual(
                        """A1"",""B1"",""C1""," & ControlChars.CrLf &
                        """A2"",""B2"",""C2"",""D2""" & ControlChars.CrLf, Table.ToCsvTable)
        End Sub

        <Test> Public Sub Cell()
            Dim Table As New TextTable
            Table.AddColumns("A", "B", "C", "D", "E", "F")
            Table.AddRow("A1", "B1", "C1")
            Table.AddRow("A2", "B2", "C2", "D2", "", Nothing)

            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Console.WriteLine()

            Assert.AreEqual("A1", Table.Cell(0, 0))
            Assert.AreEqual("A1", Table.Cell(0, "A"))
            Assert.AreEqual("B2", Table.Cell(1, 1))
            Assert.AreEqual("B2", Table.Cell(1, "B"))
        End Sub

        <Test> Public Sub ToUITable()
            Dim Table As New TextTable
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Console.WriteLine()
            Assert.AreEqual("no rows found" & System.Environment.NewLine, Table.ToUITable)

            Table.AddRow("A1", "B1", "C1")
            Table.AddRow("A2", "B2", "C2", "D2", "", Nothing)
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Assert.AreEqual(
                        "Column1|Column2|Column3|Column4|Column5|Column6" & System.Environment.NewLine &
                        "-------+-------+-------+-------+-------+-------" & System.Environment.NewLine &
                        "A1     |B1     |C1     |       |       |       " & System.Environment.NewLine &
                        "A2     |B2     |C2     |D2     |       |       " & System.Environment.NewLine,
                        Table.ToUITable)

            Table.AddRows(2)
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Console.WriteLine()
            Assert.AreEqual(
                        "Column1|Column2|Column3|Column4|Column5|Column6" & System.Environment.NewLine &
                        "-------+-------+-------+-------+-------+-------" & System.Environment.NewLine &
                        "A1     |B1     |C1     |       |       |       " & System.Environment.NewLine &
                        "A2     |B2     |C2     |D2     |       |       " & System.Environment.NewLine &
                        "       |       |       |       |       |       " & System.Environment.NewLine &
                        "       |       |       |       |       |       " & System.Environment.NewLine,
                        Table.ToUITable)

            Table.AutoTrim()
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUITable)
            Console.WriteLine("## /UI")
            Assert.AreEqual(
                        "Column1|Column2|Column3|Column4" & System.Environment.NewLine &
                        "-------+-------+-------+-------" & System.Environment.NewLine &
                        "A1     |B1     |C1     |       " & System.Environment.NewLine &
                        "A2     |B2     |C2     |D2     " & System.Environment.NewLine,
                        Table.ToUITable)
        End Sub

        <Test> Public Sub ToUIExcelTable()
            Dim Table As New TextTable
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUIExcelTable)
            Console.WriteLine("## /UI")
            Console.WriteLine()
            Assert.AreEqual("no rows found" & System.Environment.NewLine, Table.ToUIExcelTable)

            Table.AddRow("A1", "B1", "C1")
            Table.AddRow("A2", "B2", "C2", "D2", "", Nothing)
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUIExcelTable)
            Console.WriteLine("## /UI")
            Console.WriteLine()
            Assert.AreEqual(
                        "# |A |B |C |D |E |F " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |A1|B1|C1|  |  |  " & System.Environment.NewLine &
                        "2 |A2|B2|C2|D2|  |  " & System.Environment.NewLine,
                        Table.ToUIExcelTable)

            Table.AddRows(2)
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUIExcelTable)
            Console.WriteLine("## /UI")
            Console.WriteLine()
            Assert.AreEqual(
                        "# |A |B |C |D |E |F " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |A1|B1|C1|  |  |  " & System.Environment.NewLine &
                        "2 |A2|B2|C2|D2|  |  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  " & System.Environment.NewLine,
                        Table.ToUIExcelTable)

            Table.AutoTrim()
            Console.WriteLine("## UI")
            Console.WriteLine(Table.ToUIExcelTable)
            Console.WriteLine("## /UI")
            Assert.AreEqual(
                        "# |A |B |C |D " & System.Environment.NewLine &
                        "--+--+--+--+--" & System.Environment.NewLine &
                        "1 |A1|B1|C1|  " & System.Environment.NewLine &
                        "2 |A2|B2|C2|D2" & System.Environment.NewLine,
                        Table.ToUIExcelTable)
        End Sub

        <Test> Public Sub ExcelColumnName()
            Assert.Multiple(
                Sub()
                    Assert.AreEqual("A", TextTable.ExcelColumnName(0))
                    Assert.AreEqual("Z", TextTable.ExcelColumnName(25))
                    Assert.AreEqual("AA", TextTable.ExcelColumnName(26))
                    Assert.AreEqual("XFD", TextTable.ExcelColumnName(16383))
                    Assert.AreEqual("A", ExcelOps.ExcelCell.ExcelColumnName(0))
                    Assert.AreEqual("Z", ExcelOps.ExcelCell.ExcelColumnName(25))
                    Assert.AreEqual("AA", ExcelOps.ExcelCell.ExcelColumnName(26))
                    Assert.AreEqual("XFD", ExcelOps.ExcelCell.ExcelColumnName(16383))
                    Assert.AreEqual("BY", ExcelCell.ExcelColumnName(76)) 'Not: "C@"
                    Assert.AreEqual("BZ", ExcelCell.ExcelColumnName(77)) 'Not: "C@"
                    Assert.AreEqual("CA", ExcelCell.ExcelColumnName(78)) 'Not: "C@"
                End Sub)
        End Sub

        <Test> Public Sub CellAddress()
            Dim Table As New TextTable
            Assert.AreEqual("A1", TextTable.CellAddress(0, 0))
            Assert.AreEqual("A2", TextTable.CellAddress(1, 0))
            Assert.AreEqual("B1", TextTable.CellAddress(0, 1))
            Assert.AreEqual("B2", TextTable.CellAddress(1, 1))
            Assert.AreEqual("A1", ExcelOps.ExcelCell.LocalCellAddress(0, 0))
            Assert.AreEqual("A2", ExcelOps.ExcelCell.LocalCellAddress(1, 0))
            Assert.AreEqual("B1", ExcelOps.ExcelCell.LocalCellAddress(0, 1))
            Assert.AreEqual("B2", ExcelOps.ExcelCell.LocalCellAddress(1, 1))
        End Sub

        <Test> Public Sub CompareCells()
            Dim Table As New TextTable
            Table.AddRow("A1", "B1", "C1")
            Table.AddRow("A2", "B2", "C2", "D2", "", Nothing)
            Table.AddRows(2)
            Console.WriteLine("## Table 1")
            Console.WriteLine(Table.ToUIExcelTable)
            Console.WriteLine("## /Table 1")

            Dim CompTable As New TextTable
            CompTable.AddRow("A1Changed", "B1", "C1")
            CompTable.AddRow("A2", "B2", "C2", "D2Changed", "Changed", "Changed", "New", Nothing) 'Last Nothing value is a new column, but text comparison is equal, so expected to not appear as a difference
            CompTable.AddRows(3)
            Console.WriteLine("## Table 2")
            Console.WriteLine(CompTable.ToUIExcelTable)
            Console.WriteLine("## /Table 2")

            Dim CompTableWithRemovedCellData As New TextTable
            CompTableWithRemovedCellData.AddRow("A1Changed", "B1", " ")
            CompTableWithRemovedCellData.AddRow(Nothing, "B2", "C2", "", "Changed", "Changed", "New", Nothing) 'Last Nothing value is a new column, but text comparison is equal, so expected to not appear as a difference
            CompTableWithRemovedCellData.AddRows(1)
            Console.WriteLine("## Table 2")
            Console.WriteLine(CompTable.ToUIExcelTable)
            Console.WriteLine("## /Table 2")

            Dim DiffTable As TextTable

            DiffTable = Table.CompareCells(CompTable, TextTable.DiffMode.DifferentTrimmedCells, TextTable.DiffCellOutput.Bool)
            Console.WriteLine("## Table 1 DiffCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 DiffCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |D |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "2 |  |  |  |D |D |D |D |  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "5 |  |  |  |  |  |  |  |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = Table.CompareCells(CompTable, TextTable.DiffMode.EqualTrimmedCellsWithContent, TextTable.DiffCellOutput.Bool)
            Console.WriteLine("## Table 1 EqualCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 EqualCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |  |E |E |E |E |E |E |E " & System.Environment.NewLine &
                        "2 |E |E |E |  |  |  |  |E " & System.Environment.NewLine &
                        "3 |E |E |E |E |E |E |E |E " & System.Environment.NewLine &
                        "4 |E |E |E |E |E |E |E |E " & System.Environment.NewLine &
                        "5 |E |E |E |E |E |E |E |E " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = Table.CompareCells(CompTable, TextTable.DiffMode.EqualTrimmedCellsWithContent, TextTable.DiffCellOutput.Bool, 1, 2)
            Console.WriteLine("## Table 1 EqualCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 EqualCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "2 |  |  |E |  |  |  |  |E " & System.Environment.NewLine &
                        "3 |  |  |E |E |E |E |E |E " & System.Environment.NewLine &
                        "4 |  |  |E |E |E |E |E |E " & System.Environment.NewLine &
                        "5 |  |  |E |E |E |E |E |E " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = Table.CompareCells(CompTable, TextTable.DiffMode.EqualTrimmedCellsWithContent, TextTable.DiffCellOutput.Bool, 1, 2, 3, 6)
            Console.WriteLine("## Table 1 EqualCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 EqualCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "2 |  |  |E |  |  |  |  |  " & System.Environment.NewLine &
                        "3 |  |  |E |E |E |E |E |  " & System.Environment.NewLine &
                        "4 |  |  |E |E |E |E |E |  " & System.Environment.NewLine &
                        "5 |  |  |  |  |  |  |  |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = Table.CompareCells(CompTable, TextTable.DiffMode.DifferentTrimmedCells, TextTable.DiffCellOutput.ChangeType)
            Console.WriteLine("## Table 1 DiffCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 DiffCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |M |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "2 |  |  |  |M |R |R |RC|  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "5 |  |  |  |  |  |  |  |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = Table.CompareCells(CompTableWithRemovedCellData, TextTable.DiffMode.DifferentTrimmedCells, TextTable.DiffCellOutput.ChangeType)
            Console.WriteLine("## Table 1 DiffCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 DiffCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |M |  |A |  |  |  |  |  " & System.Environment.NewLine &
                        "2 |A |  |  |A |R |R |RC|  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  |  |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = CompTableWithRemovedCellData.CompareCells(Table, TextTable.DiffMode.DifferentTrimmedCells, TextTable.DiffCellOutput.ChangeType)
            Console.WriteLine("## Table 1 DiffCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 DiffCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |M |  |R |  |  |  |  |  " & System.Environment.NewLine &
                        "2 |R |  |  |R |A |A |AC|  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  |  |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = CompTable.CompareCells(Table, TextTable.DiffMode.DifferentTrimmedCells, TextTable.DiffCellOutput.Bool)
            Console.WriteLine("## Table 2 DiffCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 2 DiffCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |D |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "2 |  |  |  |D |D |D |D |  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "5 |  |  |  |  |  |  |  |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = CompTable.CompareCells(Table, TextTable.DiffMode.EqualTrimmedCellsWithContent, TextTable.DiffCellOutput.Bool)
            Console.WriteLine("## Table 2 EqualCells - Bool")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 2 EqualCells - Bool")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |  |E |E |E |E |E |E |E " & System.Environment.NewLine &
                        "2 |E |E |E |  |  |  |  |E " & System.Environment.NewLine &
                        "3 |E |E |E |E |E |E |E |E " & System.Environment.NewLine &
                        "4 |E |E |E |E |E |E |E |E " & System.Environment.NewLine &
                        "5 |E |E |E |E |E |E |E |E " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = Table.CompareCells(CompTable, TextTable.DiffMode.DifferentTrimmedCells, TextTable.DiffCellOutput.CellContentOfThisTable)
            Console.WriteLine("## Table 1 DiffCells - Content")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 DiffCells - Content")
            Console.WriteLine("## Table 1 DiffCells (CSV) - Content")
            Console.WriteLine(DiffTable.ToCsvTable)
            Console.WriteLine("## /Table 1 DiffCells (CSV) - Content")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G  |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+---+--" & System.Environment.NewLine &
                        "1 |A1|  |  |  |  |  |   |  " & System.Environment.NewLine &
                        "2 |  |  |  |D2|  |  |New|  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  |   |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  |   |  " & System.Environment.NewLine &
                        "5 |  |  |  |  |  |  |   |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)
            Assert.IsNull(DiffTable.Cell(0, 4))
            Assert.IsNull(DiffTable.Cell(0, 5))
            Assert.IsNotEmpty(DiffTable.Cell(0, 4))
            Assert.IsNotEmpty(DiffTable.Cell(0, 5))
            Assert.IsNotNull(DiffTable.Cell(1, 4))
            Assert.IsNotNull(DiffTable.Cell(1, 5))
            Assert.IsEmpty(DiffTable.Cell(1, 4))
            Assert.IsEmpty(DiffTable.Cell(1, 5))

            DiffTable = Table.CompareCells(CompTable, TextTable.DiffMode.EqualTrimmedCellsWithContent, TextTable.DiffCellOutput.CellContentOfThisTable)
            Console.WriteLine("## Table 1 EqualCells - Content")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 1 EqualCells - Content")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |  |B1|C1|  |  |  |  |  " & System.Environment.NewLine &
                        "2 |A2|B2|C2|  |  |  |  |  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "5 |  |  |  |  |  |  |  |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = CompTable.CompareCells(Table, TextTable.DiffMode.DifferentTrimmedCells, TextTable.DiffCellOutput.CellContentOfThisTable)
            Console.WriteLine("## Table 2 DiffCells - Content")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 2 DiffCells - Content")
            Assert.AreEqual(
                        "# |A        |B |C |D        |E      |F      |G  |H " & System.Environment.NewLine &
                        "--+---------+--+--+---------+-------+-------+---+--" & System.Environment.NewLine &
                        "1 |A1Changed|  |  |         |       |       |   |  " & System.Environment.NewLine &
                        "2 |         |  |  |D2Changed|Changed|Changed|New|  " & System.Environment.NewLine &
                        "3 |         |  |  |         |       |       |   |  " & System.Environment.NewLine &
                        "4 |         |  |  |         |       |       |   |  " & System.Environment.NewLine &
                        "5 |         |  |  |         |       |       |   |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

            DiffTable = CompTable.CompareCells(Table, TextTable.DiffMode.EqualTrimmedCellsWithContent, TextTable.DiffCellOutput.CellContentOfThisTable)
            Console.WriteLine("## Table 2 EqualCells - Content")
            Console.WriteLine(DiffTable.ToUIExcelTable)
            Console.WriteLine("## /Table 2 EqualCells - Content")
            Assert.AreEqual(
                        "# |A |B |C |D |E |F |G |H " & System.Environment.NewLine &
                        "--+--+--+--+--+--+--+--+--" & System.Environment.NewLine &
                        "1 |  |B1|C1|  |  |  |  |  " & System.Environment.NewLine &
                        "2 |A2|B2|C2|  |  |  |  |  " & System.Environment.NewLine &
                        "3 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "4 |  |  |  |  |  |  |  |  " & System.Environment.NewLine &
                        "5 |  |  |  |  |  |  |  |  " & System.Environment.NewLine,
                        DiffTable.ToUIExcelTable)

        End Sub

        <Test> Public Sub CellExists()
            Dim Table As New TextTable
            Assert.False(Table.CellExists(0, 0))
            Table.AddRow("A1", "B1", "C1")
            Table.AddRow("A2", "B2", "C2", "D2", "", Nothing)
            Table.AddRows(2)
            Console.WriteLine("## Table")
            Console.WriteLine(Table.ToUIExcelTable)
            Console.WriteLine("## /Table")
            Assert.True(Table.CellExists(0, 0))
            Assert.True(Table.CellExists(3, 5))
            Assert.False(Table.CellExists(4, 6))
            Assert.False(Table.CellExists(3, 6))
            Assert.False(Table.CellExists(4, 5))
        End Sub

        <Test> Public Sub EqualsAndHashCodes()
            Dim Table1 As New TextTable
            Table1.AddRow("A1", "B1", "C1")
            Table1.AddRow("A2", "B2", "C2", "D2", "", Nothing)
            Dim Table2 As New TextTable
            Table2.AddRow("A1", "B1", "C1")
            Table2.AddRow("A2", "B2", "C2", "D2", "", Nothing)
            Table2.AddRows(2)
            Dim Table3 As New TextTable
            Table3.AddRow("A1", "B1", "C1")
            Table3.AddRow("A2", "B2", "C2", "D2", "", Nothing)
            Table3.AddRow("A3")
            Assert.AreEqual(Table1, Table2)
            Assert.AreNotEqual(Table1, Table3)
            Assert.AreNotEqual(Table3, Table2)
            Assert.IsTrue(Table1 = Table2)
            Assert.IsFalse(Table1 = Table3)
            Assert.IsFalse(Table3 = Table2)
        End Sub

    End Class
End Namespace