Imports NUnit.Framework
Imports CompuMaster.Excel.ExcelOps

Namespace ExcelOpsEngineTests
    <TestFixture> Public Class ExcelRangeTest

        <SetUp> Public Sub ResetConsoleForTestOutput()
            Test.Console.ResetConsoleForTestOutput()
        End Sub

        <Test> Public Sub ExcelRange()
            Dim Cell1 As New ExcelOps.ExcelCell("Grunddaten", "A1", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Cell2 As New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All)
            Dim Range As New ExcelOps.ExcelRange(Cell1, Cell2)
            Assert.AreEqual("A1:B3", Range.ToString(False))
            Assert.AreEqual("'Grunddaten'!A1:B3", Range.ToString(True))
            Assert.AreEqual("'Grunddaten'!A1:B3", Range.ToString())
            Assert.AreEqual(1, New ExcelOps.ExcelRange(Cell2, Cell2).CellCount)
            Assert.AreEqual(6, New ExcelOps.ExcelRange(Cell1, Cell2).CellCount)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Dim InvalidRange As New ExcelOps.ExcelRange(Cell2, Cell1)
                                               End Sub)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Dim InvalidRange As New ExcelOps.ExcelRange(
                                                        New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All),
                                                        New ExcelOps.ExcelCell("Grunddaten", "A3", ExcelOps.ExcelCell.ValueTypes.All)
                                                        )
                                               End Sub)
            Assert.Catch(Of ArgumentException)(Sub()
                                                   Dim InvalidRange As New ExcelOps.ExcelRange(
                                                        New ExcelOps.ExcelCell("Grunddaten", "B3", ExcelOps.ExcelCell.ValueTypes.All),
                                                        New ExcelOps.ExcelCell("Grunddaten", "B2", ExcelOps.ExcelCell.ValueTypes.All)
                                                        )
                                               End Sub)
            Dim MyExcelCellEnumCounter As Integer = 0
            For Each MyExcelCell As ExcelCell In Range
                MyExcelCellEnumCounter += 1
                Select Case MyExcelCellEnumCounter
                    Case 1
                        Assert.AreEqual("A1", MyExcelCell.ToString(False))
                    Case 2
                        Assert.AreEqual("B1", MyExcelCell.ToString(False))
                    Case 3
                        Assert.AreEqual("A2", MyExcelCell.ToString(False))
                    Case 4
                        Assert.AreEqual("B2", MyExcelCell.ToString(False))
                    Case 5
                        Assert.AreEqual("A3", MyExcelCell.ToString(False))
                    Case 6
                        Assert.AreEqual("B3", MyExcelCell.ToString(False))
                    Case Else
                        Assert.Fail("Wrong cell count returned by ExcelRange enumerator")
                End Select
            Next
            Assert.AreEqual(6, MyExcelCellEnumCounter)
            Assert.AreEqual(6, Range.CellCount)
            Dim ClonedRange As ExcelRange = Range.Clone
            ClonedRange.SheetName = "Replaced"
            ClonedRange.AddressStart.Address = "Z5"
            ClonedRange.AddressEnd.Address = "Z7"
            Assert.AreEqual(3, ClonedRange.CellCount)
            Assert.AreEqual("Replaced", ClonedRange.SheetName)
            Assert.AreEqual(6, Range.CellCount)
            Assert.AreEqual("Grunddaten", Range.SheetName)
        End Sub

    End Class

End Namespace