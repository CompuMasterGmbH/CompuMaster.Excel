Imports NUnit.Framework

Public NotInheritable Class TestFiles

    <Obsolete("Use TestFileGrund01 instead", True)> Public Shared ReadOnly Property TestFileV0SRH As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx"))
        End Get
    End Property

    <Obsolete("Use TestFileGrund02 instead", True)> Public Shared ReadOnly Property TestFileV21SampleData2019 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund02.xlsx"))
        End Get
    End Property

    Public Shared ReadOnly Property TestFileMergedCells As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsMergedCells.xlsx"))
        End Get
    End Property

    Public Shared ReadOnly Property TestFileChartSheet01 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ChartSheet01.xlsx"))
        End Get
    End Property

    Public Shared ReadOnly Property TestFileGrund01 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund01.xlsx"))
        End Get
    End Property

    Public Shared ReadOnly Property TestFileGrund02 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsGrund02.xlsx"))
        End Get
    End Property

    Public Shared ReadOnly Property TestFileExcelOpsTestCollection As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsTestCollection.xlsx"))
        End Get
    End Property

End Class