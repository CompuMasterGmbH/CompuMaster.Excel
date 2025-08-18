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

    Public Shared ReadOnly Property TestFileCircularReference01 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "CircularRefs01.xlsx"))
        End Get
    End Property

    ''' <summary>
    ''' Embedded "picture": a chart in diagram sheet
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property TestFileEmbeddedPicture01 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "EmbeddedPicture01.xlsx"))
        End Get
    End Property

    ''' <summary>
    ''' Embedded "picture": a chart in worksheet
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property TestFileEmbeddedPicture02 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "EmbeddedPicture02.xlsx"))
        End Get
    End Property

    ''' <summary>
    ''' Embedded static picture
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property TestFileEmbeddedPicture03 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "EmbeddedPicture02.xlsx"))
        End Get
    End Property

    Public Shared ReadOnly Property TestFileChartSheet01 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ChartSheet01.xlsx"))
        End Get
    End Property

    Public Shared ReadOnly Property TestFileHtmlExport01 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "HtmlExport01.xlsx"))
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

    Public Shared ReadOnly Property TestFileExcelOpsErrorValues As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "ExcelOpsErrorValues.xlsx"))
        End Get
    End Property

    Public Shared ReadOnly Property TestFileSampleTable01 As System.IO.FileInfo
        Get
            Return New System.IO.FileInfo(TestEnvironment.FullPathOfExistingTestFile("test_data", "SampleTable01.xlsx"))
        End Get
    End Property

End Class