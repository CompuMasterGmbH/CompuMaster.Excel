Imports NUnit.Framework
Imports NUnit.Framework.Legacy
Imports System.Drawing
#Disable Warning BC40056
#Disable Warning IDE0005
Imports XnaFan.ImageComparison
Imports ImageComparison
#Enable Warning IDE0005
#Enable Warning BC40056

Public Class TestImageComparison

#Disable Warning CA1416
    ''' <summary>
    ''' Test 2 images for equality
    ''' </summary>
    ''' <param name="expected"></param>
    ''' <param name="current"></param>
    ''' <remarks>WARNING: Platform support for image comparison feature only provided by windows</remarks>
    Public Shared Sub AssertImagesAreEqual(expected As System.Drawing.Image, current As System.Drawing.Image, Optional acceptedTolerance As Single = 0F)
        '1st: compare image dimensions
        ClassicAssert.AreEqual(expected.Size, current.Size, "Image pixel size difference")

        '2nd: compare pixels
        Dim Threshold As Byte = 0 'https://www.codeproject.com/Articles/374386/Simple-image-comparison-in-NET recommends thresholds between 0 - 4
        'Dim DiffPercentage = Global.ImageComparison.ExtensionMethods.PercentageDifference(expected, current, Threshold)
        Dim DiffPercentage = PercentageDifference(expected, current, Threshold, expected.Width, expected.Height)
        Dim MaxDiff = acceptedTolerance
        Console.WriteLine("Image diff [%]: " & (DiffPercentage * 100).ToString)
        ClassicAssert.LessOrEqual(DiffPercentage, MaxDiff, "Image difference > " & MaxDiff & " found: " & (DiffPercentage * 100).ToString & " %")
    End Sub

    <Test>
    Public Shared Sub AssertImagesAreEqualTest()
        If Not TestTools.IsWindowsPlatform Then
            ClassicAssert.Ignore("Platform not supported for image comparison feature (supported only by windows)")
        Else
            Dim MasterImg As Image = System.Drawing.Image.FromFile(TestEnvironment.FullPathOfExistingTestFile("test_comparison_masters", "excel_test_chart.png"))
            Dim ComparisonImg As New Bitmap(MasterImg)

            '1st: still equal
            AssertImagesAreEqual(MasterImg, ComparisonImg)

            '2nd: now it must fail
            For x As Integer = 0 To 50
                For y As Integer = 0 To 50
                    ComparisonImg.SetPixel(x, y, System.Drawing.Color.Blue)
                Next
            Next
            ClassicAssert.Throws(Of NUnit.Framework.AssertionException)(Sub() AssertImagesAreEqual(MasterImg, ComparisonImg))
        End If
    End Sub

    Private Shared Function PercentageDifference(ByVal img1 As Image, ByVal img2 As Image, ByVal Optional threshold As Byte = 3, Optional voxelSizeX As Integer = 16, Optional voxelSizeY As Integer = 16) As Single
        Dim differences As Byte(,) = GetDifferences(img1, img2, voxelSizeX, voxelSizeY)
        Dim diffPixels As Integer = 0

        For Each b As Byte In differences

            If b > threshold Then
                diffPixels += 1
            End If
        Next

        Return CType(diffPixels / (voxelSizeX * voxelSizeY), Single)
    End Function

    Private Shared Function GetDifferences(ByVal img1 As Image, ByVal img2 As Image, Optional voxelSizeX As Integer = 16, Optional voxelSizeY As Integer = 16) As Byte(,)
        Dim bitmap As Bitmap = CType(img1.Resize(voxelSizeX, voxelSizeY).GetGrayScaleVersion(), Bitmap)
        Dim bitmap2 As Bitmap = CType(img2.Resize(voxelSizeX, voxelSizeY).GetGrayScaleVersion(), Bitmap)
        Dim array As Byte(,) = New Byte(voxelSizeX - 1, voxelSizeY - 1) {}
        Dim grayScaleValues As Byte(,) = GetGrayScaleValues(bitmap, voxelSizeX, voxelSizeY)
        Dim grayScaleValues2 As Byte(,) = GetGrayScaleValues(bitmap2, voxelSizeX, voxelSizeY)

        For x As Integer = 0 To voxelSizeX - 1
            For y As Integer = 0 To voxelSizeY - 1
                If grayScaleValues(x, y) > grayScaleValues2(x, y) Then
                    array(x, y) = CByte(Math.Abs(grayScaleValues(x, y) - grayScaleValues2(x, y)))
                Else
                    array(x, y) = CByte(Math.Abs(grayScaleValues2(x, y) - grayScaleValues(x, y)))
                End If
            Next
        Next
        Return array
    End Function

    Private Shared Function GetGrayScaleValues(ByVal img As Image, Optional voxelSizeX As Integer = 16, Optional voxelSizeY As Integer = 16) As Byte(,)
        Dim array As Byte(,) = New Byte(voxelSizeX - 1, voxelSizeY - 1) {}
        Using bitmap As Bitmap = CType(img.Resize(voxelSizeX, voxelSizeY).GetGrayScaleVersion(), Bitmap)

            For x As Integer = 0 To voxelSizeX - 1
                For y As Integer = 0 To voxelSizeY - 1
                    array(x, y) = CByte(Math.Abs(bitmap.GetPixel(x, y).R))
                Next
            Next
        End Using

        Return array
    End Function
#Enable Warning CA1416

End Class
