Imports System.Data
Imports System.IO
Imports inoGenDLL
Imports Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext
Imports NUnit.Framework

Namespace TestInoGen

    Public Class TestClsAhnentafelPDF
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private testPathSQL As String = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\TestInoGen\bin", "")

        Private DBFile As String = testPath & "\Beethoven.inoGdb"

        Private cAP As New ClsAhnentafelPDF
        Private cHelper As New ClsHelper
        Private testFolder As String
        Private currentDBVersion As Long = 8

        <SetUp>
        Public Sub Setup()
            testFolder = cHelper.CreateTestFolder("currenttestdata")
        End Sub

        <TearDown>
        Public Sub TearDown()
            cHelper.DeleteTestFolder(testFolder)
        End Sub

        <Test>
        Public Sub TestDefGen7()

            Dim testK As New Dictionary(Of Integer, ClsAhnentafelPDF.Koordinaten)
            testK = cAP.CreateAhnentafelDictionaryGen7

            For i As Long = 1 To 127
                File.AppendAllText("D:\test\test.log", "I: " & i & " Zeile " & testK(i).KZeile & " Spalte " & testK(i).KSpalte & Environment.NewLine)
            Next
        End Sub
    End Class
End Namespace
