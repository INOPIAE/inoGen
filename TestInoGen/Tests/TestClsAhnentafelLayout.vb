Imports System.Data
Imports System.IO
Imports inoGenDLL
Imports Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext
Imports NUnit.Framework

Namespace TestInoGen

    Public Class TestClsAhnentafelLayout
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private testPathSQL As String = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\TestInoGen\bin", "")

        Private DBFile As String = testPath & "\Beethoven.inoGdb"

        Private cAPL As ClsAhnentafelLayout
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
        Public Sub TestData()
            ' Test-Instanz der Klasse
            Dim layout As New ClsAhnentafelLayout((2384, 1684), False)

            ' Erste Spalte prüfen
            Dim box1 = layout.GetBox(0, 0)
            Console.WriteLine($"Erste Box X: {box1.GetX()} (erwartet: {layout.MarginLeft})")
            Assert.That(box1.GetX(), layout.MarginLeft)

            ' Letzte Spalte prüfen
            Dim box15 = layout.GetBox(0, 14)
            Dim expectedRight = layout.PageWidth - layout.MarginRight
            Console.WriteLine($"Letzte Box rechts: {box15.GetX() + box15.GetWidth()} (erwartet: {expectedRight})")
            Assert.That(box15.GetX() + box15.GetWidth(), expectedRight)

            ' Erste Zeile prüfen (oben)
            Dim boxTop = layout.GetBox(0, 0)
            Dim expectedTop = layout.PageHeight - layout.MarginTop - layout.TitleBlockHeight
            Console.WriteLine($"Erste Box oben: {boxTop.GetY() + boxTop.GetHeight()} (max: {expectedTop})")
            Assert.That(boxTop.GetY() + boxTop.GetHeight(), expectedTop)

            ' Letzte Zeile prüfen (unten)
            Dim boxBottom = layout.GetBox(14, 0)
            Console.WriteLine($"Letzte Box unten: {boxBottom.GetY()} (erwartet: {layout.MarginBottom})")
            Assert.That(boxBottom.GetY(), layout.MarginBottom)

        End Sub

        <Test>
        Public Sub TestAusgabe()

        End Sub
    End Class
End Namespace
