Imports System.IO
Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen
    Public Class TestClsOSMKarte
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private DBFile As String = testPath & "\Beethoven.inoGdb"
        Private cOK As inoGenDLL.ClsOSMKarte
        Private cHelper As New ClsHelper
        Private testFolder As String


        <SetUp>
        Public Sub Setup()
            testFolder = cHelper.CreateTestFolder("currenttestdata")
            cHelper.DeleteTestFolder(testFolder)

            Dim DBFileTest As String = testFolder & "\Beethoven.inoGdb"
            File.Copy(DBFile, DBFileTest)
            cOK = New inoGenDLL.ClsOSMKarte
        End Sub

        <TearDown>
        Public Sub TearDown()
            cHelper.DeleteTestFolder(testFolder)
        End Sub

        <Test>
        Public Sub TestFindGeoCode()
            cOK.Email = cHelper.GetEmail(Path.Combine(testPath, "TestSettings.txt"))
            Dim town As String = "Bonn"
            Dim results As List(Of ClsOSMKarte.GeoCodeResult) = cOK.FindGeoCode(town)

            Assert.That(results.Count, NUnit.Framework.Is.EqualTo(2))


            Assert.That(results(0).lon, NUnit.Framework.Is.EqualTo("7.1024635"))
            Assert.That(results(0).lat, NUnit.Framework.Is.EqualTo("50.7352621"))
            Assert.That(results(0).display_name, NUnit.Framework.Is.EqualTo("Bonn, Nordrhein-Westfalen, Deutschland"))


            Assert.That(results(1).lon, NUnit.Framework.Is.EqualTo("-2.0005556"))
            Assert.That(results(1).lat, NUnit.Framework.Is.EqualTo("46.9744444"))
            Assert.That(results(1).display_name, NUnit.Framework.Does.Contain("Bouin, Les Sables-d'Olonne, Vendée, Pays de la Loire, France"))

            town = "Wien"
            results = cOK.FindGeoCode(town)

            Assert.That(results.Count, NUnit.Framework.Is.EqualTo(1))


            Assert.That(results(0).lon, NUnit.Framework.Is.EqualTo("16.3725042"))
            Assert.That(results(0).lat, NUnit.Framework.Is.EqualTo("48.2083537"))
            Assert.That(results(0).display_name, NUnit.Framework.Is.EqualTo("Wien, Österreich"))


        End Sub
    End Class
End Namespace