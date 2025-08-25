Imports System.IO
Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen
    Public Class TestClsGenDB
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private DBFile As String = testPath & "\Beethoven.inoGdb"
        Private cGDB As inoGenDLL.ClsGenDB
        Private cHelper As New ClsHelper
        Private testFolder As String


        <SetUp>
        Public Sub Setup()
            testFolder = cHelper.CreateTestFolder("currenttestdata")
            cHelper.DeleteTestFolder(testFolder)

            Dim DBFileTest As String = testFolder & "\Beethoven.inoGdb"
            File.Copy(DBFile, DBFileTest)
            cGDB = New inoGenDLL.ClsGenDB(DBFileTest)
        End Sub

        <TearDown>
        Public Sub TearDown()
            cHelper.DeleteTestFolder(testFolder)
        End Sub

        <Test>
        Public Sub TestFillPerson()
            Dim pd = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = 1}
            cGDB.FillPerson(pd)

            Assert.That(pd.Vorname, NUnit.Framework.Is.EqualTo("Ludwig"))
            Assert.That(pd.Nachname, NUnit.Framework.Is.EqualTo("van Beethoven"))
            Assert.That(pd.Geschlecht, NUnit.Framework.Is.EqualTo("m"))
            Assert.That(pd.Konfession, NUnit.Framework.Is.EqualTo("rk"))
            Assert.That(pd.FID, NUnit.Framework.Is.EqualTo(2))
            Assert.That(pd.PS, NUnit.Framework.Is.EqualTo("VAN LUDW1770"))

        End Sub

        <Test>
        Public Sub TestFillPersonEltern()
            Dim pd = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = 1}
            pd.FID = 1
            cGDB.FillPersonEltern(pd)

            Assert.That(pd.V, NUnit.Framework.Is.EqualTo(3))
            Assert.That(pd.M, NUnit.Framework.Is.EqualTo(4))


        End Sub

        <Test>
        Public Sub TestFillPersonDaten()
            Dim pd = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = 1}
            cGDB.FillPersonDaten(pd)

            Assert.That(pd.Geburtsdatum, NUnit.Framework.Is.Null.Or.Empty)
            Assert.That(pd.Geburtsort, NUnit.Framework.Is.Null.Or.Empty)
            Assert.That(pd.Taufdatum, NUnit.Framework.Is.EqualTo("17.12.1770"))
            Assert.That(pd.Taufort, NUnit.Framework.Is.EqualTo("Bonn"))
            Assert.That(pd.Sterbedatum, NUnit.Framework.Is.EqualTo("26.03.1827"))
            Assert.That(pd.Sterbeort, NUnit.Framework.Is.EqualTo("Wien"))
        End Sub

        <Test>
        Public Sub TestFillPersonData()
            Dim pd = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = 1}
            cGDB.FillPersonData(pd)

            Assert.That(pd.Vorname, NUnit.Framework.Is.EqualTo("Ludwig"))
            Assert.That(pd.Nachname, NUnit.Framework.Is.EqualTo("van Beethoven"))
            Assert.That(pd.Geschlecht, NUnit.Framework.Is.EqualTo("m"))
            Assert.That(pd.Konfession, NUnit.Framework.Is.EqualTo("rk"))
            Assert.That(pd.FID, NUnit.Framework.Is.EqualTo(2))
            Assert.That(pd.PS, NUnit.Framework.Is.EqualTo("VAN LUDW1770"))

            Assert.That(pd.V, NUnit.Framework.Is.EqualTo(2))
            Assert.That(pd.M, NUnit.Framework.Is.EqualTo(7))

            Assert.That(pd.Geburtsdatum, NUnit.Framework.Is.Null.Or.Empty)
            Assert.That(pd.Geburtsort, NUnit.Framework.Is.Null.Or.Empty)
            Assert.That(pd.Taufdatum, NUnit.Framework.Is.EqualTo("17.12.1770"))
            Assert.That(pd.Taufort, NUnit.Framework.Is.EqualTo("Bonn"))
            Assert.That(pd.Sterbedatum, NUnit.Framework.Is.EqualTo("26.03.1827"))
            Assert.That(pd.Sterbeort, NUnit.Framework.Is.EqualTo("Wien"))


            pd = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = 2}
            cGDB.FillPersonData(pd)

            Assert.That(pd.Vorname, NUnit.Framework.Is.EqualTo("Johann"))
            Assert.That(pd.Nachname, NUnit.Framework.Is.EqualTo("van Beethoven"))
            Assert.That(pd.Geschlecht, NUnit.Framework.Is.EqualTo("m"))
            Assert.That(pd.Konfession, NUnit.Framework.Is.Null.Or.Empty)
            Assert.That(pd.FID, NUnit.Framework.Is.EqualTo(1))
            Assert.That(pd.PS, NUnit.Framework.Is.EqualTo("VAN JOHA1792"))

            Assert.That(pd.V, NUnit.Framework.Is.EqualTo(3))
            Assert.That(pd.M, NUnit.Framework.Is.EqualTo(4))

            Assert.That(pd.Geburtsdatum, NUnit.Framework.Is.EqualTo("um 1740"))
            Assert.That(pd.Geburtsort, NUnit.Framework.Is.EqualTo("Bonn"))
            Assert.That(pd.Taufdatum, NUnit.Framework.Is.Null.Or.Empty)
            Assert.That(pd.Taufort, NUnit.Framework.Is.Null.Or.Empty)

            Assert.That(pd.Sterbedatum, NUnit.Framework.Is.EqualTo("18.12.1792"))
            Assert.That(pd.Sterbeort, NUnit.Framework.Is.EqualTo("Bonn"))

        End Sub

        <Test>
        Public Sub TestPersonenDaten()
            Dim PName As String = cGDB.PersonenDaten(1)

            Assert.That(PName, NUnit.Framework.Is.EqualTo("VAN LUDW1770 Ludwig VAN BEETHOVEN"))

        End Sub

        <Test>
        Public Sub TestFillFamilieDaten()
            Dim pd = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = 2}
            cGDB.FillPersonData(pd)
            pd.EID = 1
            cGDB.FillFamilieDaten(pd)

            Assert.That(pd.Heiratdatum, NUnit.Framework.Is.Null.Or.Empty)
            Assert.That(pd.Heiratort, NUnit.Framework.Is.Null.Or.Empty)
            Assert.That(pd.KHeiratdatum, NUnit.Framework.Is.EqualTo("07.09.1733"))
            Assert.That(pd.KHeiratort, NUnit.Framework.Is.EqualTo("Bonn"))

        End Sub
    End Class
End Namespace
