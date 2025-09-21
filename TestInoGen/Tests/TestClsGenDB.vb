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

        <Test>
        Public Sub TestCalculateDatum()
            Dim testDatum As String = "12.01.1900"
            Dim result As Nullable(Of Date) = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.EqualTo(New Date(1900, 1, 12)))

            testDatum = "< 12.01.1900"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.EqualTo(New Date(1900, 1, 12)))

            testDatum = "> 12.01.1900"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.EqualTo(New Date(1900, 1, 12)))

            testDatum = "um 12.01.1900"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.EqualTo(New Date(1900, 1, 12)))

            testDatum = "um 02.1900"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.EqualTo(New Date(1900, 2, 1)))

            testDatum = "02.1900"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.EqualTo(New Date(1900, 2, 1)))

            testDatum = "um 1901"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.EqualTo(New Date(1901, 1, 1)))

            testDatum = "1901"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.EqualTo(New Date(1901, 1, 1)))

            testDatum = "nur text"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.Null)

            testDatum = "1608/1609"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.Null)

            testDatum = "1608 / 1609"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.Null)

            testDatum = "1608 - 1609"
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.Null)

            testDatum = ""
            result = cGDB.CalculateDatum(testDatum)
            Assert.That(result, NUnit.Framework.Is.Null)

        End Sub

        <Test>
        Public Sub TestGetPersonenAdditionalData()
            Dim EDL As New List(Of clsAhnentafelDaten.EventData)
            EDL = cGDB.GetPersonenAdditionalData(1)

            Assert.That(EDL.Count, [Is].EqualTo(3))

            Assert.That(EDL(0).EventLocation, [Is].EqualTo("Bonn"))
            Assert.That(EDL(0).EventID, [Is].EqualTo(9))
            Assert.That(EDL(0).Eventname, [Is].EqualTo("Beruf"))
            Assert.That(EDL(0).EventTopic, [Is].EqualTo("Bratschist"))
            Assert.That(EDL(0).EventDate, [Is].EqualTo("09.1791"))
            Assert.That(EDL(0).Person, [Is].True)

            Assert.That(EDL(1).EventLocation, [Is].EqualTo("Bonn"))
            Assert.That(EDL(1).EventID, [Is].EqualTo(9))
            Assert.That(EDL(1).Eventname, [Is].EqualTo("Beruf"))
            Assert.That(EDL(1).EventTopic, [Is].EqualTo("Organist"))
            Assert.That(EDL(1).EventDate, [Is].EqualTo("09.1791"))
            Assert.That(EDL(1).Person, [Is].True)

            Assert.That(EDL(2).EventLocation, [Is].EqualTo("Wien"))
            Assert.That(EDL(2).EventID, [Is].EqualTo(9))
            Assert.That(EDL(2).Eventname, [Is].EqualTo("Beruf"))
            Assert.That(EDL(2).EventTopic, [Is].EqualTo("Komponist"))
            Assert.That(EDL(2).EventDate, [Is].EqualTo("ab 1792"))
            Assert.That(EDL(2).Person, [Is].True)

        End Sub

        <Test>
        Public Sub TestGetFamilies()
            Dim FL As New List(Of clsAhnentafelDaten.FamilyData)
            FL = cGDB.GetFamilies(2)

            Assert.That(FL.Count, [Is].EqualTo(1))
            Assert.That(FL(0).ID, [Is].EqualTo(2))
            Assert.That(FL(0).VID, [Is].EqualTo(2))
            Assert.That(FL(0).MID, [Is].EqualTo(7))

            FL = cGDB.GetFamilies(7)

            Assert.That(FL.Count, [Is].EqualTo(2))
            Assert.That(FL(0).ID, [Is].EqualTo(68))
            Assert.That(FL(0).VID, [Is].EqualTo(128))
            Assert.That(FL(0).MID, [Is].EqualTo(7))
            Assert.That(FL(1).ID, [Is].EqualTo(2))
            Assert.That(FL(1).VID, [Is].EqualTo(2))
            Assert.That(FL(1).MID, [Is].EqualTo(7))

            FL = cGDB.GetFamilies(50)

            Assert.That(FL.Count, [Is].EqualTo(1))
            Assert.That(FL(0).ID, [Is].EqualTo(25))
            Assert.That(FL(0).VID, [Is].EqualTo(50))
            Assert.That(FL(0).MID, [Is].EqualTo(0))

            FL = cGDB.GetFamilies(1)

            Assert.That(FL.Count, [Is].EqualTo(0))

        End Sub
    End Class
End Namespace
