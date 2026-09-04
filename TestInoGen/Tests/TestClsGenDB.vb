Imports System.IO
Imports NUnit.Framework
Imports inoGenDLL
Imports System.Data

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
            DBFileTest = testFolder & "\TestVK.inoGdb"
            File.Copy(testPath & "\TestVK.inoGdb", DBFileTest)
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
            Assert.That(pd.PS, NUnit.Framework.Is.EqualTo("VAN JOHA1740"))

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

        <Test>
        Public Sub TestStatisicsVKHeirat()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.StatisicsVKHeirat()

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo(3))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo(3))
            Assert.That(dt.Rows(0).Item(2), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item(3), NUnit.Framework.Is.EqualTo(0))
            Assert.That(dt.Rows(0).Item(4), NUnit.Framework.Is.EqualTo(3))
            Assert.That(dt.Rows(0).Item(5), NUnit.Framework.Is.EqualTo(0))
            Assert.That(dt.Rows(0).Item(6), NUnit.Framework.Is.EqualTo(0))
            Assert.That(dt.Rows(0).Item(7), NUnit.Framework.Is.EqualTo(0))
            Assert.That(dt.Rows(0).Item(8), NUnit.Framework.Is.EqualTo(0))
            Assert.That(dt.Rows(0).Item(9), NUnit.Framework.Is.EqualTo(0))
            Assert.That(dt.Rows(0).Item(10), NUnit.Framework.Is.EqualTo(0))


        End Sub

        <Test>
        Public Sub TestStatisicsVornamen()
            Dim result As Integer = cGDB.StatisicsVornamen

            Assert.That(result, NUnit.Framework.Is.EqualTo(79))

        End Sub

        <Test>
        Public Sub TestStatisicsNachnamen()
            Dim result As Integer = cGDB.StatisicsNachnamen

            Assert.That(result, NUnit.Framework.Is.EqualTo(64))

        End Sub

        <Test>
        Public Sub TestStatisicsOrt()
            Dim result As Integer = cGDB.StatisicsOrte

            Assert.That(result, NUnit.Framework.Is.EqualTo(35))

        End Sub

        <Test>
        Public Sub TestStatisicsPersonen()
            Dim result As Integer = cGDB.StatisicsPersonen

            Assert.That(result, NUnit.Framework.Is.EqualTo(135))

        End Sub

        <Test>
        Public Sub TestStatisicsFamilien()
            Dim result As Integer = cGDB.StatisicsFamilien

            Assert.That(result, NUnit.Framework.Is.EqualTo(68))

        End Sub

        <Test>
        Public Sub TestVKH_Personen()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.VKH_Personen()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(7))

            Assert.That(dt.Rows(0).Item("Nachname"), NUnit.Framework.Is.EqualTo("Müller"))
            Assert.That(dt.Rows(0).Item("Vorname"), NUnit.Framework.Is.EqualTo("Dieter"))
            Assert.That(dt.Rows(0).Item("Person"), NUnit.Framework.Is.EqualTo("Bräutigam"))

            Assert.That(dt.Rows(1).Item("Nachname"), NUnit.Framework.Is.EqualTo("Müller"))
            Assert.That(dt.Rows(1).Item("Vorname"), NUnit.Framework.Is.EqualTo("Wilhelmine"))
            Assert.That(dt.Rows(1).Item("Person"), NUnit.Framework.Is.EqualTo("Braut"))

            Assert.That(dt.Rows(2).Item("Nachname"), NUnit.Framework.Is.EqualTo("Musterfrau"))
            Assert.That(dt.Rows(2).Item("Vorname"), NUnit.Framework.Is.EqualTo("Petra"))
            Assert.That(dt.Rows(2).Item("Person"), NUnit.Framework.Is.EqualTo("Braut"))

            Assert.That(dt.Rows(3).Item("Nachname"), NUnit.Framework.Is.EqualTo("Mustermann"))
            Assert.That(dt.Rows(3).Item("Vorname"), NUnit.Framework.Is.EqualTo("Bernhard"))
            Assert.That(dt.Rows(3).Item("Person"), NUnit.Framework.Is.EqualTo("Vater Bräutigam"))

            Assert.That(dt.Rows(4).Item("Nachname"), NUnit.Framework.Is.EqualTo("Mustermann"))
            Assert.That(dt.Rows(4).Item("Vorname"), NUnit.Framework.Is.EqualTo("Peter"))
            Assert.That(dt.Rows(4).Item("Person"), NUnit.Framework.Is.EqualTo("Bräutigam"))



            dt = cGDB.VKH_Personen("Bräutigam")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))

            dt = cGDB.VKH_Personen("Vater Bräutigam")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            dt = cGDB.VKH_Personen("Zeuge")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(0))



            dt = cGDB.VKH_Personen("Bräutigam", 1)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))

            dt = cGDB.VKH_Personen("Bräutigam", 3)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))

            dt = cGDB.VKH_Personen("Vater Bräutigam", 1)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            dt = cGDB.VKH_Personen("Vater Bräutigam", 3)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            dt = cGDB.VKH_Personen("Zeuge", 1)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(0))

            dt = cGDB.VKH_Personen("Zeuge", 3)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(0))



            dt = cGDB.VKH_Personen(1)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(7))

            Assert.That(dt.Rows(0).Item("Nachname"), NUnit.Framework.Is.EqualTo("Müller"))
            Assert.That(dt.Rows(0).Item("Vorname"), NUnit.Framework.Is.EqualTo("Dieter"))
            Assert.That(dt.Rows(0).Item("Person"), NUnit.Framework.Is.EqualTo("Bräutigam"))

            Assert.That(dt.Rows(3).Item("Nachname"), NUnit.Framework.Is.EqualTo("Mustermann"))
            Assert.That(dt.Rows(3).Item("Vorname"), NUnit.Framework.Is.EqualTo("Bernhard"))
            Assert.That(dt.Rows(3).Item("Person"), NUnit.Framework.Is.EqualTo("Vater Bräutigam"))

            Assert.That(dt.Rows(4).Item("Nachname"), NUnit.Framework.Is.EqualTo("Mustermann"))
            Assert.That(dt.Rows(4).Item("Vorname"), NUnit.Framework.Is.EqualTo("Peter"))
            Assert.That(dt.Rows(4).Item("Person"), NUnit.Framework.Is.EqualTo("Bräutigam"))

            dt = cGDB.VKH_Personen(3)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(7))

            Assert.That(dt.Rows(0).Item("Nachname"), NUnit.Framework.Is.EqualTo("Müller"))
            Assert.That(dt.Rows(0).Item("Vorname"), NUnit.Framework.Is.EqualTo("Dieter"))
            Assert.That(dt.Rows(0).Item("Person"), NUnit.Framework.Is.EqualTo("Bräutigam"))

            Assert.That(dt.Rows(3).Item("Nachname"), NUnit.Framework.Is.EqualTo("Mustermann"))
            Assert.That(dt.Rows(3).Item("Vorname"), NUnit.Framework.Is.EqualTo("Peter"))
            Assert.That(dt.Rows(3).Item("Person"), NUnit.Framework.Is.EqualTo("Bräutigam"))

            Assert.That(dt.Rows(4).Item("Nachname"), NUnit.Framework.Is.EqualTo("Mustermann"))
            Assert.That(dt.Rows(4).Item("Vorname"), NUnit.Framework.Is.EqualTo("Bernhard"))
            Assert.That(dt.Rows(4).Item("Person"), NUnit.Framework.Is.EqualTo("Vater Bräutigam"))
        End Sub

        <Test>
        Public Sub TestVKH_Locations()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.StatisticsVKHLocations()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo("Musterdorf"))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(1).Item(0), NUnit.Framework.Is.EqualTo("Musterstadt"))
            Assert.That(dt.Rows(1).Item(1), NUnit.Framework.Is.EqualTo(3))

        End Sub

        <Test>
        Public Sub TestVKH_LocationsI()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.StatisticsVKHLocations("Musterdorf", "")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo("Musterstadt"))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo(2))


            dt = cGDB.StatisticsVKHLocations("Musterdorf", "D")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo("Musterstadt"))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo(2))


            dt = cGDB.StatisticsVKHLocations("Musterdorf", "E")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(0))

            dt = cGDB.StatisticsVKHLocations("Musterstadt", "")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

        End Sub


        <Test>
        Public Sub TestVKH_LocationsExtern()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.StatisticsVKHLocationsExtern("Musterstadt", "")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item(2), NUnit.Framework.Is.EqualTo("Musterdorf"))



            dt = cGDB.StatisticsVKHLocationsExtern("Musterdorf", "D")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))

            Assert.That(dt.Rows(0).Item(2), NUnit.Framework.Is.EqualTo("Musterstadt"))

            dt = cGDB.StatisticsVKHLocationsExtern("Musterdorf", "E")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(0))

            dt = cGDB.StatisticsVKHLocationsExtern("Musterstadt", "")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

        End Sub

        <Test>
        Public Sub TestGetNachname()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetNachname()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(4))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo("Müller"))

            Assert.That(dt.Rows(1).Item(0), NUnit.Framework.Is.EqualTo(3))
            Assert.That(dt.Rows(1).Item(1), NUnit.Framework.Is.EqualTo("Musterfrau"))

        End Sub

        <Test>
        Public Sub TestUpdateNachname()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetNachname()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(4))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo("Müller"))
            Assert.That(dt.Rows(1).Item(0), NUnit.Framework.Is.EqualTo(3))
            Assert.That(dt.Rows(1).Item(1), NUnit.Framework.Is.EqualTo("Musterfrau"))

            Dim Nachname As String = "Doe"
            Dim ID As Integer = 1

            cGDB.UpdateNachname(ID, Nachname)

            dt = cGDB.GetNachname()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(4))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo(ID))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo(Nachname))
            Assert.That(dt.Rows(1).Item(0), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(1).Item(1), NUnit.Framework.Is.EqualTo("Müller"))
        End Sub

        <Test>
        Public Sub TestCleanDate()
            Dim dateString As String = "12.01.1900"
            Dim result As Boolean = cGDB.CleanDate(dateString)

            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("12.01.1900"))

            dateString = "12-01-1900"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("12.01.1900"))

            dateString = "12,01,1900"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("12.01.1900"))

            dateString = "12;01;1900"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("12.01.1900"))

            dateString = "12_01_1900"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("12.01.1900"))

            dateString = " 12. 01 . 1900 "
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("12.01.1900"))

            dateString = "2.1.1900"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("02.01.1900"))

            dateString = "29.2.1904"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("29.02.1904"))

            dateString = "29.2.1902"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(False))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("29.2.1902"))

            dateString = "29-2-1902"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(False))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("29.2.1902"))

            dateString = ""
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo(""))

            dateString = "   "
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo(""))


            dateString = "hallo"
            result = cGDB.CleanDate(dateString)
            Assert.That(result, NUnit.Framework.Is.EqualTo(False))
            Assert.That(dateString, NUnit.Framework.Is.EqualTo("hallo"))

        End Sub

        <Test>
        Public Sub TestVKH_Orte()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.VKH_Orte()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo("1901/001"))
            Assert.That(dt.Rows(0).Item(2), NUnit.Framework.Is.EqualTo("Musterdorf"))


            dt = cGDB.VKH_Orte("Musterstadt")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))


            dt = cGDB.VKH_Orte("Musterdorf")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))


            dt = cGDB.VKH_Orte("", "D")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo("1900/1"))
            Assert.That(dt.Rows(0).Item(2), NUnit.Framework.Is.EqualTo("Musterstadt"))


            dt = cGDB.VKH_Orte("", "E")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item(1), NUnit.Framework.Is.EqualTo("1901/001"))
            Assert.That(dt.Rows(0).Item(2), NUnit.Framework.Is.EqualTo("Musterdorf"))


            dt = cGDB.VKH_Orte("", "F")

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(0))

        End Sub

        <Test>
        Public Sub TestToTitleCase()
            Dim test As String = "this is a test STRING"
            Dim result As String = cGDB.ToTitleCase(test)

            Assert.That(result, NUnit.Framework.Is.EqualTo("This Is A Test String"))

            test = ""
            result = cGDB.ToTitleCase(test)

            Assert.That(result, NUnit.Framework.Is.EqualTo(""))

            test = " "
            result = cGDB.ToTitleCase(test)

            Assert.That(result, NUnit.Framework.Is.EqualTo(""))

            test = " new  TEST  string "
            result = cGDB.ToTitleCase(test)

            Assert.That(result, NUnit.Framework.Is.EqualTo("New Test String"))
        End Sub

        <Test>
        Public Sub TestGetVKH_Books()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetVKH_Books()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))

            Assert.That(dt.Rows(0).Item(0), NUnit.Framework.Is.EqualTo("D"))
            Assert.That(dt.Rows(1).Item(0), NUnit.Framework.Is.EqualTo("E"))

        End Sub

        <Test>
        Public Sub TestVKH_ReportData()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.VKH_ReportData()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))


            dt = cGDB.VKH_ReportData(False)
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))

            dt = cGDB.VKH_ReportData(True)
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item(3), NUnit.Framework.Is.EqualTo("1900/2"))


            dt = cGDB.VKH_ReportData("D")
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))

            dt = cGDB.VKH_ReportData("E")
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            dt = cGDB.VKH_ReportData("F")
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(0))


            dt = cGDB.VKH_ReportData(False, "D")
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))

            dt = cGDB.VKH_ReportData(True, "D")
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item(3), NUnit.Framework.Is.EqualTo("1900/2"))


        End Sub

        <Test>
        Public Sub TestGetVKH_Table()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetVKH_Table()

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))

            Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/1"))

            Assert.That(dt.Rows(1).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            Assert.That(dt.Rows(1).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(1).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/2"))

            Assert.That(dt.Rows(2).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("E"))
            Assert.That(dt.Rows(2).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(2).Item("NR_H"), NUnit.Framework.Is.EqualTo("1901/001"))
        End Sub

        <Test>
        Public Sub TestGetVKH_TableEntry()
            Dim DBFileT As String = testFolder & "\TestVK.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetVKH_TableEntry(2)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/1"))


            dt = cGDB.GetVKH_TableEntry(3)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/2"))

        End Sub

        <Test>
        Public Sub TestGetGedcomPersonenEreignisse()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetGedcomPersonenEreignisse

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(214))

            'Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            'Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            'Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/1"))


            'dt = cGDB.GetVKH_TableEntry(3)

            'Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            'Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            'Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            'Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/2"))

        End Sub

        <Test>
        Public Sub TestGetGedcomPersonFamilie()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetGedcomPersonFamilie(2)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item("tblFamilieID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(0).Item("tblPersonIDV"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(0).Item("tblPersonIDM"), NUnit.Framework.Is.EqualTo(7))
        End Sub

        <Test>
        Public Sub TestGetGedcomFamilie()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetGedcomFamilie

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(74))

            'Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            'Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            'Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/1"))


            'dt = cGDB.GetVKH_TableEntry(3)

            'Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            'Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            'Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            'Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/2"))

        End Sub

        <Test>
        Public Sub TestGetQuellen()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetQuellen

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))


            Assert.That(dt.Rows(0).Item("Quelle"), NUnit.Framework.Is.EqualTo("Bonn Münster"))
            Assert.That(dt.Rows(0).Item("QuelleKurz"), NUnit.Framework.Is.EqualTo("BN Münster"))
            Assert.That(dt.Rows(1).Item("Quelle"), NUnit.Framework.Is.EqualTo("Bonn St. Remigius"))
            Assert.That(dt.Rows(1).Item("QuelleKurz"), NUnit.Framework.Is.EqualTo("BN Remigius"))

            'dt = cGDB.GetVKH_TableEntry(3)

            'Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            'Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            'Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            'Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/2"))

        End Sub

        <Test>
        Public Sub TestWorkQuellen()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetQuellen

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))


            Dim ID As Integer = cGDB.SetQuelle("Koblenz St. Kastor", "KO Kastor", "Hinweis")

            dt = cGDB.GetQuellen

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))

            dt = cGDB.GetQuelleByID(ID)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item("Quelle"), NUnit.Framework.Is.EqualTo("Koblenz St. Kastor"))
            Assert.That(dt.Rows(0).Item("QuelleKurz"), NUnit.Framework.Is.EqualTo("KO Kastor"))
            Assert.That(dt.Rows(0).Item("QuelleBeschreibung"), NUnit.Framework.Is.EqualTo("Hinweis"))

            Dim check As Boolean = cGDB.UpdateQuelle(ID, "Koblenz St. Kastor Updated", "KO Kastor Updated", "Hinweis Updated")
            Assert.That(check, NUnit.Framework.Is.EqualTo(True))

            dt = cGDB.GetQuelleByID(ID)

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            Assert.That(dt.Rows(0).Item("Quelle"), NUnit.Framework.Is.EqualTo("Koblenz St. Kastor Updated"))
            Assert.That(dt.Rows(0).Item("QuelleKurz"), NUnit.Framework.Is.EqualTo("KO Kastor Updated"))
            Assert.That(dt.Rows(0).Item("QuelleBeschreibung"), NUnit.Framework.Is.EqualTo("Hinweis Updated"))


            check = cGDB.DeleteQuelle(2)

            Assert.That(check, NUnit.Framework.Is.EqualTo(True))

            dt = cGDB.GetQuellen

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))

            Assert.That(dt.Rows(0).Item("Quelle"), NUnit.Framework.Is.EqualTo("Bonn St. Remigius"))
            Assert.That(dt.Rows(0).Item("QuelleKurz"), NUnit.Framework.Is.EqualTo("BN Remigius"))

        End Sub


        <Test>
        Public Sub TestGetQuellZitate()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetQuellZitate

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(5))


            Assert.That(dt.Rows(0).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(4))
            Assert.That(dt.Rows(0).Item("Jahr"), NUnit.Framework.Is.EqualTo(1767))
            Assert.That(dt.Rows(0).Item("Anzahl"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(1).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(1).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(6))
            Assert.That(dt.Rows(1).Item("Jahr"), NUnit.Framework.Is.EqualTo(1769))
            Assert.That(dt.Rows(1).Item("Anzahl"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(4).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(4).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(4).Item("Jahr"), NUnit.Framework.Is.EqualTo(1786))
            Assert.That(dt.Rows(4).Item("Anzahl"), NUnit.Framework.Is.EqualTo(0))

        End Sub

        <Test>
        Public Sub TestGetQuellZitateF()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetQuellZitateF()
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(5))


            Assert.That(dt.Rows(0).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(4))
            Assert.That(dt.Rows(0).Item("Jahr"), NUnit.Framework.Is.EqualTo(1767))
            Assert.That(dt.Rows(0).Item("Anzahl"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(1).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(1).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(6))
            Assert.That(dt.Rows(1).Item("Jahr"), NUnit.Framework.Is.EqualTo(1769))
            Assert.That(dt.Rows(1).Item("Anzahl"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(4).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(4).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(4).Item("Jahr"), NUnit.Framework.Is.EqualTo(1786))
            Assert.That(dt.Rows(4).Item("Anzahl"), NUnit.Framework.Is.EqualTo(0))

            Dim filter As New Dictionary(Of String, Object) From {
                    {"tblEreignisArtID", 2}
                }

            dt = cGDB.GetQuellZitateF(filter)
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(3))

            filter = New Dictionary(Of String, Object) From {
                {"Jahr", 1769}
            }

            dt = cGDB.GetQuellZitateF(filter)
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(2))
            filter = New Dictionary(Of String, Object) From {
                {"tblEreignisArtID", 6},
                {"Jahr", 1769}
            }

            dt = cGDB.GetQuellZitateF(filter)
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

        End Sub

        <Test>
        Public Sub TestWorkQuellZitate()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetQuellZitate

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(5))

            Dim ID As Integer = cGDB.SetQuellZitat(1, 2, 1790, "3", "1", "7", New Date(1790, 1, 1), "URL", "URLB", "ZitatB")

            dt = cGDB.GetQuellZitate

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(6))
            Assert.That(dt.Rows(5).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(5).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(5).Item("Jahr"), NUnit.Framework.Is.EqualTo(1790))
            Assert.That(dt.Rows(5).Item("Seite"), NUnit.Framework.Is.EqualTo("3"))
            Assert.That(dt.Rows(5).Item("Bd"), NUnit.Framework.Is.EqualTo("1"))
            Assert.That(dt.Rows(5).Item("Datum"), NUnit.Framework.Is.EqualTo(New Date(1790, 1, 1)))
            Assert.That(dt.Rows(5).Item("InternetAdresse"), NUnit.Framework.Is.EqualTo("URL"))
            Assert.That(dt.Rows(5).Item("URLBeschreibung"), NUnit.Framework.Is.EqualTo("URLB"))
            Assert.That(dt.Rows(5).Item("ZitatBeschreibung"), NUnit.Framework.Is.EqualTo("ZitatB"))
            Assert.That(dt.Rows(5).Item("Anzahl"), NUnit.Framework.Is.EqualTo(0))

            Dim check As Boolean = cGDB.UpdateQuellZitat(ID, 1, 2, 1791, "4", "2", "8", New Date(1791, 1, 1), "URL2", "URLB2", "ZitatB2")
            Assert.That(check, NUnit.Framework.Is.EqualTo(True))
            dt = cGDB.GetQuellZitate
            Assert.That(dt.Rows(5).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(5).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(5).Item("Jahr"), NUnit.Framework.Is.EqualTo(1791))
            Assert.That(dt.Rows(5).Item("Seite"), NUnit.Framework.Is.EqualTo("4"))
            Assert.That(dt.Rows(5).Item("Bd"), NUnit.Framework.Is.EqualTo("2"))
            Assert.That(dt.Rows(5).Item("Datum"), NUnit.Framework.Is.EqualTo(New Date(1791, 1, 1)))
            Assert.That(dt.Rows(5).Item("InternetAdresse"), NUnit.Framework.Is.EqualTo("URL2"))
            Assert.That(dt.Rows(5).Item("URLBeschreibung"), NUnit.Framework.Is.EqualTo("URLB2"))
            Assert.That(dt.Rows(5).Item("ZitatBeschreibung"), NUnit.Framework.Is.EqualTo("ZitatB2"))

            check = cGDB.DeleteQuellZitat(ID)
            Assert.That(check, NUnit.Framework.Is.EqualTo(True))
            dt = cGDB.GetQuellZitate
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(5))

            ID = cGDB.SetQuellZitat(1, 2, "a", "3", "1", "7", Nothing, "URL", "URLB", "ZitatB")

            dt = cGDB.GetQuellZitate

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(6))
            Assert.That(dt.Rows(0).Item("tblQuellZitatID"), NUnit.Framework.Is.EqualTo(ID))
            Assert.That(dt.Rows(0).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(0).Item("Jahr"), NUnit.Framework.Is.EqualTo(DBNull.Value))
            Assert.That(dt.Rows(0).Item("Seite"), NUnit.Framework.Is.EqualTo("3"))
            Assert.That(dt.Rows(0).Item("Bd"), NUnit.Framework.Is.EqualTo("1"))
            Assert.That(dt.Rows(0).Item("Datum"), NUnit.Framework.Is.EqualTo(DBNull.Value))
            Assert.That(dt.Rows(0).Item("InternetAdresse"), NUnit.Framework.Is.EqualTo("URL"))
            Assert.That(dt.Rows(0).Item("URLBeschreibung"), NUnit.Framework.Is.EqualTo("URLB"))
            Assert.That(dt.Rows(0).Item("ZitatBeschreibung"), NUnit.Framework.Is.EqualTo("ZitatB"))

            check = cGDB.UpdateQuellZitat(ID, 1, 2, "b", "4", "2", "8", Nothing, "URL2", "URLB2", "ZitatB2")
            Assert.That(check, NUnit.Framework.Is.EqualTo(True))
            dt = cGDB.GetQuellZitate
            Assert.That(dt.Rows(0).Item("tblQuellZitatID"), NUnit.Framework.Is.EqualTo(ID))
            Assert.That(dt.Rows(0).Item("tblQuelleID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("tblEreignisArtID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(0).Item("Jahr"), NUnit.Framework.Is.EqualTo(DBNull.Value))
            Assert.That(dt.Rows(0).Item("Seite"), NUnit.Framework.Is.EqualTo("4"))
            Assert.That(dt.Rows(0).Item("Bd"), NUnit.Framework.Is.EqualTo("2"))
            Assert.That(dt.Rows(0).Item("Datum"), NUnit.Framework.Is.EqualTo(DBNull.Value))
            Assert.That(dt.Rows(0).Item("InternetAdresse"), NUnit.Framework.Is.EqualTo("URL2"))
            Assert.That(dt.Rows(0).Item("URLBeschreibung"), NUnit.Framework.Is.EqualTo("URLB2"))
            Assert.That(dt.Rows(0).Item("ZitatBeschreibung"), NUnit.Framework.Is.EqualTo("ZitatB2"))

        End Sub


        <Test>
        Public Sub TestGetEreignisZitat()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetEreignisZitat

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(5))


            Assert.That(dt.Rows(0).Item("tblQuellZitatID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("tblEreignisID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("tblPersonID"), NUnit.Framework.Is.EqualTo(1))
            Assert.That(dt.Rows(0).Item("EventTag"), NUnit.Framework.Is.EqualTo("_PROB"))
            Assert.That(dt.Rows(1).Item("tblQuellZitatID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(1).Item("tblEreignisID"), NUnit.Framework.Is.EqualTo(16))
            Assert.That(dt.Rows(1).Item("tblPersonID"), NUnit.Framework.Is.EqualTo(2))
            Assert.That(dt.Rows(1).Item("EventTag"), NUnit.Framework.Is.EqualTo("HUSB"))
            Assert.That(dt.Rows(4).Item("tblQuellZitatID"), NUnit.Framework.Is.EqualTo(4))
            Assert.That(dt.Rows(4).Item("tblEreignisID"), NUnit.Framework.Is.EqualTo(259))
            Assert.That(dt.Rows(4).Item("tblPersonID"), NUnit.Framework.Is.EqualTo(130))
            Assert.That(dt.Rows(4).Item("EventTag"), NUnit.Framework.Is.EqualTo(DBNull.Value))
            'dt = cGDB.GetVKH_TableEntry(3)

            'Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(1))

            'Assert.That(dt.Rows(0).Item("BUCH_H"), NUnit.Framework.Is.EqualTo("D"))
            'Assert.That(dt.Rows(0).Item("SEITE_H"), NUnit.Framework.Is.EqualTo(1))
            'Assert.That(dt.Rows(0).Item("NR_H"), NUnit.Framework.Is.EqualTo("1900/2"))

        End Sub

        <Test>
        Public Sub TestWorkEreignisZitat()
            Dim DBFileT As String = testFolder & "\Beethoven.inoGdb"
            cGDB = New inoGenDLL.ClsGenDB(DBFileT)

            Dim dt As DataTable = cGDB.GetEreignisZitat

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(5))

            Dim ID As Integer = cGDB.SetEreignisZitat(6, 7, 8, "WIFE")

            dt = cGDB.GetEreignisZitat

            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(6))

            Assert.That(dt.Rows(5).Item("tblQuellZitatID"), NUnit.Framework.Is.EqualTo(6))
            Assert.That(dt.Rows(5).Item("tblEreignisID"), NUnit.Framework.Is.EqualTo(7))
            Assert.That(dt.Rows(5).Item("tblPersonID"), NUnit.Framework.Is.EqualTo(8))
            Assert.That(dt.Rows(5).Item("EventTag"), NUnit.Framework.Is.EqualTo("WIFE"))

            Dim check As Boolean = cGDB.UpdateEreignisZitat(ID, 9, 10, 11, "HUSB")

            dt = cGDB.GetEreignisZitat
            Assert.That(dt.Rows(5).Item("tblQuellZitatID"), NUnit.Framework.Is.EqualTo(9))
            Assert.That(dt.Rows(5).Item("tblEreignisID"), NUnit.Framework.Is.EqualTo(10))
            Assert.That(dt.Rows(5).Item("tblPersonID"), NUnit.Framework.Is.EqualTo(11))
            Assert.That(dt.Rows(5).Item("EventTag"), NUnit.Framework.Is.EqualTo("HUSB"))

            check = cGDB.DeleteEreignisZitat(ID)

            dt = cGDB.GetEreignisZitat
            Assert.That(dt.Rows.Count, NUnit.Framework.Is.EqualTo(5))

        End Sub
    End Class
End Namespace
