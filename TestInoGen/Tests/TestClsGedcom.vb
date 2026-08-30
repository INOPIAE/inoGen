Imports System.Data
Imports System.IO
Imports inoGenDLL
Imports Microsoft.ApplicationInsights.MetricDimensionNames.TelemetryContext
Imports NUnit.Framework

Namespace TestInoGen
    Public Class TestClsGedcom
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private testPathSQL As String = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\TestInoGen\bin", "")

        Private DBFile As String = testPath & "\Beethoven.inoGdb"

        Private cGC As ClsGedcom
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
        Public Sub TestBasis()
            cGC = New ClsGedcom()
            Dim filePath As String = testFolder & "\test_mein_stammbaum.ged"
            cGC.Basis(filePath)

        End Sub

        <Test>
        Public Sub TestAddIndividual()
            Dim GedcomFileT As String = testFolder & "\testI.ged"

            Try
                cGC = New ClsGedcom(GedcomFileT, DBFile)

                Dim person1 As New Dictionary(Of String, Object) From {
                    {"IndividualId", "1"},
                    {"Name", "Max /Mustermann/"},
                    {"GivenName", "Max"},
                    {"Surname", "Mustermann"},
                    {"Sex", "M"},
                    {"BirthDate", "12.5.1950"},
                    {"BirthPlace", "Berlin, Deutschland"},
                    {"DeathDate", "15.06.2020"},
                    {"DeathPlace", "München, Deutschland"},
                    {"FamilySIds", New List(Of String) From {"1"}},
                    {"FamilyCId", "0"}
                }

                cGC.AddIndividual(person1)
            Finally
                If cGC IsNot Nothing Then
                    cGC.Dispose()
                    cGC = Nothing
                End If
            End Try

            Threading.Thread.Sleep(100)

            ' Datei einlesen und prüfen
            Dim fileContent As String = File.ReadAllText(GedcomFileT)

            ' Assertions - Prüfen ob bestimmte Werte vorhanden sind

            ' NUnit Constraint Model - lesbarere Assertions
            Assert.That(fileContent, Does.Contain("0 @I1@ INDI"), "Individual ID nicht gefunden")
            Assert.That(fileContent, Does.Contain("1 NAME Max /Mustermann/"), "Name nicht gefunden")
            Assert.That(fileContent, Does.Contain("2 GIVN Max"), "Vorname nicht gefunden")
            Assert.That(fileContent, Does.Contain("2 SURN Mustermann"), "Nachname nicht gefunden")
            Assert.That(fileContent, Does.Contain("1 SEX M"), "Geschlecht nicht gefunden")
            Assert.That(fileContent, Does.Contain("1 BIRT"), "Geburtseintrag nicht gefunden")
            Assert.That(fileContent, Does.Contain("2 DATE 12 MAY 1950"), "Geburtsdatum nicht gefunden")
            Assert.That(fileContent, Does.Contain("2 PLAC Berlin, Deutschland"), "Geburtsort nicht gefunden")
            Assert.That(fileContent, Does.Contain("1 DEAT"), "Todeseintrag nicht gefunden")
            Assert.That(fileContent, Does.Contain("2 DATE 15 JUN 2020"), "Todesdatum nicht gefunden")
            Assert.That(fileContent, Does.Contain("2 PLAC München, Deutschland"), "Todesort nicht gefunden")
            Assert.That(fileContent, Does.Contain("1 FAMS @F1@"), "Familien-Verweis nicht gefunden")
            Assert.That(fileContent, Does.Contain("1 FAMC @F0@"), "FamilienKind-Verweis nicht gefunden")
        End Sub

        <Test>
        Public Sub TestCreateGedcom()

            Dim GedcomFileT As String = testFolder & "\testG.ged"
            Dim SubName As String = "John Doe"

            cGC = New ClsGedcom(GedcomFileT, DBFile)

            cGC.AddHeader(SubName)

            Dim person1 As New Dictionary(Of String, Object) From {
                {"IndividualId", "1"},
                {"Name", "Max /Mustermann/"},
                {"GivenName", "Max"},
                {"Surname", "Mustermann"},
                {"Sex", "M"},
                {"BirthDate", "12.5.1950"},
                {"BirthPlace", "Berlin, Deutschland"},
                {"DeathDate", "15.6.2020"},
                {"DeathPlace", "München, Deutschland"},
                {"FamilySIds", New List(Of String) From {"1"}},
                {"FamilyCId", "0"}
            }

            cGC.AddIndividual(person1)

            cGC.AddTrailer()

            ' Datei einlesen und prüfen
            Dim fileContent As String = File.ReadAllText(GedcomFileT)

            ' Assertions - Prüfen ob bestimmte Werte vorhanden sind
            Assert.That(fileContent, Does.Contain("1 SUBM " + SubName), "Submitter nicht gefunden")
        End Sub

        <Test>
        Public Sub TestCreateGedcomDatabase()
            Dim GedcomFileT As String = testFolder & "\testG1.ged"
            Dim SubName As String = "Jane Doe"

            cGC = New ClsGedcom(GedcomFileT, DBFile)

            cGC.AddHeader(SubName)

            cGC.GetPersons()
            cGC.GetFamily()
            cGC.AddTrailer()

            Dim fileContent As String = File.ReadAllText(GedcomFileT)

            ' Assertions - Prüfen ob bestimmte Werte vorhanden sind
            Assert.That(fileContent, Does.Contain("1 SUBM " + SubName), "Submitter nicht gefunden")
        End Sub


        <Test>
        Public Sub TestConvertDateToGedcom()
            cGC = New ClsGedcom()

            Dim testCases As New Dictionary(Of String, String) From {
                {"12.5.1850", "12 MAY 1850"},
                {"3.11.1875", "3 NOV 1875"},
                {"10.1875", "OCT 1875"},
                {"15.6.2020", "15 JUN 2020"},
                {"< 1875", "BEF 1875"},
                {"<1875", "BEF 1875"},
                {"vor 1875", "BEF 1875"},
                {"vor1875", "BEF 1875"},
                {"> 1875", "AFT 1875"},
                {">1875", "AFT 1875"},
                {"nach 3.1875", "AFT MAR 1875"},
                {"nach 03.1875", "AFT MAR 1875"},
                {"um 1875", "ABT 1875"},
                {"um2.1875", "ABT FEB 1875"},
                {"err 1875", "CAL 1875"},
                {"err1.1875", "CAL JAN 1875"},
                {"err 1.4.1875", "CAL 1 APR 1875"}
            }

            Assert.Multiple(Sub()
                                For Each testCase In testCases
                                    Dim result As String = cGC.ConvertDateToGedcom(testCase.Key)
                                    Assert.That(result, NUnit.Framework.Is.EqualTo(testCase.Value),
                                               $"Input: '{testCase.Key}'")
                                Next
                            End Sub)
        End Sub

        <Test>
        Public Sub TestCreateGedcomFile()
            cGC = New ClsGedcom()
            Dim filePath As String = testFolder & "\complete.ged"
            Dim Submitter As String = "James Doe"
            cGC.CreateGedcomFile(filePath, DBFile, Submitter)

            Dim fileContent As String = File.ReadAllText(filePath)

            ' Assertions - Prüfen ob bestimmte Werte vorhanden sind
            Assert.That(fileContent, Does.Contain("1 SUBM " + Submitter), "Submitter nicht gefunden")

        End Sub
    End Class

End Namespace
