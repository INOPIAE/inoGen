Imports System.IO
Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen
    Public Class TestClsAhnenTafelDaten
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private DBFile As String = testPath & "\Beethoven.inoGdb"
        Private cAT As inoGenDLL.clsAhnentafelDaten
        Private cHelper As New ClsHelper
        Private testFolder As String


        <SetUp>
        Public Sub Setup()
            testFolder = cHelper.CreateTestFolder("currenttestdata")
            cHelper.DeleteTestFolder(testFolder)
            Dim DBFileTest As String = testFolder & "\Beethoven.inoGdb"
            File.Copy(DBFile, DBFileTest)
            cAT = New inoGenDLL.clsAhnentafelDaten(DBFileTest)
        End Sub

        <TearDown>
        Public Sub TearDown()
            cHelper.DeleteTestFolder(testFolder)
        End Sub

        <Test>
        Public Sub TestFillPerson()
            cAT.RootPersonID = 1
            cAT.NewList()

            Assert.That(cAT.Persons.Count, NUnit.Framework.Is.EqualTo(7))

        End Sub

        <Test>
        Public Sub TestAusgabePersDetails()
            cAT.RootPersonID = 1
            cAT.NewList()
            Dim pd = cAT.GetPersonByID(1)
            Dim filePath As String = Path.Combine(testFolder, "test.md")
            Using writer As New StreamWriter(filePath, False, System.Text.Encoding.UTF8)
                cAT.AusgabePersDetails(writer, pd)
            End Using

            Dim lines As String() = File.ReadAllLines(filePath)

            Dim expected As String() = {
                "~ 17.12.1770 Bonn",
                "† 26.03.1827 Wien",
                "⚰ 29.03.1827 Wien"
            }

            For Each line In expected
                Assert.That(lines, Does.Contain(line), $"Zeile fehlt: {line}")
            Next
        End Sub

        <Test>
        Public Sub TestOutputVorname()
            Dim Vorname As String = "Karl *Walter Dieter"

            Dim result As String = cAT.OutputVorname(Vorname)

            Assert.That(result, Does.Contain("*Walter*"))


            Vorname = "Martin Heinrich August Hermann *Erich"

            result = cAT.OutputVorname(Vorname)

            Assert.That(result, Does.Contain("*Erich*"))
        End Sub

        <Test>
        Public Sub TestToRoman()
            Dim testValue As Integer = 1999

            Dim result As String = cAT.ToRoman(testValue)

            Assert.That(result, [Is].EqualTo("MCMXCIX"))

            testValue = 1

            result = cAT.ToRoman(testValue)

            Assert.That(result, [Is].EqualTo("I"))
        End Sub

        <Test>
        Public Sub TestToRoman_Exception()

            Dim ex = Assert.Throws(Of ArgumentOutOfRangeException)(Sub() cAT.ToRoman(4000))
            Assert.That(ex.Message, Does.Contain("1 - 3999"))

            ex = Assert.Throws(Of ArgumentOutOfRangeException)(Sub() cAT.ToRoman(0))
            Assert.That(ex.Message, Does.Contain("1 - 3999"))

            ex = Assert.Throws(Of ArgumentOutOfRangeException)(Sub() cAT.ToRoman(5000))
            Assert.That(ex.Message, Does.Contain("1 - 3999"))

            ex = Assert.Throws(Of ArgumentOutOfRangeException)(Sub() cAT.ToRoman(-1))
            Assert.That(ex.Message, Does.Contain("1 - 3999"))
        End Sub
    End Class
End Namespace