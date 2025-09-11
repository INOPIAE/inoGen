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

            Assert.That(cAT.Persons.Count, NUnit.Framework.Is.EqualTo(126))

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

        <Test>
        Public Sub TestCalculateChild()
            cAT.RootPersonID = 1
            cAT.NewList()

            Dim p As clsAhnentafelDaten.PersonData = cAT.Persons(1)
            Dim c As Integer = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(1))

            p = cAT.Persons(2)
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(1))

            p = cAT.Persons(3)
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(1))

            p.Pos = 16
            p.Gen = 5
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(8))

            p.Pos = 17
            p.Gen = 5
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(8))

            p.Pos = 18
            p.Gen = 5
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(9))

            p.Pos = 19
            p.Gen = 5
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(9))

            p.Pos = 32
            p.Gen = 6
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(8))

            p.Pos = 33
            p.Gen = 6
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(8))

            p.Pos = 36
            p.Gen = 6
            c = cAT.CalculateChild(p)

            Assert.That(c, [Is].EqualTo(9))

        End Sub
        <Test>
        Public Sub CalculateChildPosChart()
            cAT.RootPersonID = 1
            cAT.NewList()

            Dim p As clsAhnentafelDaten.PersonData = cAT.Persons(15)
            Dim c As Integer = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(1), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(16)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(2), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(17)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(1), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(18)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(2), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(18)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(2), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            ' 6. Generation

            p = cAT.Persons(29)
            c = cAT.CalculateChildPosChart(p) 'Markus van Beethoven
            Assert.That(c, [Is].EqualTo(3), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(30)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(4), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(31)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(5), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(32)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(6), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")


            p = cAT.Persons(33)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(3), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(34)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(4), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(35)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(5), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(36)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(6), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")


            ' 7. Generation

            p = cAT.Persons(49)
            c = cAT.CalculateChildPosChart(p) 'Hendrik van Beethoven
            Assert.That(c, [Is].EqualTo(7), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(50)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(8), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(51)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(9), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(52)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(10), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(53)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(11), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(54)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(12), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(55)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(13), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(56)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(14), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")




            p = cAT.Persons(57)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(7), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(58)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(8), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(59)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(9), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(60)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(10), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")


            p = cAT.Persons(61) 'Thomas Vogt
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(9), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(62)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(10), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")


            ' 8. Generation
            p = cAT.Persons(71) 'Aert van Beethoven
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(1), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(72)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(2), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(73)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(1), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(74)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(1), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(75)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(2), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            ' 9. Generation
            p = cAT.Persons(84) 'Marck van Beethoven
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(3), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(85)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(4), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(86)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(3), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos} Geburt {p.Geburtsdatum}")
            p = cAT.Persons(87)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(4), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(88)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(5), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(89)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(6), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(90)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(3), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")


            ' 10. Generation
            p = cAT.Persons(94) 'Jan van Beethoven
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(7), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(95)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(7), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos} Geburt {p.Geburtsdatum}")
            p = cAT.Persons(96)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(8), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(97)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(9), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(98)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(10), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(99)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(11), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(100)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(12), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(101)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(13), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(102)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(14), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(103)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(7), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            ' 11. Generation
            p = cAT.Persons(111) 'Theiß Scheitter
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(1), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(112)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(2), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(113)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(1), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(114)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(2), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            ' 12. Generation
            p = cAT.Persons(115) 'Hans Scheitter
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(3), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(116)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(4), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(117)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(5), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(118)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(3), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(119)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(5), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            ' 13. Generation
            p = cAT.Persons(120) 'Peter Scheitter
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(7), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(121)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(9), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(122)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(10), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            p = cAT.Persons(123)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(7), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")
            p = cAT.Persons(124)
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(8), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

            ' 14. Generation
            p = cAT.Persons(125) 'Jeckel Hoffmann
            c = cAT.CalculateChildPosChart(p)
            Assert.That(c, [Is].EqualTo(1), $"Name: {p.Vorname} {p.Nachname} Pos: {p.Pos}")

        End Sub
    End Class
End Namespace