Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen

    Public Class TestClsFranzKalender
        Private cFK As inoGenDLL.ClsFranzKalender


        <SetUp>
        Public Sub Setup()
            cFK = New inoGenDLL.ClsFranzKalender
        End Sub

        <TearDown>
        Public Sub TearDown()

        End Sub

        <Test>
        Public Sub TestGregorianToFranz()
            Dim franzDate As String
            franzDate = cFK.GregorianToFranz(New Date(1792, 9, 22))
            Assert.That(franzDate, NUnit.Framework.Is.EqualTo("1 Vendémiaire 1"))

            franzDate = cFK.GregorianToFranz(New Date(1793, 9, 20))
            Assert.That(franzDate, NUnit.Framework.Is.EqualTo("4 Sansculotides 1 (Jour de l'Opinion)"))

            franzDate = cFK.GregorianToFranz(New Date(1793, 10, 5))
            Assert.That(franzDate, NUnit.Framework.Is.EqualTo("14 Vendémiaire 2"))

            franzDate = cFK.GregorianToFranz(New Date(1795, 1, 15))
            Assert.That(franzDate, NUnit.Framework.Is.EqualTo("26 Nivôse 3"))

            franzDate = cFK.GregorianToFranz(New Date(1792, 9, 21))
            Assert.That(franzDate, NUnit.Framework.Is.EqualTo("Datum außerhalb des Franz. Kalenders"))

            franzDate = cFK.GregorianToFranz(New Date(1806, 1, 1))
            Assert.That(franzDate, NUnit.Framework.Is.EqualTo("Datum außerhalb des Franz. Kalenders"))
        End Sub

        <Test>
        Public Sub TestFranzToGregorian()
            Dim gregDate As Nullable(Of Date)
            gregDate = cFK.FranzToGregorian(1, ClsFranzKalender.Monate.Vendémiaire, 1)
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1792, 9, 22)))
            gregDate = cFK.FranzToGregorian("1 Vendémiaire 1")
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1792, 9, 22)))

            gregDate = cFK.FranzToGregorian(2, ClsFranzKalender.Monate.Vendémiaire, 14)
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1793, 10, 5)))
            gregDate = cFK.FranzToGregorian("14 Vendémiaire 2")
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1793, 10, 5)))

            gregDate = cFK.FranzToGregorian(3, ClsFranzKalender.Monate.Nivôse, 26)
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1795, 1, 15)))
            gregDate = cFK.FranzToGregorian("26 Nivôse 3")
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1795, 1, 15)))

            gregDate = cFK.FranzToGregorian(15, ClsFranzKalender.Monate.Nivôse, 26)
            Assert.That(gregDate, NUnit.Framework.Is.Null)
            gregDate = cFK.FranzToGregorian("26 Nivôse 15")
            Assert.That(gregDate, NUnit.Framework.Is.Null)

            For jahr As Integer = 1 To 14
                gregDate = cFK.FranzToGregorian($"26 Nivôse {jahr}")
                Assert.That(gregDate.HasValue, $"Jahr {jahr} sollte ein gültiges gregorianisches Datum liefern.")
            Next

            Dim invalidYears() As Integer = {0, 15, 16, 100}
            For Each jahr In invalidYears
                gregDate = cFK.FranzToGregorian($"26 Nivôse {jahr}")
                Assert.That(gregDate, [Is].Null, $"Jahr {jahr} sollte Nothing zurückgeben.")
            Next

            gregDate = cFK.FranzToGregorian(1, ClsFranzKalender.Monate.Sansculotides, 4)
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1793, 9, 20)))
            gregDate = cFK.FranzToGregorian("4 Sansculotides 1")
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1793, 9, 20)))
            gregDate = cFK.FranzToGregorian("Jour de l'Opinion 1")
            Assert.That(gregDate, NUnit.Framework.Is.EqualTo(New Date(1793, 9, 20)))
        End Sub
    End Class
End Namespace