Imports System.IO
Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen
    Public Class TestClsKirchenjahr
        Private cKJ As inoGenDLL.ClsKirchenjahr


        <SetUp>
        Public Sub Setup()
            cKJ = New inoGenDLL.ClsKirchenjahr
        End Sub

        <TearDown>
        Public Sub TearDown()

        End Sub

        <Test>
        Public Sub TestGetEasterSunday()
            Dim year As Integer = 2025

            Dim easter As Date = cKJ.GetEasterSunday(year)

            Assert.That(easter, NUnit.Framework.Is.EqualTo(#4/20/2025#))

            year = 2026
            easter = cKJ.GetEasterSunday(year)

            Assert.That(easter, NUnit.Framework.Is.EqualTo(#4/5/2026#))
        End Sub


        <Test>
        Public Sub TestGetSundayAfterEpiphany()
            Dim year As Integer = 2025

            Dim sunday As Date = cKJ.GetSundayAfterEpiphany(year, 1)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#1/12/2025#))

            sunday = cKJ.GetSundayAfterEpiphany(year, 2)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#1/19/2025#))

            sunday = cKJ.GetSundayAfterEpiphany(year, 3)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#1/26/2025#))

            year = 2026
            sunday = cKJ.GetSundayAfterEpiphany(year, 1)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#1/11/2026#))
        End Sub

        <Test>
        Public Sub TestGetSundayAroundEastery()
            Dim year As Integer = 2025

            Dim sunday As Date = cKJ.GetSundayAroundEaster(year, 1)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#4/27/2025#))

            sunday = cKJ.GetSundayAroundEaster(year, 2)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#5/4/2025#))

            sunday = cKJ.GetSundayAroundEaster(year, 3)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#5/11/2025#))

            sunday = cKJ.GetSundayAroundEaster(year, 0)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#4/20/2025#))

            sunday = cKJ.GetSundayAroundEaster(year, -1)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#4/13/2025#))

            sunday = cKJ.GetSundayAroundEaster(year, -2)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#4/6/2025#))


            year = 2026
            sunday = cKJ.GetSundayAroundEaster(year, 1)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#4/12/2026#))
        End Sub

        <Test>
        Public Sub TestGetLastSundayAfterTrinity()
            Dim year As Integer = 2025

            Dim sunday As Date = cKJ.GetLastSundayAfterTrinity(year)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#11/23/2025#))

            year = 2026
            sunday = cKJ.GetLastSundayAfterTrinity(year)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#11/22/2026#))
        End Sub


        <Test>
        Public Sub TestGetAdventSunday()
            Dim year As Integer = 2025

            Dim sunday As Date = cKJ.GetAdventSunday(year, 1)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#11/30/2025#))

            sunday = cKJ.GetAdventSunday(year, 2)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#12/7/2025#))

            sunday = cKJ.GetAdventSunday(year, 3)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#12/14/2025#))

            sunday = cKJ.GetAdventSunday(year, 4)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#12/21/2025#))

            year = 2026
            sunday = cKJ.GetAdventSunday(year, 4)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#12/20/2026#))
        End Sub


        <Test>
        Public Sub TestGetTrinitySunday()
            Dim year As Integer = 2025

            Dim sunday As Date = cKJ.GetTrinitySunday(year)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#6/15/2025#))

            year = 2026
            sunday = cKJ.GetTrinitySunday(year)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#5/31/2026#))
        End Sub

        <Test>
        Public Sub TestGetSundayAfterTrinity()
            Dim year As Integer = 2025

            Dim sunday As Date = cKJ.GetSundayAfterTrinity(year, 1)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#6/22/2025#))

            sunday = cKJ.GetSundayAfterTrinity(year, 2)
            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#6/29/2025#))

            year = 2026
            sunday = cKJ.GetSundayAfterTrinity(year, 1)

            Assert.That(sunday, NUnit.Framework.Is.EqualTo(#6/7/2026#))
        End Sub
    End Class
End Namespace