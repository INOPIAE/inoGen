
Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen
    Public Class TestClsGregJulian
        Private cGJ As inoGenDLL.ClsGregJulian

        <SetUp>
        Public Sub Setup()
            cGJ = New inoGenDLL.ClsGregJulian
        End Sub

        <TearDown>
        Public Sub TearDown()

        End Sub

        <Test>
        Public Sub TestToJulianDate()
            Dim gDate As New DateTime(1582, 10, 15) ' greg. Kalenderstart
            Dim jDate As DateTime = cGJ.ToJulianDate(gDate)

            Assert.That(jDate, NUnit.Framework.Is.EqualTo(New DateTime(1582, 10, 5)))

            gDate = New DateTime(1917, 11, 7)
            jDate = cGJ.ToJulianDate(gDate)

            Assert.That(jDate, NUnit.Framework.Is.EqualTo(New DateTime(1917, 10, 25)))
        End Sub

        <Test>
        Public Sub TestToGregorianDate()
            Dim gDate As New DateTime(1582, 10, 5) ' greg. Kalenderstart
            Dim jDate As DateTime = cGJ.ToGregorianDate(gDate)

            Assert.That(jDate, NUnit.Framework.Is.EqualTo(New DateTime(1582, 10, 15)))

            gDate = New DateTime(1917, 10, 25) ' greg. Kalenderstart
            jDate = cGJ.ToGregorianDate(gDate)

            Assert.That(jDate, NUnit.Framework.Is.EqualTo(New DateTime(1917, 11, 7)))
        End Sub
    End Class
End Namespace
