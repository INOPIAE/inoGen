
Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen
    Public Class TestClsAlter
        Private cA As inoGenDLL.ClsAlter

        <SetUp>
        Public Sub Setup()
            cA = New inoGenDLL.ClsAlter
        End Sub

        <TearDown>
        Public Sub TearDown()

        End Sub

        <Test>
        Public Sub TestCalculateBirthday()

            Dim result As String = cA.CalculateBirthday("10.02.1801", "0", "0", "0", "0")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 10.02.1801"))

            result = cA.CalculateBirthday("10.02.1801", "", "", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 10.02.1801"))

            result = cA.CalculateBirthday("10.02.1801", "1", "", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 10.02.1800"))

            result = cA.CalculateBirthday("10.02.1801", "10", "", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 10.02.1791"))

            result = cA.CalculateBirthday("10.02.1801", "", "1", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 10.01.1801"))

            result = cA.CalculateBirthday("10.02.1801", "", "10", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 10.04.1800"))

            result = cA.CalculateBirthday("10.02.1801", "", "", "1", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 03.02.1801"))

            result = cA.CalculateBirthday("10.02.1801", "", "", "10", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 02.12.1800"))

            result = cA.CalculateBirthday("10.02.1801", "", "", "", "1")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 09.02.1801"))

            result = cA.CalculateBirthday("10.02.1801", "", "", "", "10")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 31.01.1801"))

            result = cA.CalculateBirthday("10.02.1801", "5", "", "3", "1")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 19.01.1796"))

            result = cA.CalculateBirthday("10.02.1801", "5", "", "-3", "-1")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 03.03.1796"))

            result = cA.CalculateBirthday("err 10.02.1801", "5", "", "-3", "-1")
            Assert.That(result, NUnit.Framework.Is.EqualTo("Invalid death date"))
        End Sub

        <Test>
        Public Sub TestCalculateBirthdayYear()

            Dim result As String = cA.CalculateBirthday("1801", "0", "0", "0", "0")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 1801"))

            result = cA.CalculateBirthday("1801", "", "", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 1801"))

            result = cA.CalculateBirthday("1801", "5", "", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 1796"))

            result = cA.CalculateBirthday("1801", "", "10", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 1800"))

            result = cA.CalculateBirthday("1801", "", "", "1", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 1800"))

            result = cA.CalculateBirthday("1801", "", "", "", "1")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 1800"))

            result = cA.CalculateBirthday("1801", "5", "4", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 1795"))
        End Sub

        <Test>
        Public Sub TestCalculateBirthdayYearMonth()

            Dim result As String = cA.CalculateBirthday("02.1801", "0", "0", "0", "0")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 02.1801"))

            result = cA.CalculateBirthday("02.1801", "", "", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 02.1801"))

            result = cA.CalculateBirthday("02.1801", "5", "", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 02.1796"))

            result = cA.CalculateBirthday("02.1801", "", "10", "", "")
            Assert.That(result, NUnit.Framework.Is.EqualTo("err. 04.1800"))

        End Sub
    End Class
End Namespace
