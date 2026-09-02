Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen
    Public Class TestClsGenHelper
        Private cGH As inoGenDLL.ClsGenHelper


        <SetUp>
        Public Sub Setup()
            cGH = New inoGenDLL.ClsGenHelper
        End Sub

        <TearDown>
        Public Sub TearDown()

        End Sub

        <Test>
        Public Sub TestGetIndexString()
            Dim test As String = "Peter"
            Dim result As String = cGH.GetIndexString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("PETE"))

            test = " "
            result = cGH.GetIndexString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("____"))

            test = ""
            result = cGH.GetIndexString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("____"))

            test = "Au"
            result = cGH.GetIndexString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("AU__"))


        End Sub

        <Test>
        Public Sub TestNextChar()
            Dim test As String = "A"
            Dim result As String = cGH.NextChar(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("B"))

            test = "Z"
            result = cGH.NextChar(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("A"))

            test = "a"
            result = cGH.NextChar(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("b"))

            test = "z"
            result = cGH.NextChar(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("a"))

            test = "%"
            result = cGH.NextChar(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("%"))

            test = "ß"
            result = cGH.NextChar(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("ß"))

        End Sub

        <Test>
        Public Sub TestCleanupDateString()
            Dim test As String = " 12 , 05 , 2020 "
            Dim result As String = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("12.05.2020"))
            test = " 12 , , 2020 "
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("12.2020"))
            test = " , , "
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(""))
            test = ""
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(""))
            test = "12.05.2020"
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("12.05.2020"))
            test = "12,05.2020"
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("12.05.2020"))
            test = "<12.05.2020"
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("< 12.05.2020"))
            test = "< 12.05.2020"
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("< 12.05.2020"))
            test = "err12.05.2020"
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("err 12.05.2020"))
            test = "err 12.05.2020"
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("err 12.05.2020"))
            test = "err   12.05.2020  p"
            result = cGH.CleanupDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo("err 12.05.2020 p"))
        End Sub

        <Test>
        Public Sub TestIsValidDateString()
            Dim test As String = "12.05.2020"
            Dim result As Boolean = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "12.05"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "12"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "abt 12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "um 12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "vor 12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "nach 12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "ca 12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "< 12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "> 12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "err12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "circa 12.05.2020"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(False))
            test = ""
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(True))
            test = "."
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(False))
            test = "err12.05.2020 p"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(False))
            test = "12.05.2020 p"
            result = cGH.IsValidDateString(test)
            Assert.That(result, NUnit.Framework.Is.EqualTo(False))
        End Sub
    End Class
End Namespace