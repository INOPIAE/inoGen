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
    End Class
End Namespace