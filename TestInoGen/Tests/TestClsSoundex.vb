Imports System.Data
Imports System.IO
Imports inoGenDLL
Imports NUnit.Framework

Namespace TestInoGen
    Public Class TestClsSoundex
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private testPathSQL As String = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\TestInoGen\bin", "")

        Private DBFile As String = testPath & "\Beethoven.inoGdb"

        Private cSo As New ClsSoundex
        Private cHelper As New ClsHelper
        Private testFolder As String


        <SetUp>
        Public Sub Setup()
            testFolder = cHelper.CreateTestFolder("currenttestdata")
        End Sub

        <TearDown>
        Public Sub TearDown()
            cHelper.DeleteTestFolder(testFolder)
        End Sub

        <Test>
        Public Sub TestEncode()
            Assert.That(cSo.Encode("Beethoven"), NUnit.Framework.Is.EqualTo("B315"))
            Assert.That(cSo.Encode("Ludwig"), NUnit.Framework.Is.EqualTo("L320"))
            Assert.That(cSo.Encode("Ludewig"), NUnit.Framework.Is.EqualTo("L320"))
            Assert.That(cSo.Encode("Johann"), NUnit.Framework.Is.EqualTo("J500"))
            Assert.That(cSo.Encode("Johan"), NUnit.Framework.Is.EqualTo("J500"))
            Assert.That(cSo.Encode("Johannes"), NUnit.Framework.Is.EqualTo("J520"))
            Assert.That(cSo.Encode("Johannis"), NUnit.Framework.Is.EqualTo("J520"))
            Assert.That(cSo.Encode("Müller"), NUnit.Framework.Is.EqualTo("M460"))
            Assert.That(cSo.Encode("Mueller"), NUnit.Framework.Is.EqualTo("M460"))
            Assert.That(cSo.Encode("Miller"), NUnit.Framework.Is.EqualTo("M460"))
        End Sub
    End Class
End Namespace
