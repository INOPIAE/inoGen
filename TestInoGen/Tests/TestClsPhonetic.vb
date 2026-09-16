Imports System.Data
Imports System.IO
Imports inoGenDLL
Imports NUnit.Framework

Namespace TestInoGen
    Public Class TestClsPhonetic
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private testPathSQL As String = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\TestInoGen\bin", "")

        Private DBFile As String = testPath & "\Beethoven.inoGdb"

        Private cCP As New ClsPhonetic
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
        Public Sub TestGetNameCPhonetik()
            Assert.That(cCP.GetNameCPhonetik("Beethoven"), NUnit.Framework.Is.EqualTo("10020306"))
            Assert.That(cCP.GetNameCPhonetik("Ludwig"), NUnit.Framework.Is.EqualTo("502304"))
            Assert.That(cCP.GetNameCPhonetik("Ludewig"), NUnit.Framework.Is.EqualTo("5020304"))
            Assert.That(cCP.GetNameCPhonetik("Johann"), NUnit.Framework.Is.EqualTo("66"))
            Assert.That(cCP.GetNameCPhonetik("Johan"), NUnit.Framework.Is.EqualTo("6"))
            Assert.That(cCP.GetNameCPhonetik("Johannes"), NUnit.Framework.Is.EqualTo("6608"))
            Assert.That(cCP.GetNameCPhonetik("Johannis"), NUnit.Framework.Is.EqualTo("6608"))
        End Sub

        <Test>
        Public Sub TestGetNameCPhonetikRaw()
            Assert.That(cCP.GetNameCPhonetikRaw("Beethoven"), NUnit.Framework.Is.EqualTo("10020306"))
            Assert.That(cCP.GetNameCPhonetikRaw("Ludwig"), NUnit.Framework.Is.EqualTo("502304"))
            Assert.That(cCP.GetNameCPhonetikRaw("Ludewig"), NUnit.Framework.Is.EqualTo("5020304"))
            Assert.That(cCP.GetNameCPhonetikRaw("Johann"), NUnit.Framework.Is.EqualTo("00066"))
            Assert.That(cCP.GetNameCPhonetikRaw("Johan"), NUnit.Framework.Is.EqualTo("0006"))
            Assert.That(cCP.GetNameCPhonetikRaw("Johannes"), NUnit.Framework.Is.EqualTo("0006608"))
            Assert.That(cCP.GetNameCPhonetikRaw("Johannis"), NUnit.Framework.Is.EqualTo("0006608"))
        End Sub

        <Test>
        Public Sub TestGetNameSoundex()
            Assert.That(cCP.GetNameSoundex("Beethoven"), NUnit.Framework.Is.EqualTo("B315"))
            Assert.That(cCP.GetNameSoundex("Ludwig"), NUnit.Framework.Is.EqualTo("L320"))
            Assert.That(cCP.GetNameSoundex("Ludewig"), NUnit.Framework.Is.EqualTo("L320"))
            Assert.That(cCP.GetNameSoundex("Johann"), NUnit.Framework.Is.EqualTo("J550"))
            Assert.That(cCP.GetNameSoundex("Johan"), NUnit.Framework.Is.EqualTo("J500"))
            Assert.That(cCP.GetNameSoundex("Johannes"), NUnit.Framework.Is.EqualTo("J552"))
            Assert.That(cCP.GetNameSoundex("Johannis"), NUnit.Framework.Is.EqualTo("J552"))
        End Sub

    End Class
End Namespace
