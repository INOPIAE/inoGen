Imports System.IO
Imports NUnit.Framework
Imports inoGenDLL

Namespace TestInoGen
    Public Class TestClsDatabase
        Private testPath As String = Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\bin", ""), "TestData")
        Private testPathSQL As String = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)).Replace("\TestInoGen\bin", "")

        Private DBFile As String = testPath & "\Beethoven.inoGdb"

        Private cDB As New ClsDatabase(DBFile)
        Private cHelper As New ClsHelper
        Private testFolder As String
        Private currentDBVersion As Long = 3

        <SetUp>
        Public Sub Setup()
            testFolder = cHelper.CreateTestFolder("currenttestdata")
        End Sub

        <TearDown>
        Public Sub TearDown()
            cHelper.DeleteTestFolder(testFolder)
        End Sub

        <Test>
        Public Sub TestCreateDBSteps()
            Dim DBFileT As String = testFolder & "\inoGen.inoGdb"
            Dim DBFileL As String = testFolder & "\inoGen.laccdb"
            Dim cDBT As New ClsDatabase(DBFileT)

            Assert.That(File.Exists(DBFileT), NUnit.Framework.Is.False)


            Dim strResult As String = cDBT.CreateDB()

            Assert.That(File.Exists(DBFileT), NUnit.Framework.Is.True)
            Assert.That(strResult, NUnit.Framework.Is.EqualTo("Database Created Successfully"))


            Dim strSQLFile As String = testPathSQL & "\inoGenDLL\SQL\db.sql"
            strResult = cDBT.FillDatabase(strSQLFile)

            Assert.That(strResult, NUnit.Framework.Is.EqualTo("SQL processed"))


            Dim version As Long = cDBT.ReadDBVersion

            Assert.That(version, NUnit.Framework.Is.EqualTo(currentDBVersion))

        End Sub

        <Test>
        Public Sub TestReadDBVersion()
            Dim version As Long = cDB.ReadDBVersion

            Assert.That(version, NUnit.Framework.Is.EqualTo(currentDBVersion))
        End Sub

        <Test>
        Public Sub TestCheckDBVersion()
            Dim version As Long = cDB.CheckDBVersion

            Assert.That(version, NUnit.Framework.Is.EqualTo(currentDBVersion))
        End Sub

        <Test>
        Public Sub TestUpdateDB()
            Dim DBFileT As String = testFolder & "\inoGen.inoGdb"
            Dim DBFileL As String = testFolder & "\inoGen.laccdb"
            Dim cDBT As New ClsDatabase(DBFileT)

            Assert.That(File.Exists(DBFileT), NUnit.Framework.Is.False)

            Dim strResult As String = cDBT.CreateDB()

            Assert.That(File.Exists(DBFileT), NUnit.Framework.Is.True)
            Assert.That(strResult, NUnit.Framework.Is.EqualTo("Database Created Successfully"))

            Dim strSQLFile As String = testPath & "\SQL\firstDB.sql"
            strResult = cDBT.FillDatabase(strSQLFile)

            Assert.That(strResult, NUnit.Framework.Is.EqualTo("SQL processed"))

            Dim version As Long = cDBT.ReadDBVersion

            Assert.That(version, NUnit.Framework.Is.EqualTo(1))

            version = cDBT.CheckDBVersion

            Assert.That(version, NUnit.Framework.Is.EqualTo(currentDBVersion))
        End Sub

    End Class

End Namespace