Imports System.IO
Imports NUnit.Framework
Public Class ClsHelper
    Public Function CreateTestFolder(Optional strSubfolder As String = "") As String
        Dim strFolder As String = Path.GetDirectoryName(TestContext.CurrentContext.TestDirectory)

        strFolder &= "\" & strSubfolder
        Directory.CreateDirectory(strFolder)
        Return strFolder
    End Function

    Public Sub DeleteTestFolder(strFolder)
        For Each d In Directory.GetDirectories(strFolder)
            Directory.Delete(d, True)
        Next

        For Each f In Directory.GetFiles(strFolder)
            File.Delete(f)
        Next
    End Sub
End Class
