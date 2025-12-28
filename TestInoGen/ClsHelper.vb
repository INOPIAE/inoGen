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

    Public Function GetEmail(settingsFile As String) As String
        If Not File.Exists(settingsFile) Then
            Throw New FileNotFoundException("Settings.txt nicht gefunden")
        End If

        Dim email As String =
            File.ReadAllLines(settingsFile).
                 Select(Function(l) l.Trim()).
                 Where(Function(l) l.StartsWith("email=", StringComparison.OrdinalIgnoreCase)).
                 Select(Function(l) l.Substring("email=".Length)).
                 FirstOrDefault()

        If String.IsNullOrWhiteSpace(email) Then
            Throw New Exception("E-Mail nicht in Settings.txt gefunden")
        End If
        Return email
    End Function
End Class
