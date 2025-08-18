'Class Application

Partial Public Class Application
    Inherits System.Windows.Application

    Public Shared ReadOnly Property MyAppFolder As String
        Get
            Dim appDataPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            Dim folder As String = IO.Path.Combine(appDataPath, "inoGen")
            If Not IO.Directory.Exists(folder) Then
                IO.Directory.CreateDirectory(folder)
            End If
            Return folder
        End Get
    End Property
End Class
