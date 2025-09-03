Imports System.Reflection

Public Class AboutWindow

    Public Sub New()
        InitializeComponent()

        Dim version As String = Assembly.GetExecutingAssembly().GetName().Version.ToString()
        txtVersion.Text = "Version: " & version
    End Sub

    Private Sub Ok_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub Legal_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New Legal()
        win.ShowDialog()
    End Sub
End Class
