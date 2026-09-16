Public Class allgemeinesFenster
    Public Sub New(control As UserControl,
                   Optional windowTitle As String = "inoGen")

        InitializeComponent()

        Me.Title = windowTitle

        contentControl.Content = control

    End Sub
End Class
