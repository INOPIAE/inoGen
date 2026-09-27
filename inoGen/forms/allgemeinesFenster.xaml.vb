Public Class allgemeinesFenster

    Private _formular As IFormularClipboard
    Public Sub New(control As UserControl,
                   Optional windowTitle As String = "inoGen")

        InitializeComponent()

        Me.Title = windowTitle

        contentControl.Content = control

        If TypeOf control Is IFormularClipboard Then

            _formular =
                DirectCast(control, IFormularClipboard)

            ClsFormularClipboard.SetActiveForm(_formular)

        End If


        '-----------------------------------------
        ' Tastatursteuerung aktivieren
        '-----------------------------------------

        ClsShortcutManager.Attach(Me)

    End Sub
End Class
