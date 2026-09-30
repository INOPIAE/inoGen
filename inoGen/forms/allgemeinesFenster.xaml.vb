Public Class allgemeinesFenster
    Private _navigation As INavigationService
    Private _formular As IFormularClipboard

    Public Sub New(control As UserControl, navigation As INavigationService,
                   Optional windowTitle As String = "inoGen")

        InitializeComponent()

        Me.Title = windowTitle

        _navigation = navigation

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

    Public ReadOnly Property Navigation _
    As INavigationService

        Get
            Return _navigation
        End Get

    End Property
End Class
