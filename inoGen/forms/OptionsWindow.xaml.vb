Public Class OptionsWindow
    Private Sub Save_Click(sender As Object, e As RoutedEventArgs)
        My.Settings.Email = Me.txtEmail.Text
        My.Settings.Save()
        Close()
    End Sub

    Private Sub Cancel_Click(sender As Object, e As RoutedEventArgs)
        Close()
    End Sub

    Private Sub OptionsWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Me.txtEmail.Text = My.Settings.Email

    End Sub
End Class
