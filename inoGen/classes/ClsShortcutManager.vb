Public Class ClsShortcutManager
    Public Shared Sub Attach(window As Window)

        AddHandler window.PreviewKeyDown,
            AddressOf Window_PreviewKeyDown

    End Sub


    Private Shared Sub Window_PreviewKeyDown(
        sender As Object,
        e As KeyEventArgs)

        '---------------------------------------------
        ' CTRL + D
        '---------------------------------------------

        If Keyboard.Modifiers = ModifierKeys.Control AndAlso
           e.Key = Key.D Then

            ClsFormularClipboard.Copy()

            e.Handled = True

            Return

        End If


        '---------------------------------------------
        ' CTRL + E
        '---------------------------------------------

        If Keyboard.Modifiers = ModifierKeys.Control AndAlso
           e.Key = Key.E Then

            ClsFormularClipboard.Paste()

            e.Handled = True

            Return

        End If

    End Sub
End Class
