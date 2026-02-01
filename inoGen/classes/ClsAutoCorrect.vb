Imports System.Text.RegularExpressions

Public Class ClsAutoCorrect
    Private AutoCorrections As New Dictionary(Of String, String)

    Public Class AutoCorrectEntry
        Public Property ReplaceText As String
        Public Property WithText As String
    End Class

    Public Sub New()
        LoadAutoCorrections()
    End Sub

    Public Sub CheckAutoCorrection(txt As TextBox)

        Dim text = txt.Text
        Dim sortedKeys = AutoCorrections.Keys.OrderByDescending(Function(k) k.Length)

        For Each entry In sortedKeys

            Dim suggestion = AutoCorrections(entry)

            Dim correctPattern = $"\b{Regex.Escape(suggestion)}\b"
            If Regex.IsMatch(text, correctPattern) Then
                Continue For
            End If

            Dim wrongPattern = $"\b{Regex.Escape(entry)}\b"
            If Not Regex.IsMatch(text, wrongPattern, RegexOptions.IgnoreCase) Then
                Continue For
            End If

            Dim result = MessageBox.Show(
            $"Möchtest du '{entry}' durch '{suggestion}' ersetzen?",
            "Autokorrektur",
            MessageBoxButton.YesNoCancel,
            MessageBoxImage.Question)

            If result = MessageBoxResult.Yes Then
                txt.Text = ReplaceIgnoreCase(txt.Text, entry, suggestion)
                text = txt.Text ' Text aktualisieren für weitere Prüfungen

            ElseIf result = MessageBoxResult.Cancel Then
                Exit Sub
            End If

            ' bei No → einfach weiter
        Next
    End Sub

    Private Function ReplaceIgnoreCase(
        input As String,
        search As String,
        replacement As String) As String

        Dim pattern = $"\b{Regex.Escape(search)}\b"

        Return Regex.Replace(
            input,
            pattern,
            Function(m)
                If m.Value.All(AddressOf Char.IsUpper) Then
                    Return replacement.ToUpper()
                ElseIf Char.IsUpper(m.Value(0)) Then
                    Return Char.ToUpper(replacement(0)) & replacement.Substring(1)
                Else
                    Return replacement.ToLower()
                End If
            End Function,
            RegexOptions.IgnoreCase)
    End Function

    Public Sub LoadAutoCorrections()
        AutoCorrections.Clear()

        If My.Settings.AutoCorrection Is Nothing Then Exit Sub

        For Each entry In My.Settings.AutoCorrection
            Dim parts = entry.Split("="c)
            If parts.Length = 2 Then
                AutoCorrections(parts(0)) = parts(1)
            End If
        Next
    End Sub

    Public Sub SaveAutoCorrections()
        Dim sc As New Specialized.StringCollection
        For Each kv In AutoCorrections
            sc.Add($"{kv.Key}={kv.Value}")
        Next
        My.Settings.AutoCorrection = sc
        My.Settings.Save()
    End Sub
End Class
