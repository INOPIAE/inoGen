Public Class ClsGenHelper
    Function GetIndexString(text As String) As String
        If text Is Nothing Then text = ""

        text = text.ToUpper.Trim()

        If text.Length >= 4 Then
            Return text.Substring(0, 4)
        Else
            Return text.PadRight(4, "_"c)
        End If
    End Function

    Function NextChar(c As Char) As Char
        If c >= "A"c AndAlso c <= "Z"c Then
            Return ChrW(((AscW(c) - AscW("A"c) + 1) Mod 26) + AscW("A"c))
        ElseIf c >= "a"c AndAlso c <= "z"c Then
            Return ChrW(((AscW(c) - AscW("a"c) + 1) Mod 26) + AscW("a"c))
        Else
            Return c
        End If
    End Function

End Class
