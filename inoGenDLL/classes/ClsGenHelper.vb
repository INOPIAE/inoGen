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

End Class
