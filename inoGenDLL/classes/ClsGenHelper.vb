Imports System.Text.RegularExpressions

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

    Function CleanupDateString(dateString As String) As String
        If String.IsNullOrWhiteSpace(dateString) Then
            Return ""
        End If
        dateString = dateString.Trim().Replace(",", ".")
        dateString = Regex.Replace(dateString, "([A-Za-zÄÖÜäöüß<])(\d)", "$1 $2")
        dateString = Regex.Replace(dateString, "(\d)([A-Za-zÄÖÜäöüß])", "$1 $2")
        dateString = Regex.Replace(dateString, "\s+", " ").Trim()

        Dim parts() As String = dateString.Split("."c)
        Dim cleanedParts As New List(Of String)
        For Each part In parts
            Dim trimmedPart As String = part.Trim()
            If Not String.IsNullOrEmpty(trimmedPart) Then
                cleanedParts.Add(trimmedPart)
            End If
        Next
        Return String.Join(".", cleanedParts)
    End Function

    Function IsValidDateString(dateString As String) As Boolean
        If String.IsNullOrWhiteSpace(dateString) Then
            Return True
        End If
        Dim cleanedDateString As String = CleanupDateString(dateString)
        Dim parts() As String = cleanedDateString.Split(" ")
        If parts.Length < 1 OrElse parts.Length > 2 Then
            Return False
        End If
        If parts.Length = 2 Then
            Select Case parts(0).ToLower()
                Case "abt", "um", "vor", "nach", "ca", "<", ">", "err"
                    '  Return True
                Case Else
                    Return False
            End Select
        End If
        Dim testDate As String
        If parts.Length = 1 Then
            testDate = parts(0)
        Else
            testDate = parts(1)
        End If
        If Regex.IsMatch(testDate, "\d") = False Then
            Return False
        End If
        Return True
    End Function

End Class
