Imports System.Text
Imports System.Globalization

Public Class ClsSoundex

    ''' <summary>
    ''' Erzeugt einen klassischen Soundex-Code mit 4 Zeichen.
    ''' Beispiel: Müller -> M460
    ''' </summary>
    Public Shared Function Encode(text As String) As String

        If String.IsNullOrWhiteSpace(text) Then
            Return ""
        End If

        ' Großbuchstaben und deutsche Sonderzeichen vereinheitlichen
        Dim value As String = Normalize(text)

        If value.Length = 0 Then
            Return ""
        End If

        ' Erster Buchstabe bleibt erhalten
        Dim firstLetter As Char = value(0)

        ' Buchstaben in Soundex-Ziffern umwandeln
        Dim result As New StringBuilder()
        result.Append(firstLetter)

        Dim previousCode As String = GetCode(firstLetter)

        For i As Integer = 1 To value.Length - 1

            Dim currentChar As Char = value(i)
            Dim currentCode As String = GetCode(currentChar)

            ' Vokale und H/W/Y setzen den Vergleich zurück
            If currentCode = "0" Then
                previousCode = "0"
                Continue For
            End If

            ' Gleiche Codes direkt hintereinander nur einmal übernehmen
            If currentCode <> previousCode Then
                result.Append(currentCode)
            End If

            previousCode = currentCode

            ' Maximal 4 Zeichen
            If result.Length >= 4 Then
                Exit For
            End If

        Next

        ' Mit Nullen auffüllen
        While result.Length < 4
            result.Append("0")
        End While

        Return result.ToString()

    End Function


    ''' <summary>
    ''' Vereinheitlicht deutsche Zeichen und Sonderzeichen.
    ''' </summary>
    Private Shared Function Normalize(text As String) As String

        Dim value As String = text.Trim().ToUpperInvariant()

        value = value.Replace("Ä", "A")
        value = value.Replace("Ö", "O")
        value = value.Replace("Ü", "U")
        value = value.Replace("ß", "S")

        ' Nur A-Z behalten
        Dim result As New StringBuilder()

        For Each c As Char In value
            If c >= "A"c AndAlso c <= "Z"c Then
                result.Append(c)
            End If
        Next

        Return result.ToString()

    End Function


    ''' <summary>
    ''' Liefert die Soundex-Klasse eines Buchstabens.
    ''' </summary>
    Private Shared Function GetCode(c As Char) As String

        Select Case c

            ' Vokale
            Case "A"c, "E"c, "I"c, "O"c, "U"c
                Return "0"

            ' Lippenlaute
            Case "B"c, "F"c, "P"c, "V"c
                Return "1"

            ' Gutturale / Zischlaute
            Case "C"c, "K"c, "G"c, "J"c, "Q"c, "S"c, "X"c, "Z"c
                Return "2"

            ' Dentale
            Case "D"c, "T"c
                Return "3"

            ' Laterale
            Case "L"c
                Return "4"

            ' Nasale
            Case "M"c, "N"c
                Return "5"

            ' R-Laute
            Case "R"c
                Return "6"

            ' H, W, Y werden nicht codiert
            Case "H"c, "W"c, "Y"c
                Return "0"

            Case Else
                Return "0"

        End Select

    End Function

End Class
