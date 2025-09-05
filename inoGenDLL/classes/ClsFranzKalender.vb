Public Class ClsFranzKalender
    Public Enum Monate
        Vendémiaire = 0
        Brumaire = 1
        Frimaire = 2
        Nivôse = 3
        Pluviôse = 4
        Ventôse = 5
        Germinal = 6
        Floréal = 7
        Prairial = 8
        Messidor = 9
        Thermidor = 10
        Fructidor = 11
        Sansculotides = 12
    End Enum

    Public Sondertage As String() = {"Jour de la Vertu", "Jour du Génie", "Jour du Travail", "Jour de l'Opinion", "Jour des Récompenses", "Jour de la Révolution"}

    Public Structure Jahr
        Public Jahr As Integer
        Public Beginn As Date
        Public Ende As Date
        Public Schaltjahr As Boolean
    End Structure

    Public Jahre As New List(Of Jahr)
    Public Sub New()
        InitJahre()
    End Sub

    Private Sub InitJahre()
        Jahre.Add(New Jahr With {.Jahr = 1, .Beginn = New Date(1792, 9, 22), .Ende = New Date(1793, 9, 21), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 2, .Beginn = New Date(1793, 9, 22), .Ende = New Date(1794, 9, 21), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 3, .Beginn = New Date(1794, 9, 22), .Ende = New Date(1795, 9, 22), .Schaltjahr = True})
        Jahre.Add(New Jahr With {.Jahr = 4, .Beginn = New Date(1795, 9, 23), .Ende = New Date(1796, 9, 21), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 5, .Beginn = New Date(1796, 9, 22), .Ende = New Date(1797, 9, 21), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 6, .Beginn = New Date(1797, 9, 22), .Ende = New Date(1798, 9, 21), .Schaltjahr = True})
        Jahre.Add(New Jahr With {.Jahr = 7, .Beginn = New Date(1798, 9, 22), .Ende = New Date(1799, 9, 22), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 8, .Beginn = New Date(1799, 9, 23), .Ende = New Date(1800, 9, 22), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 9, .Beginn = New Date(1800, 9, 23), .Ende = New Date(1801, 9, 22), .Schaltjahr = True})
        Jahre.Add(New Jahr With {.Jahr = 10, .Beginn = New Date(1801, 9, 23), .Ende = New Date(1802, 9, 22), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 11, .Beginn = New Date(1802, 9, 23), .Ende = New Date(1803, 9, 23), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 12, .Beginn = New Date(1803, 9, 24), .Ende = New Date(1804, 9, 22), .Schaltjahr = True})
        Jahre.Add(New Jahr With {.Jahr = 13, .Beginn = New Date(1804, 9, 23), .Ende = New Date(1805, 9, 22), .Schaltjahr = False})
        Jahre.Add(New Jahr With {.Jahr = 14, .Beginn = New Date(1805, 9, 23), .Ende = New Date(1805, 12, 31), .Schaltjahr = False})
    End Sub


    Public Function GregorianToFranz(ByVal gregDate As Date) As String
        Dim franzDate As String = ""
        Dim franzJahr As Jahr = Nothing
        For Each jahr As Jahr In Jahre
            If gregDate >= jahr.Beginn And gregDate <= jahr.Ende Then
                franzJahr = jahr
                Exit For
            End If
        Next
        If franzJahr.Jahr = 0 Then
            Return "Datum außerhalb des Franz. Kalenders"
        End If
        Dim tageSeitBeginn As Integer = (gregDate - franzJahr.Beginn).Days
        Dim monatIndex As Integer = tageSeitBeginn \ 30
        Dim tagImMonat As Integer = (tageSeitBeginn Mod 30) + 1
        Dim tagName As String = ""
        If monatIndex > 11 Then
            ' Ergänzungstage
            monatIndex = 12
            tagImMonat = tageSeitBeginn - (12 * 30) + 1
            If tagImMonat > 5 And Not franzJahr.Schaltjahr Then
                Return "Datum außerhalb des Franz. Kalenders"
            ElseIf tagImMonat > 6 And franzJahr.Schaltjahr Then
                Return "Datum außerhalb des Franz. Kalenders"
            End If
            tagName = String.Format(" ({0})", Sondertage(tagImMonat - 1))
        End If
        franzDate = String.Format("{0} {1} {2}", tagImMonat, [Enum].GetName(GetType(Monate), monatIndex), franzJahr.Jahr) & tagName
        Return franzDate
    End Function

    Public Function FranzToGregorian(ByVal franzDate As String) As Nullable(Of Date)
        Dim franzJahr As Jahr
        Dim jahrNummer As Integer
        Dim parts As String() = franzDate.Split(" "c)
        If Not Integer.TryParse(parts(parts.Length - 1), jahrNummer) Then
            Return Nothing
        End If
        If parts.Length <> 3 Then
            If franzDate.Contains("Jour") Then
                Dim st As Integer
                For st = 1 To 6
                    If franzDate.Contains(Sondertage(st - 1)) Then

                        franzJahr = Jahre.Find(Function(j) j.Jahr = jahrNummer)
                        If franzJahr.Jahr = 0 Then
                            Return Nothing
                        End If
                        Return franzJahr.Beginn.AddDays(360 + st - 1)
                    End If
                Next
                ' Ergänzungstage
                parts = franzDate.Split(New Char() {" "c, "("c}, StringSplitOptions.RemoveEmptyEntries)
                If parts.Length <> 3 Then
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
            Return Nothing
        End If
        Dim tag As Integer
        Dim monat As Monate

        If Not Integer.TryParse(parts(0), tag) Then
            Return Nothing
        End If

        If Not [Enum].TryParse(parts(1), True, monat) Then
            Return Nothing
        End If

        franzJahr = Jahre.Find(Function(j) j.Jahr = jahrNummer)
        If franzJahr.Jahr = 0 Then
            Return Nothing
        End If
        Dim monatIndex As Integer = CType(monat, Integer)
        Dim tageSeitBeginn As Integer = (monatIndex * 30) + (tag - 1)
        If monatIndex = 12 Then
            ' Ergänzungstage
            If (tag > 5 And Not franzJahr.Schaltjahr) Or (tag > 6 And franzJahr.Schaltjahr) Then
                Return Nothing
            End If
            tageSeitBeginn = 12 * 30 + (tag - 1)
        End If
        Dim gregDate As Date = franzJahr.Beginn.AddDays(tageSeitBeginn)
        Return gregDate
    End Function

    Public Function FranzToGregorian(Jahr As Integer, ByVal franzDate As String, Tag As Integer) As Nullable(Of Date)
        Return FranzToGregorian(String.Format("{0} {1} {2}", Tag, franzDate, Jahr))
    End Function
End Class
