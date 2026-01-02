Public Class ClsKirchenjahr
    Public Function GetEasterSunday(year As Integer) As Date
        Dim a = year Mod 19
        Dim b = year \ 100
        Dim c = year Mod 100
        Dim d = b \ 4
        Dim e = b Mod 4
        Dim f = (b + 8) \ 25
        Dim g = (b - f + 1) \ 3
        Dim h = (19 * a + b - d - g + 15) Mod 30
        Dim i = c \ 4
        Dim k = c Mod 4
        Dim l = (32 + 2 * e + 2 * i - h - k) Mod 7
        Dim m = (a + 11 * h + 22 * l) \ 451
        Dim month = (h + l - 7 * m + 114) \ 31
        Dim day = ((h + l - 7 * m + 114) Mod 31) + 1

        Return New Date(year, month, day)
    End Function

    Public Function GetTrinitySunday(year As Integer) As Date
        Dim easter = GetEasterSunday(year)
        Return easter.AddDays(56)
    End Function

    Public Function GetSundayAfterTrinity(year As Integer, n As Integer) As Date
        If n < 1 Then Throw New ArgumentException("n muss >= 1 sein")

        Dim trinity = GetTrinitySunday(year)
        Return trinity.AddDays(7 * n)
    End Function

    Public Function GetAdventSunday(year As Integer, number As Integer) As Date
        Dim christmas As New Date(year, 12, 24)

        '4th Advent
        While christmas.DayOfWeek <> DayOfWeek.Sunday
            christmas = christmas.AddDays(-1)
        End While
        Select Case number
            Case 1
                Return christmas.AddDays(-21)
            Case 2
                Return christmas.AddDays(-14)
            Case 3
                Return christmas.AddDays(-7)
            Case 4
                Return christmas
            Case Else
                Throw New ArgumentException("number must be 1 to 4")
        End Select

    End Function

    Public Function GetLastSundayAfterTrinity(year As Integer) As Date
        Dim firstAdvent = GetAdventSunday(year, 1)
        Return firstAdvent.AddDays(-7)
    End Function

    Public Function GetSundayAfterEpiphany(year As Integer, n As Integer) As Date
        Dim epiphany As New Date(year, 1, 6)
        ' N-ter Sonntag nach Epiphanias
        Return epiphany.AddDays((7 - CInt(epiphany.DayOfWeek)) Mod 7 + 7 * (n - 1))
    End Function

    Public Function GetSundayAroundEaster(year As Integer, n As Integer) As Date
        'If n < 1 Then n += 1

        Dim easter = GetEasterSunday(year)
        Return easter.AddDays(7 * n)
    End Function
End Class
