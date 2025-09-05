Public Class ClsGregJulian
    ' Gregorianisch -> Julian
    Function ToJulianDate(gregDate As DateTime) As DateTime
        Dim jd As Integer = GregorianToJulianDayNumber(gregDate)
        Return JulianDayNumberToJulianDate(jd)
    End Function

    ' Julian -> Gregorian
    Function ToGregorianDate(julianDate As DateTime) As DateTime
        Dim jd As Integer = JulianToJulianDayNumber(julianDate)
        Return JulianDayNumberToGregorianDate(jd)
    End Function

    ' --- Hilfsfunktionen ---

    ' Gregorianisches Datum → Julian Day Number
    Private Function GregorianToJulianDayNumber(d As DateTime) As Integer
        Dim a As Integer = (14 - d.Month) \ 12
        Dim y As Integer = d.Year + 4800 - a
        Dim m As Integer = d.Month + 12 * a - 3

        Return d.Day + ((153 * m + 2) \ 5) + 365 * y + (y \ 4) - (y \ 100) + (y \ 400) - 32045
    End Function

    ' Julianisches Datum → Julian Day Number
    Private Function JulianToJulianDayNumber(d As DateTime) As Integer
        Dim a As Integer = (14 - d.Month) \ 12
        Dim y As Integer = d.Year + 4800 - a
        Dim m As Integer = d.Month + 12 * a - 3

        Return d.Day + ((153 * m + 2) \ 5) + 365 * y + (y \ 4) - 32083
    End Function

    ' Julian Day Number → Gregorianisches Datum
    Private Function JulianDayNumberToGregorianDate(jd As Integer) As DateTime
        Dim a As Integer = jd + 32044
        Dim b As Integer = (4 * a + 3) \ 146097
        Dim c As Integer = a - (146097 * b) \ 4
        Dim d As Integer = (4 * c + 3) \ 1461
        Dim e As Integer = c - (1461 * d) \ 4
        Dim m As Integer = (5 * e + 2) \ 153

        Dim day As Integer = e - (153 * m + 2) \ 5 + 1
        Dim month As Integer = m + 3 - 12 * (m \ 10)
        Dim year As Integer = 100 * b + d - 4800 + (m \ 10)

        Return New DateTime(year, month, day)
    End Function

    ' Julian Day Number → Julianisches Datum
    Private Function JulianDayNumberToJulianDate(jd As Integer) As DateTime
        Dim b As Integer = 0
        Dim c As Integer = jd + 32082
        Dim d As Integer = (4 * c + 3) \ 1461
        Dim e As Integer = c - (1461 * d) \ 4
        Dim m As Integer = (5 * e + 2) \ 153

        Dim day As Integer = e - (153 * m + 2) \ 5 + 1
        Dim month As Integer = m + 3 - 12 * (m \ 10)
        Dim year As Integer = d - 4800 + (m \ 10)

        Return New DateTime(year, month, day)
    End Function

End Class
