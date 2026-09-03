Public Class ClsAlter
    Function CalculateBirthday(deathdate As String, Years As String, Months As String, Weeks As String, Days As String) As String
        Dim DateOfDeath As Date
        Dim Quality As String = ""
        Dim parts() As String = deathdate.Split(".")

        If IsDate(deathdate) And parts.Length = 3 Then
            DateOfDeath = CDate(deathdate)
        ElseIf IsNumeric(deathdate) And parts.Length = 1 Then
            DateOfDeath = DateSerial(deathdate, 1, 1)
            Quality = "y"
        ElseIf deathdate.Contains(".") Then
            Quality = "ym"


            For Each p In parts
                If Not IsNumeric(p) Then
                    Return "Invalid death date"
                End If
            Next
            If parts.Length = 3 Then
                Dim day As Integer = CInt(parts(0))
                Dim month As Integer = CInt(parts(1))
                Dim year As Integer = CInt(parts(2))
                DateOfDeath = New Date(year, month, day)
            ElseIf parts.Length = 2 Then
                Dim month As Integer = CInt(parts(0))
                Dim year As Integer = CInt(parts(1))
                DateOfDeath = New Date(year, month, 1)
            Else
                Return "Invalid death date"
            End If
        Else
            Return "Invalid death date"
        End If

        If IsNumeric(Years) Then
            DateOfDeath = DateOfDeath.AddYears(-CInt(Years))
        End If

        If IsNumeric(Months) Then
            DateOfDeath = DateOfDeath.AddMonths(-CInt(Months))
        End If

        If IsNumeric(Weeks) Then
            DateOfDeath = DateOfDeath.AddDays(-CInt(Weeks) * 7)
        End If

        If IsNumeric(Days) Then
            DateOfDeath = DateOfDeath.AddDays(-CInt(Days))
        End If

        If Quality = "y" Then
            Return "err. " & DateOfDeath.Year.ToString()
        End If

        If Quality = "ym" Then
            Return "err. " & DateOfDeath.Month.ToString("00") & "." & DateOfDeath.Year.ToString()
        End If

        Return "err. " & DateOfDeath
    End Function
End Class
