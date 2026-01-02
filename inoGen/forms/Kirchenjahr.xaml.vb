Imports inoGenDLL

Public Class Kirchenjahr
    Private cKJ As New ClsKirchenjahr

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Private Sub BtnCopy_Click(sender As Object, e As RoutedEventArgs) Handles BtnCopy.Click
        Clipboard.SetText(TxtResult.Text)
        If IsNumeric(TxtYear.Text) Then
            My.Settings.LastKJ = TxtYear.Text
            My.Settings.Save()
        End If
    End Sub

    Private Sub Kirchenjahr_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        If My.Settings.LastKJ = 0 Then
            My.Settings.LastKJ = Now.Year
            My.Settings.Save()
        End If

        TxtYear.Text = My.Settings.LastKJ

        CmbNamedSunday.ItemsSource = New List(Of String) From {
            "Septuagesimae / Circumdederunt",
            "Sexagesimae / Exsurge",
            "Quinquagesimae / Estomihi",
            "Quadragesimae / Invokavit",
            "Reminiszere",
            "Okuli",
            "Lätare",
            "Judika",
            "Palmsonntag",
            "Ostern",
            "Quasimodogeniti",
            "Misericordias Domini",
            "Jubilate",
            "Kantate",
            "Rogate",
            "Exaudi",
            "Pfingsten",
            "Trinitatis",
            "Letzter Sonntag nach Trinitatis"
        }
    End Sub

    Private Sub Calculate()
        If IsNumeric(TxtYear.Text) = False Then Exit Sub
        If RbNamedSunday.IsChecked Then
            Select Case CStr(CmbNamedSunday.SelectedItem)
                Case "Septuagesimae / Circumdederunt"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -9)
                Case "Sexagesimae / Exsurge"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -8)
                Case "Quinquagesimae / Estomihi"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -7)
                Case "Quadragesimae / Invokavit"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -6)
                Case "Reminiszere"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -5)
                Case "Okuli"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -4)
                Case "Lätare"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -3)
                Case "Judika"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -2)
                Case "Palmsonntag"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -1)
                Case "Ostern"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 0)
                Case "Quasimodogeniti"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 1)
                Case "Misericordias Domini"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 2)
                Case "Jubilate"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 3)
                Case "Kantate"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 4)
                Case "Rogate"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 5)
                Case "Exaudi"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 6)
                Case "Pfingsten"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 7)
                Case "Trinitatis"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 8)
                Case "Letzter Sonntag nach Trinitatis"
                    TxtResult.Text = cKJ.GetLastSundayAfterTrinity(TxtYear.Text)
            End Select
        End If
        If RbAdvent.IsChecked Then
            If IsNumeric(TxtAdvent.Text) Then
                TxtResult.Text = cKJ.GetAdventSunday(TxtYear.Text, TxtAdvent.Text)
            End If
        End If
        If RbEpi.IsChecked Then
            If IsNumeric(TxtEpi.Text) Then
                TxtResult.Text = cKJ.GetSundayAfterEpiphany(TxtYear.Text, TxtEpi.Text)
            End If
        End If
        If RbTrinit.IsChecked Then
            If IsNumeric(TxtTrinit.Text) Then
                TxtResult.Text = cKJ.GetSundayAfterTrinity(TxtYear.Text, TxtTrinit.Text)
            End If
        End If
        Clipboard.SetText(TxtResult.Text)
    End Sub

    Private Sub CmbNamedSunday_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CmbNamedSunday.SelectionChanged
        If CmbNamedSunday.SelectedItem IsNot Nothing Then
            RbNamedSunday.IsChecked = True
        End If
        Calculate()
    End Sub

    Private Sub TxtAdvent_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TxtAdvent.TextChanged
        If IsNumeric(TxtAdvent.Text) Then
            RbAdvent.IsChecked = True
        End If
        Calculate()
    End Sub

    Private Sub TxtEpi_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TxtEpi.TextChanged
        If IsNumeric(TxtEpi.Text) Then
            RbEpi.IsChecked = True
        End If
        Calculate()
    End Sub

    Private Sub TxtTrinit_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TxtTrinit.TextChanged
        If IsNumeric(TxtTrinit.Text) Then
            RbTrinit.IsChecked = True
        End If
        Calculate()
    End Sub
End Class
