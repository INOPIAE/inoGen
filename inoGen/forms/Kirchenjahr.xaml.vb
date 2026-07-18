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

        CmbWeekday.ItemsSource = New List(Of String) From {
            "Sonntag",
            "Montag",
            "Dienstag",
            "Mittwoch",
            "Donnerstag",
            "Freitag",
            "Samstag"
        }
        CmbNamedDay.ItemsSource = New List(Of String) From {
            "Epiphanias",
            "Aschermittwoch",
            "Tag der Darstellung Jesu im Tempel",
            "Tag der Verkündigung Marias",
            "Tag der Heimsuchung Mariä",
            "Michaelistag",
            "Reformationsfest",
            "Buß- und Bettag",
            "Totensonntag"
        }
    End Sub

    Private Sub Calculate()
        If IsNumeric(TxtYear.Text) = False Then Exit Sub
        If RbNamedSunday.IsChecked Then
            Select Case CStr(CmbNamedSunday.SelectedItem)
                Case "Septuagesimae / Circumdederunt"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -9).AddDays(CmbWeekday.SelectedIndex)
                Case "Sexagesimae / Exsurge"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -8).AddDays(CmbWeekday.SelectedIndex)
                Case "Quinquagesimae / Estomihi"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -7).AddDays(CmbWeekday.SelectedIndex)
                Case "Quadragesimae / Invokavit"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -6).AddDays(CmbWeekday.SelectedIndex)
                Case "Reminiszere"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -5).AddDays(CmbWeekday.SelectedIndex)
                Case "Okuli"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -4).AddDays(CmbWeekday.SelectedIndex)
                Case "Lätare"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -3).AddDays(CmbWeekday.SelectedIndex)
                Case "Judika"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -2).AddDays(CmbWeekday.SelectedIndex)
                Case "Palmsonntag"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -1).AddDays(CmbWeekday.SelectedIndex)
                Case "Ostern"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 0).AddDays(CmbWeekday.SelectedIndex)
                Case "Quasimodogeniti"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 1).AddDays(CmbWeekday.SelectedIndex)
                Case "Misericordias Domini"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 2).AddDays(CmbWeekday.SelectedIndex)
                Case "Jubilate"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 3).AddDays(CmbWeekday.SelectedIndex)
                Case "Kantate"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 4).AddDays(CmbWeekday.SelectedIndex)
                Case "Rogate"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 5).AddDays(CmbWeekday.SelectedIndex)
                Case "Exaudi"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 6).AddDays(CmbWeekday.SelectedIndex)
                Case "Pfingsten"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 7).AddDays(CmbWeekday.SelectedIndex)
                Case "Trinitatis"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, 8).AddDays(CmbWeekday.SelectedIndex)
                Case "Letzter Sonntag nach Trinitatis"
                    TxtResult.Text = cKJ.GetLastSundayAfterTrinity(TxtYear.Text).AddDays(CmbWeekday.SelectedIndex)
            End Select
        End If
        If RbNamedDay.IsChecked Then
            Select Case CStr(CmbNamedDay.SelectedItem)
                Case "Epiphanias"
                    TxtResult.Text = DateSerial(TxtYear.Text, 1, 6)
                Case "Aschermittwoch"
                    TxtResult.Text = cKJ.GetSundayAroundEaster(TxtYear.Text, -7).AddDays(3)
                Case "Tag der Darstellung Jesu im Tempel"
                    TxtResult.Text = DateSerial(TxtYear.Text, 2, 2)
                Case "Tag der Verkündigung Marias"
                    TxtResult.Text = DateSerial(TxtYear.Text, 3, 25)
                Case "Johannisfest"
                    TxtResult.Text = DateSerial(TxtYear.Text, 6, 24)
                Case "Tag der Heimsuchung Mariä"
                    TxtResult.Text = DateSerial(TxtYear.Text, 7, 2)
                Case "Michaelistag"
                    TxtResult.Text = DateSerial(TxtYear.Text, 9, 29)
                Case "Reformationsfest"
                    TxtResult.Text = DateSerial(TxtYear.Text, 10, 31)
                Case "Buß- und Bettag"
                    TxtResult.Text = cKJ.GetLastSundayAfterTrinity(TxtYear.Text).AddDays(-4)
                Case "Totensonntag"
                    TxtResult.Text = cKJ.GetLastSundayAfterTrinity(TxtYear.Text)
            End Select
        End If
        If RbAdvent.IsChecked Then
            If IsNumeric(TxtAdvent.Text) Then
                Try
                    TxtResult.Text = cKJ.GetAdventSunday(TxtYear.Text, TxtAdvent.Text).AddDays(CmbWeekday.SelectedIndex)
                Catch ex As ArgumentException
                    TxtResult.Text = "Advent muss 1-4 sein"
                End Try
            End If
        End If
        If RbEpi.IsChecked Then
            If IsNumeric(TxtEpi.Text) Then
                TxtResult.Text = cKJ.GetSundayAfterEpiphany(TxtYear.Text, TxtEpi.Text).AddDays(CmbWeekday.SelectedIndex)
            End If
        End If
        If RbTrinit.IsChecked Then
            If IsNumeric(TxtTrinit.Text) Then
                TxtResult.Text = cKJ.GetSundayAfterTrinity(TxtYear.Text, TxtTrinit.Text).AddDays(CmbWeekday.SelectedIndex)
            End If
        End If
        Clipboard.SetText(TxtResult.Text)
    End Sub

    Private Sub CmbNamedSunday_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CmbNamedSunday.SelectionChanged
        If CmbNamedSunday.SelectedItem IsNot Nothing Then
            RbNamedSunday.IsChecked = True
        End If
        CmbWeekday.SelectedIndex = 0
        Calculate()
    End Sub

    Private Sub CmbNamedDay_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CmbNamedDay.SelectionChanged
        If CmbNamedDay.SelectedItem IsNot Nothing Then
            RbNamedDay.IsChecked = True
        End If
        CmbWeekday.SelectedIndex = 0
        Calculate()
    End Sub
    Private Sub TxtAdvent_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TxtAdvent.TextChanged
        If IsNumeric(TxtAdvent.Text) Then
            RbAdvent.IsChecked = True
        End If
        CmbWeekday.SelectedIndex = 0
        Calculate()
    End Sub

    Private Sub TxtEpi_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TxtEpi.TextChanged
        If IsNumeric(TxtEpi.Text) Then
            RbEpi.IsChecked = True
        End If
        CmbWeekday.SelectedIndex = 0
        Calculate()
    End Sub

    Private Sub TxtTrinit_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TxtTrinit.TextChanged
        If IsNumeric(TxtTrinit.Text) Then
            RbTrinit.IsChecked = True
        End If
        CmbWeekday.SelectedIndex = 0
        Calculate()
    End Sub

    Private Sub CmbWeekday_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles CmbWeekday.SelectionChanged
        Calculate()
    End Sub
End Class
