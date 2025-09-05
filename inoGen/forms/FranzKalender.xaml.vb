Imports System.Globalization
Imports inoGenDLL

Public Class FranzKalender

    Private cFK As New ClsFranzKalender
    Private cGJ As New ClsGregJulian
    Public Sub New()
        InitializeComponent()

        ' Beispielwerte für alternative Datum-Dropdowns
        For y As Integer = 1 To 15
            cbYear.Items.Add(y)
        Next
        Dim months = [Enum].GetValues(GetType(ClsFranzKalender.Monate)) _
            .Cast(Of ClsFranzKalender.Monate)() _
            .Select(Function(m) New With {.Key = CInt(m), .Value = m.ToString()}) _
            .ToList()

        cbMonth.ItemsSource = months
        cbMonth.SelectedIndex = 0   ' optional: ersten Monat auswählen
        For d As Integer = 1 To 30
            cbDay.Items.Add(d)
        Next
    End Sub

    Private Sub BtnGregConvert_Click(sender As Object, e As RoutedEventArgs)
        If dpGregorian.SelectedDate.HasValue Then
            Dim strF As String = cFK.GregorianToFranz(dpGregorian.SelectedDate.Value)
            Me.fInfo.Text = strF
            If strF.Contains("Datum") Then
                Exit Sub
            End If
            Dim parts As String() = strF.Split(" "c)
            cbDay.SelectedValue = Integer.Parse(parts(0))
            Dim monthName As String = parts(1)
            For Each item In cbMonth.Items
                Dim m = CType(item, Object)
                If m.Value.ToString().Equals(monthName, StringComparison.OrdinalIgnoreCase) Then
                    cbMonth.SelectedItem = item
                    Exit For
                End If
            Next
            cbYear.SelectedValue = Integer.Parse(parts(2))
        End If
    End Sub

    Private Sub BtnGregCopy_Click(sender As Object, e As RoutedEventArgs)
        If dpGregorian.SelectedDate.HasValue Then
            Clipboard.SetText(dpGregorian.SelectedDate.Value.ToShortDateString())
            MessageBox.Show("Datum kopiert!")
        End If
    End Sub

    Private Sub BtnAltConvert_Click(sender As Object, e As RoutedEventArgs)
        If cbYear.SelectedItem IsNot Nothing AndAlso cbMonth.SelectedItem IsNot Nothing AndAlso cbDay.SelectedItem IsNot Nothing Then
            dpGregorian.SelectedDate = cFK.FranzToGregorian(cbYear.Text, cbMonth.Text, cbDay.Text)
        End If
    End Sub

    Private Sub BtnAltCopy_Click(sender As Object, e As RoutedEventArgs)
        If fInfo.Text <> "Ungültiges Datum" Then
            Clipboard.SetText(fInfo.Text)
            MessageBox.Show("Datum kopiert!")
        End If
    End Sub

    Private Sub BtnClose_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub FranzDatum()
        If cbMonth.SelectedItem IsNot Nothing AndAlso cbDay.SelectedItem IsNot Nothing AndAlso cbYear.SelectedItem IsNot Nothing Then
            Dim franzDate As String = $"{cbDay.SelectedItem} {cbMonth.SelectedItem.Value} {cbYear.SelectedItem}"

            fInfo.Text = franzDate
        Else
            fInfo.Text = "Ungültiges Datum"
        End If
    End Sub

    Private Sub cbDay_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbDay.SelectionChanged
        FranzDatum()
    End Sub

    Private Sub cbMonth_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbMonth.SelectionChanged
        FranzDatum()
    End Sub

    Private Sub cbYear_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbYear.SelectionChanged
        FranzDatum()
    End Sub


    Private Sub BtnGtJConvert_Click(sender As Object, e As RoutedEventArgs)
        If dpGregorian.SelectedDate.HasValue Then
            dpJulian.SelectedDate = cGJ.ToJulianDate(dpGregorian.SelectedDate.Value)
        End If
    End Sub

    Private Sub BtnJulianConvert_Click(sender As Object, e As RoutedEventArgs)
        If dpJulian.SelectedDate.HasValue Then
            dpGregorian.SelectedDate = cGJ.ToGregorianDate(dpJulian.SelectedDate.Value)
        End If
    End Sub

    Private Sub BtnJulianCopy_Click(sender As Object, e As RoutedEventArgs)
        If dpJulian.SelectedDate.HasValue Then
            Clipboard.SetText(dpJulian.SelectedDate.Value.ToShortDateString())
            MessageBox.Show("Datum kopiert!")
        End If
    End Sub
End Class
