Imports System.Collections.ObjectModel
Imports System.ComponentModel

Public Class OptionsWindow
    Private _BerufListe As New ObservableCollection(Of String)
    Private BerufView As ICollectionView

    Public ReadOnly Property BerufListe As ObservableCollection(Of String)
        Get
            Return _BerufListe
        End Get
    End Property

    Private Sub OptionsWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        BerufView = CollectionViewSource.GetDefaultView(BerufListe)
        BerufView.SortDescriptions.Clear()
        BerufView.SortDescriptions.Add(
            New SortDescription("", ListSortDirection.Ascending))
        If Not IsNothing(My.Settings.CurrentWork) Then
            For Each entry In My.Settings.CurrentWork
                If entry.StartsWith("VKHOrt:") Then
                    txtVKHOrt.Text = entry.Replace("VKHOrt:", "")
                    Continue For
                End If
                entry = entry.Replace("Beruf:", "")
                If Not BerufListe.Contains(entry) Then
                    BerufListe.Add(entry)
                End If
            Next
        End If

    End Sub

    Private Sub btnBAdd_Click(sender As Object, e As RoutedEventArgs) Handles btnBAdd.Click
        Dim neuerBeruf = txtBeruf.Text.Trim()

        If String.IsNullOrEmpty(neuerBeruf) Then Exit Sub

        If Not BerufListe.Contains(neuerBeruf) Then
            BerufListe.Add(neuerBeruf)
        End If

        txtBeruf.Clear()
    End Sub

    Private Sub Save_Click(sender As Object, e As RoutedEventArgs)
        My.Settings.Email = Me.txtEmail.Text

        Dim sc As New System.Collections.Specialized.StringCollection()
        For Each beruf As String In BerufListe
            sc.Add(String.Format("Beruf:{0}", beruf))
        Next
        sc.Add(String.Format("VKHOrt:{0}", txtVKHOrt.Text.Trim))

        My.Settings.CurrentWork = sc

        My.Settings.Save()
        Close()
    End Sub

    Private Sub Cancel_Click(sender As Object, e As RoutedEventArgs)
        Close()
    End Sub

    Private Sub OptionsWindow_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Me.txtEmail.Text = My.Settings.Email
        DataContext = Me
        LoadColorSettings()
    End Sub


    Private Sub btnBRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnBRemove.Click
        Dim selectedBeruf As String = TryCast(lstBeruf.SelectedItem, String)

        If selectedBeruf Is Nothing Then Exit Sub

        BerufListe.Remove(selectedBeruf)
    End Sub

    Private Sub rectColor1_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor1.MouseLeftButtonDown
        If ShowColorDialog(rectColor1, My.Settings.Gen71) Then
            My.Settings.Gen71 = GetDrawingColorFromRect(rectColor1)
        End If
    End Sub

    Private Sub rectColor2_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor2.MouseLeftButtonDown
        If ShowColorDialog(rectColor2, My.Settings.Gen72) Then
            My.Settings.Gen72 = GetDrawingColorFromRect(rectColor2)
        End If
    End Sub

    Private Sub rectColor3_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor3.MouseLeftButtonDown
        If ShowColorDialog(rectColor3, My.Settings.Gen73) Then
            My.Settings.Gen73 = GetDrawingColorFromRect(rectColor3)
        End If
    End Sub

    Private Sub rectColor4S_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor4S.MouseLeftButtonDown
        If ShowColorDialog(rectColor4S, My.Settings.Gen74S) Then
            My.Settings.Gen74S = GetDrawingColorFromRect(rectColor4S)
        End If
    End Sub

    Private Sub rectColor4E_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor4E.MouseLeftButtonDown
        If ShowColorDialog(rectColor4E, My.Settings.Gen74E) Then
            My.Settings.Gen74E = GetDrawingColorFromRect(rectColor4E)
        End If
    End Sub

    Private Sub rectColor5S_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor5S.MouseLeftButtonDown
        If ShowColorDialog(rectColor5S, My.Settings.Gen75S) Then
            My.Settings.Gen75S = GetDrawingColorFromRect(rectColor5S)
        End If
    End Sub

    Private Sub rectColor5E_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor5E.MouseLeftButtonDown
        If ShowColorDialog(rectColor5E, My.Settings.Gen75E) Then
            My.Settings.Gen75E = GetDrawingColorFromRect(rectColor5E)
        End If
    End Sub

    Private Sub rectColor6S_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor6S.MouseLeftButtonDown
        If ShowColorDialog(rectColor6S, My.Settings.Gen76S) Then
            My.Settings.Gen76S = GetDrawingColorFromRect(rectColor6S)
        End If
    End Sub

    Private Sub rectColor6E_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor6E.MouseLeftButtonDown
        If ShowColorDialog(rectColor6E, My.Settings.Gen76E) Then
            My.Settings.Gen76E = GetDrawingColorFromRect(rectColor6E)
        End If
    End Sub


    Private Sub rectColor7S_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor7S.MouseLeftButtonDown
        If ShowColorDialog(rectColor7S, My.Settings.Gen77S) Then
            My.Settings.Gen77S = GetDrawingColorFromRect(rectColor7S)
        End If
    End Sub

    Private Sub rectColor7E_MouseLeftButtonDown(sender As Object, e As MouseButtonEventArgs) Handles rectColor7E.MouseLeftButtonDown
        If ShowColorDialog(rectColor7E, My.Settings.Gen77E) Then
            My.Settings.Gen77E = GetDrawingColorFromRect(rectColor7E)
        End If
    End Sub


    Private Function ShowColorDialog(rect As Rectangle, currentColor As System.Drawing.Color) As Boolean
        Dim colorDialog As New System.Windows.Forms.ColorDialog() With {
            .Color = currentColor
        }

        If colorDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim wpfColor = Color.FromArgb(colorDialog.Color.A, colorDialog.Color.R, colorDialog.Color.G, colorDialog.Color.B)
            rect.Fill = New SolidColorBrush(wpfColor)
            Return True
        End If

        Return False
    End Function
    Private Function GetDrawingColorFromRect(rect As Rectangle) As System.Drawing.Color
        Dim brush = TryCast(rect.Fill, SolidColorBrush)
        If brush IsNot Nothing Then
            Dim wpfColor = brush.Color
            Return System.Drawing.Color.FromArgb(wpfColor.A, wpfColor.R, wpfColor.G, wpfColor.B)
        End If
        Return System.Drawing.Color.Black
    End Function

    Private Sub LoadColorSettings()
        rectColor1.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen71))
        rectColor2.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen72))
        rectColor3.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen73))
        rectColor4S.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen74S))
        rectColor4E.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen74E))
        rectColor4S.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen74S))
        rectColor5E.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen75E))
        rectColor5S.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen75S))
        rectColor6E.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen76E))
        rectColor6S.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen76S))
        rectColor7E.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen77E))
        rectColor7S.Fill = New SolidColorBrush(ConvertToWpfColor(My.Settings.Gen77S))
    End Sub

    Private Function ConvertToWpfColor(drawingColor As System.Drawing.Color) As Color
        Return Color.FromArgb(drawingColor.A, drawingColor.R, drawingColor.G, drawingColor.B)
    End Function
End Class
