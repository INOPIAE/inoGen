Imports System.Collections.ObjectModel
Imports inoGen.ClsAutoCorrect

Public Class AutoCorrection
    Private Entries As New ObservableCollection(Of ClsAutoCorrect.AutoCorrectEntry)
    Private ReadOnly _main As MainWindow
    Public Sub New(main As MainWindow)
        InitializeComponent()

        _main = main
    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As RoutedEventArgs) Handles BtnDelete.Click
        Dim entry = TryCast(lstAutoCorrect.SelectedItem, AutoCorrectEntry)
        If entry Is Nothing Then Exit Sub

        Entries.Remove(entry)
        ClearFields()
    End Sub

    Private Sub BtnNew_Click(sender As Object, e As RoutedEventArgs) Handles BtnNew.Click
        If txtReplace.Text = "" OrElse txtWith.Text = "" Then Exit Sub

        Entries.Add(New AutoCorrectEntry With {
            .ReplaceText = txtReplace.Text.Trim(),
            .WithText = txtWith.Text.Trim()
        })

        ClearFields()

        SortEntries()
    End Sub

    Private Sub SortEntries()
        Dim sorted = Entries.OrderBy(Function(x) x.ReplaceText).ToList()
        Entries.Clear()
        For Each entry In sorted
            Entries.Add(entry)
        Next
    End Sub

    Private Sub BtnUpdate_Click(sender As Object, e As RoutedEventArgs) Handles BtnUpdate.Click
        Dim entry = TryCast(lstAutoCorrect.SelectedItem, AutoCorrectEntry)
        If entry Is Nothing Then Exit Sub

        entry.ReplaceText = txtReplace.Text.Trim()
        entry.WithText = txtWith.Text.Trim()

        lstAutoCorrect.Items.Refresh()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs) Handles BtnCancel.Click
        Close()
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs) Handles BtnSave.Click

        Dim sc As New System.Collections.Specialized.StringCollection()


        For Each entry In Entries
            sc.Add($"{entry.ReplaceText}={entry.WithText}")
        Next
        My.Settings.AutoCorrection = sc
        My.Settings.Save()

        _main.CAutoCorrect.LoadAutoCorrections()

        Close()
    End Sub

    Private Sub AutoCorrection_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        lstAutoCorrect.ItemsSource = Entries

        ' Beispiel: aus Settings laden
        If My.Settings.AutoCorrection Is Nothing Then Exit Sub
        For Each s In My.Settings.AutoCorrection
            Dim p = s.Split("="c)
            If p.Length = 2 Then
                Entries.Add(New AutoCorrectEntry With {
                    .ReplaceText = p(0),
                    .WithText = p(1)
                })
            End If
        Next
    End Sub
    Private Sub LstAutoCorrect_DoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim entry = TryCast(lstAutoCorrect.SelectedItem, AutoCorrectEntry)
        If entry Is Nothing Then Exit Sub

        txtReplace.Text = entry.ReplaceText
        txtWith.Text = entry.WithText
    End Sub
    Private Sub ClearFields()
        txtReplace.Clear()
        txtWith.Clear()
    End Sub
End Class
