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
            For Each beruf In My.Settings.CurrentWork
                If Not BerufListe.Contains(beruf) Then
                    BerufListe.Add(beruf)
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
            sc.Add(beruf)
        Next
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

    End Sub


    Private Sub btnBRemove_Click(sender As Object, e As RoutedEventArgs) Handles btnBRemove.Click
        Dim selectedBeruf As String = TryCast(lstBeruf.SelectedItem, String)

        If selectedBeruf Is Nothing Then Exit Sub

        BerufListe.Remove(selectedBeruf)
    End Sub

End Class
