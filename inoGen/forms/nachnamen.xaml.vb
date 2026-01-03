Imports System.Data
Imports inoGenDLL

Public Class nachnamen
    Private cGDB As New ClsGenDB(My.Settings.DBPath)

    Private dtNachname As DataTable
    Private dvNachname As DataView

    Private Sub TxtSuche_TextChanged(sender As Object, e As TextChangedEventArgs) Handles TxtSuche.TextChanged
        If dvNachname Is Nothing Then Exit Sub

        Dim filter = TxtSuche.Text.Replace("'", "''") ' Schutz
        dvNachname.RowFilter = $"Nachname LIKE '%{filter}%'"
    End Sub

    Private Sub nachnamen_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        UpdateData()
        TxtID.IsEnabled = False
    End Sub

    Private Sub DgNachname_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles DgNachname.MouseDoubleClick
        Dim rowView = TryCast(DgNachname.SelectedItem, DataRowView)
        If rowView Is Nothing Then Exit Sub

        TxtID.Text = rowView("tblNachnameID").ToString()
        TxtNachname.Text = rowView("Nachname").ToString()

    End Sub

    Private Sub BtnSelect_Click(sender As Object, e As RoutedEventArgs)

        Dim rowView = TryCast(DgNachname.SelectedItem, DataRowView)
        If rowView Is Nothing Then
            MessageBox.Show("Bitte Datensatz auswählen")
            Return
        End If

        Dim id As Integer = CInt(rowView("ID"))
        Dim nachname As String = rowView("Nachname").ToString()

    End Sub

    Private Sub UpdateData()
        dtNachname = cGDB.GetNachname
        dvNachname = dtNachname.DefaultView
        DgNachname.ItemsSource = dvNachname
    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)

        cGDB.UpdateNachname(TxtID.Text, TxtNachname.Text)
        UpdateData()

    End Sub
End Class
