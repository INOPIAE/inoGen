Imports System.Data
Imports System.Windows.Forms
Imports inoGenDLL

Public Class VKH_Ausgabe
    Private connectionString As String =
        String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)

    Private cGeoD As New ClsGeoDaten(My.Settings.DBPath)
    Private cGen As New inoGenDLL.ClsGenDB(My.Settings.DBPath)
    Private cFH As New ClsFileHandling
    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnOSMMap_Click(sender As Object, e As RoutedEventArgs) Handles btnOSMMap.Click
        cGeoD.ErstelleLocationList(txtOrt.Text.Trim, cmbBuch.Text)
        Dim win As OSMKarte
        If txtOrt.Text.Trim = vbNullString Then
            win = New OSMKarte(cGeoD.LocationList)
        Else
            Dim Grundort As New ClsOSMKarte.marker
            Grundort = cGeoD.GetGeoData(txtOrt.Text.Trim, 10)
            win = New OSMKarte(cGeoD.LocationList, Grundort)
        End If

        win.Show()
    End Sub

    Private Sub btnPDF_Click(sender As Object, e As RoutedEventArgs) Handles btnPDF.Click
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "PDF-Dateien (*.pdf)|*.pdf"
        saveFileDialog.Title = "PDF speichern"
        saveFileDialog.DefaultExt = "pdf"
        saveFileDialog.AddExtension = True

        ' Dialog anzeigen
        If saveFileDialog.ShowDialog() = Forms.DialogResult.OK Then
            Dim strFileName As String = saveFileDialog.FileName
            My.Settings.LastPlace = txtOrt.Text.Trim
            My.Settings.Save()
            Try
                MdlPDFVKHReport.GenerateReport(saveFileDialog.FileName, chkCheck.IsChecked, txtOrt.Text.Trim, cmbBuch.Text)
                If MessageBox.Show("PDF erfolgreich gespeichert!: " & vbCrLf & "Datei öffnen?", "Hinweis", MessageBoxButtons.YesNo) = System.Windows.MessageBoxResult.Yes Then
                    Dim strReturn As String = cFH.OpenPdfFile(strFileName)
                    If strReturn <> "" Then
                        MessageBox.Show(strReturn, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show("Fehler beim Speichern der PDF: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub VKH_Ausgabe_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        txtOrt.Text = My.Settings.LastPlace
        Dim dt As DataTable = cGen.GetVKH_Books
        Dim emptyRow As DataRow = dt.NewRow()
        emptyRow("BUCH_H") = String.Empty
        dt.Rows.InsertAt(emptyRow, 0)

        cmbBuch.ItemsSource = Nothing
        cmbBuch.ItemsSource = dt.DefaultView
        cmbBuch.DisplayMemberPath = "BUCH_H"
        cmbBuch.SelectedIndex = 0
    End Sub
End Class
