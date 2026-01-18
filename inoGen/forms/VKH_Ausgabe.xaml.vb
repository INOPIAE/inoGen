Imports System.Windows.Forms
Imports inoGenDLL

Public Class VKH_Ausgabe
    Private connectionString As String =
        String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)

    Private cGeoD As New ClsGeoDaten(My.Settings.DBPath)
    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnOSMMap_Click(sender As Object, e As RoutedEventArgs) Handles btnOSMMap.Click
        cGeoD.ErstelleLocationList()

        Dim win As New OSMKarte(cGeoD.LocationList)
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
            My.Settings.LastPlace = txtOrt.Text.Trim
            My.Settings.Save()
            Try
                MdlPDFVKHReport.GenerateReport(saveFileDialog.FileName, txtOrt.Text.Trim)
                MessageBox.Show("PDF erfolgreich gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Fehler beim Speichern der PDF: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub VKH_Ausgabe_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        txtOrt.Text = My.Settings.LastPlace
    End Sub
End Class
