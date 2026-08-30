Imports System.IO
Imports inoGenDLL
Public Class GEDCOM_Export
    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub BtnExport_Click(sender As Object, e As RoutedEventArgs)
        Dim cGC As New ClsGedcom
        cGC.CreateGedcomFile(TxtDatei.Text, My.Settings.DBPath, "Marcus Mängel")
    End Sub

    Private Sub BtnFile_Click(sender As Object, e As RoutedEventArgs)
        Dim saveDialog As New Microsoft.Win32.SaveFileDialog() With {
            .Filter = "GEDCOM-Dateien (*.ged)|*.ged",
            .DefaultExt = ".ged",
            .Title = "GEDCOM-Datei speichern",
            .OverwritePrompt = True
        }

        ' Aktuellen Pfad aus TextBox als Startpunkt verwenden (falls vorhanden)
        If Not String.IsNullOrEmpty(TxtDatei.Text) Then
            saveDialog.FileName = Path.GetFileName(TxtDatei.Text)
            saveDialog.InitialDirectory = Path.GetDirectoryName(TxtDatei.Text)
        End If

        If saveDialog.ShowDialog() = True Then
            TxtDatei.Text = saveDialog.FileName
        End If
    End Sub
End Class
