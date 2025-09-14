Imports System.IO
Imports System.Windows.Forms

Public Class AhnenTafel
    Private cAT As New inoGenDLL.clsAhnentafelDaten(My.Settings.DBPath)
    Private cGenDB As New inoGenDLL.ClsGenDB(My.Settings.DBPath)
    Dim PID As Integer = 1
    Public Sub New()
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        btnCSV.IsEnabled = False
        btnOK.IsEnabled = False
        btnChart.IsEnabled = False
        btnMap.IsEnabled = False
    End Sub
    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
        Dim mdFilePath As String = IO.Path.Combine(Application.MyAppFolder, "Ahnenbericht.md")
        cAT.RootPersonID = PID
        cAT.NewList()
        If ckbCompress.IsChecked Then
            cAT.WriteCompTreeToFile(mdFilePath)
        Else
            cAT.WriteTreeToFile(mdFilePath)
        End If

        Dim md As String = File.ReadAllText(mdFilePath)
        MdView.Markdown = md
        btnCSV.IsEnabled = True
        btnChart.IsEnabled = True
        btnMap.IsEnabled = True
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New SuchePerson()
        AddHandler win.PersonSelected, Sub(id, persontext)

                                           PID = id
                                           txtPerson.Text = cGenDB.PersonenDaten(id)
                                           btnOK.IsEnabled = True
                                       End Sub

        win.Show()
    End Sub

    Private Sub btnCSV_Click(sender As Object, e As RoutedEventArgs)
        Dim mdFilePath As String = IO.Path.Combine(Application.MyAppFolder, "ahnentafel.csv")
        cAT.RootPersonID = PID
        cAT.NewList()

        cAT.WriteToCSV(mdFilePath)
        MessageBox.Show("abgeschlossen")
    End Sub
    Private Sub btnMap_Click(sender As Object, e As RoutedEventArgs)

        Dim win As New OSMKarte(cAT.LocationList, cAT.Persons)
        win.Show()

    End Sub

    Private Sub btnChart_Click(sender As Object, e As RoutedEventArgs)
        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "PDF-Dateien (*.pdf)|*.pdf"
        saveFileDialog.Title = "PDF speichern"
        saveFileDialog.DefaultExt = "pdf"
        saveFileDialog.AddExtension = True

        ' Dialog anzeigen
        If saveFileDialog.ShowDialog() = Forms.DialogResult.OK Then
            Try
                MdlPdfAhnentafel.AT(cAT.Persons, saveFileDialog.FileName)
                MessageBox.Show("PDF erfolgreich gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Fehler beim Speichern der PDF: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

    End Sub
End Class
