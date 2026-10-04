Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Forms

Public Class AhnenTafel
    Private cAT As New inoGenDLL.clsAhnentafelDaten(My.Settings.DBPath)
    Private cGenDB As New inoGenDLL.ClsGenDB(My.Settings.DBPath)
    Private PID As Integer = 1

    Private mdFilePath As String = IO.Path.Combine(Application.MyAppFolder, "Ahnenbericht.md")

    Public Sub New()
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        btnCSV.IsEnabled = False
        btnReport.IsEnabled = False
        btnOK.IsEnabled = False
        btnChart.IsEnabled = False
        btnMap.IsEnabled = False
    End Sub

    Private Async Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
        btnOK.IsEnabled = False
        Dim blnCheck As Boolean = ckbCompress.IsChecked
        Dim blnSource As Boolean = ckbWithSources.IsChecked
        ShowProgress("Daten werden zusammengestellt...", True)
        Await Task.Run(Sub()
                           cAT.RootPersonID = PID
                           cAT.NewList()

                           If blnCheck Then
                               cAT.WriteCompTreeToFile(mdFilePath)
                           Else
                               cAT.WriteTreeToFile(mdFilePath, blnSource)
                           End If
                       End Sub)
        HideProgress()

        Dim md As String = File.ReadAllText(mdFilePath)
        MdView.Markdown = md

        btnOK.IsEnabled = True
        btnCSV.IsEnabled = True
        btnReport.IsEnabled = True
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

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub btnChart_Click(sender As Object, e As RoutedEventArgs)

        Dim Ergebnis As Boolean = True

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "PDF-Dateien (*.pdf)|*.pdf"
        saveFileDialog.Title = "PDF speichern"
        saveFileDialog.DefaultExt = "pdf"
        saveFileDialog.AddExtension = True



        ' Dialog anzeigen
        If saveFileDialog.ShowDialog() = Forms.DialogResult.OK Then
            If rbA1.IsChecked = True Then
                My.Settings.LastGenPapersize = "A1"
            ElseIf rbA2.IsChecked = True Then
                My.Settings.LastGenPapersize = "A2"
            ElseIf rbA3.IsChecked = True Then
                My.Settings.LastGenPapersize = "A3"
            Else
                My.Settings.LastGenPapersize = "A4"
            End If
            If rbFO.IsChecked = True Then
                My.Settings.LastGenColortype = "ohne"
            ElseIf rbFG.IsChecked = True Then
                My.Settings.LastGenColortype = "Geschlecht"
            Else
                My.Settings.LastGenColortype = "Zweig"
            End If
            If rbGen4.IsChecked = True Then
                My.Settings.LastGenPrintout = "Gen4"
            Else
                My.Settings.LastGenPrintout = "Gen7"
            End If
            My.Settings.Save()
            Try
                If rbGen4.IsChecked = True Then
                    MdlPdfAhnentafel.AT(cAT.Persons, saveFileDialog.FileName)
                Else
                    Ergebnis = mdlPDFAhnentafelGen.PrintAhnentafelGen7(cAT.Persons, saveFileDialog.FileName)
                End If

                If Ergebnis = True Then
                    MessageBox.Show("PDF erfolgreich gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Catch ex As Exception
                MessageBox.Show("Fehler beim Speichern der PDF: " & vbCrLf & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

    End Sub

    Private Sub btnReport_Click(sender As Object, e As RoutedEventArgs)

        Dim saveFileDialog As New SaveFileDialog()
        saveFileDialog.Filter = "PDF-Dateien (*.pdf)|*.pdf"
        saveFileDialog.Title = "PDF speichern"
        saveFileDialog.DefaultExt = "pdf"
        saveFileDialog.AddExtension = True

        If saveFileDialog.ShowDialog() = Forms.DialogResult.OK Then
            Try
                MdlPdfAncestorReport.GenerateReport(mdFilePath, saveFileDialog.FileName, $"{cAT.Persons(0).Vorname} {cAT.Persons(0).Nachname}")
                MessageBox.Show("PDF erfolgreich gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show("Fehler beim Speichern der PDF: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
    End Sub

    Private Sub AhnenTafel_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        Select Case My.Settings.LastGenPapersize
            Case "A4"
                rbA4.IsChecked = True
            Case "A3"
                rbA3.IsChecked = True
            Case "A2"
                rbA2.IsChecked = True
            Case "A1"
                rbA1.IsChecked = True
            Case Else
                rbA4.IsChecked = True
        End Select
        Select Case My.Settings.LastGenPrintout
            Case "Gen4"
                rbGen4.IsChecked = True
            Case "Gen7"
                rbGen7.IsChecked = True
            Case Else
                rbGen4.IsChecked = True
        End Select
        Select Case My.Settings.LastGenColortype
            Case "ohne"
                rbFO.IsChecked = True
            Case "Geschlecht"
                rbFG.IsChecked = True
            Case "Zweig"
                rbFZ.IsChecked = True
            Case Else
                rbFO.IsChecked = True
        End Select

        ckbCompress.IsChecked = My.Settings.ATCompressed
        ckbWithSources.IsChecked = My.Settings.ATDetails
    End Sub
    ' Code-Behind: Steuert Anzeige und Inhalt der Fortschrittsanzeige

    Private Sub ShowProgress(message As String, Optional indeterminate As Boolean = True, Optional value As Double = 0)
        ' UI-Updates auf UI-Thread ausführen
        Dispatcher.Invoke(Sub()
                              txtProgressLabel.Text = message
                              pbProgress.IsIndeterminate = indeterminate
                              If Not indeterminate Then
                                  pbProgress.Value = value
                              End If
                              progressBorder.Visibility = Visibility.Visible
                          End Sub)
    End Sub

    Private Sub UpdateProgress(value As Double, Optional message As String = Nothing)
        Dispatcher.Invoke(Sub()
                              pbProgress.IsIndeterminate = False
                              pbProgress.Value = value
                              If Not String.IsNullOrEmpty(message) Then txtProgressLabel.Text = message
                          End Sub)
    End Sub

    Private Sub HideProgress()
        Dispatcher.Invoke(Sub()
                              progressBorder.Visibility = Visibility.Collapsed
                          End Sub)
    End Sub

    Private Sub AhnenTafel_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        My.Settings.ATCompressed = ckbCompress.IsChecked
        My.Settings.ATDetails = ckbWithSources.IsChecked
        My.Settings.Save()
        MyBase.Finalize()
    End Sub
End Class
