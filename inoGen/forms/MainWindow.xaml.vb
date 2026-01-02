Imports System.Collections.Specialized
Imports System.IO
Imports System.Windows.Forms
Imports inoGenDLL
Imports Microsoft.Win32
Imports OpenFileDialog = System.Windows.Forms.OpenFileDialog
Imports SaveFileDialog = System.Windows.Forms.SaveFileDialog

Class MainWindow

    Public connectionString As String =
        String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", "D:\Daten\programierung neu\inoGen\Daten\Drews.accdb")

    Public Shared fsWindow As FamilySearchWeb = Nothing

    Public Sub New()

        InitializeComponent()

        If My.Settings.RecentFiles IsNot Nothing Then
            Start()
        Else
            New_Click(Nothing, New RoutedEventArgs())
        End If


    End Sub

    Private Sub Orte_Click(sender As Object, e As RoutedEventArgs)
        MainContent.Content = New orte()
    End Sub

    Private Sub Kreise_Click(sender As Object, e As RoutedEventArgs)
        MainContent.Content = New kreise()
    End Sub

    Private Sub Person_Click(sender As Object, e As RoutedEventArgs)
        MainContent.Content = New personen()
        My.Settings.LastContent = "Person"
        My.Settings.Save()
    End Sub

    Private Sub Familie_Click(sender As Object, e As RoutedEventArgs)
        MainContent.Content = New familien()
        My.Settings.LastContent = "Familie"
        My.Settings.Save()
    End Sub

    Private Sub Konfession_Click(sender As Object, e As RoutedEventArgs)
        MainContent.Content = New konfession()
    End Sub

    Private Sub Quit_Click(sender As Object, e As RoutedEventArgs)
        ShutDown()
    End Sub

    Private Sub ShutDown()
        StatusText.Text = "Programm wird beendet..."
        Application.Current.Shutdown()
    End Sub

    Private Sub Start()
        Dim strDB As String
        Dim cDB As ClsDatabase
        Dim DbVersion As Long

        StatusText.Text = "Datenquelle wird geladen..."

        strDB = My.Settings.DBPath
        cDB = New ClsDatabase(strDB)
        DbVersion = cDB.CheckDBVersion

        If My.Settings.LastContent <> "" Then
            Select Case My.Settings.LastContent
                Case "Person"
                    Person_Click(Nothing, Nothing)
                Case "Familie"
                    Familie_Click(Nothing, Nothing)
                Case "VHK"
                    VKH_Click(Nothing, Nothing)
            End Select
        End If


        AddToRecentFiles(strDB)
        RefreshRecentFilesMenu()

        DBText.Text = "Datenbank: " & Path.GetFileNameWithoutExtension(strDB)
        StatusText.Text = ""

        If My.Settings.Email = "" Then
            Options_Click(Nothing, Nothing)
        End If
    End Sub

    Public Sub ShowContent(ctrl As UserControl)
        MainContent.Content = ctrl
    End Sub

    Private Sub FranzKalender_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New FranzKalender()

        win.Show()
    End Sub

    Private Sub Ahnentafel_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New AhnenTafel()

        win.Show()
    End Sub
    Private Sub PersonReport_Click(sender As Object, e As RoutedEventArgs)
        Dim PID As Integer
        Dim win As New SuchePerson()
        AddHandler win.PersonSelected, Sub(id, persontext)
                                           PID = id
                                           Dim saveFileDialog As New SaveFileDialog()
                                           saveFileDialog.Filter = "PDF-Dateien (*.pdf)|*.pdf"
                                           saveFileDialog.Title = "PDF speichern"
                                           saveFileDialog.DefaultExt = "pdf"
                                           saveFileDialog.AddExtension = True

                                           ' Dialog anzeigen
                                           If saveFileDialog.ShowDialog() = Forms.DialogResult.OK Then
                                               Try
                                                   MdlPdfPersonReport.GenerateReport(PID, saveFileDialog.FileName)

                                                   MessageBox.Show("PDF erfolgreich gespeichert!", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                               Catch ex As Exception
                                                   MessageBox.Show("Fehler beim Speichern der PDF: " & ex.Message, "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End Try
                                           End If
                                       End Sub
        win.Show()

    End Sub
    Private Sub Info_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New AboutWindow()

        win.Show()
    End Sub

    Private Sub Ereignisart_Click(sender As Object, e As RoutedEventArgs)
        MainContent.Content = New ereignisart()
    End Sub

    Private Sub New_Click(sender As Object, e As RoutedEventArgs)
        Dim saveDialog As New SaveFileDialog()
        saveDialog.Title = "Neue Datenquelle anlegen"
        saveDialog.Filter = "Daten (*.inoGdb)|*.inoGdb"
        saveDialog.FileName = "Datenquelle.inoGdb"
        saveDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        If saveDialog.ShowDialog() = Forms.DialogResult.OK Then
            Dim filePath As String = saveDialog.FileName
            SaveSettingAfterFileChange(filePath)
            Start()
        End If
    End Sub


    Private Sub Open_Click(sender As Object, e As RoutedEventArgs)
        Dim openDialog As New OpenFileDialog()
        openDialog.Title = "Datei öffnen"
        openDialog.Filter = "Daten (*.inoGdb)|*.inoGdb"
        openDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        If openDialog.ShowDialog() = True Then
            Dim filePath As String = openDialog.FileName
            SaveSettingAfterFileChange(filePath)
            Start()
        End If
    End Sub

    Private Shared Sub SaveSettingAfterFileChange(filePath As String)
        My.Settings.DBPath = filePath
        My.Settings.LastVKHID = 0
        My.Settings.LastPID = 0
        My.Settings.LastFID = 0
        My.Settings.Save()
    End Sub

    Public Sub AddToRecentFiles(filePath As String)
        If My.Settings.RecentFiles Is Nothing Then
            My.Settings.RecentFiles = New StringCollection()
        End If

        If My.Settings.RecentFiles.Contains(filePath) Then
            My.Settings.RecentFiles.Remove(filePath)
        End If

        My.Settings.RecentFiles.Insert(0, filePath)

        While My.Settings.RecentFiles.Count > 5
            My.Settings.RecentFiles.RemoveAt(My.Settings.RecentFiles.Count - 1)
        End While

        My.Settings.Save()
    End Sub

    Private Sub RefreshRecentFilesMenu()
        RecentFilesMenu.Items.Clear()

        If My.Settings.RecentFiles IsNot Nothing Then
            For Each filePath In My.Settings.RecentFiles
                Dim item As New MenuItem() With {.Header = filePath}
                AddHandler item.Click, Sub()
                                           My.Settings.DBPath = filePath
                                           SaveSettingAfterFileChange(filePath)
                                           Start()
                                       End Sub
                RecentFilesMenu.Items.Add(item)
            Next
        End If
    End Sub

    Private Sub MainWindow_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        ShutDown()
    End Sub

    Private Sub VKH_Click(sender As Object, e As RoutedEventArgs)

        Dim ctrl As New vkHeirat()

        AddHandler ctrl.RequestResizeMainWindow, AddressOf OnRequestResize

        MainContent.Content = ctrl
        My.Settings.LastContent = "VHK"
        My.Settings.Save()
    End Sub

    Private Sub OnRequestResize(newWidth As Double)
        If Me.Width < newWidth Then
            Me.Width = newWidth
        End If
    End Sub

    Public Sub ShowContent(content As Object)
        MainContent.Content = content
    End Sub

    Private Sub Options_Click(sender As Object, e As RoutedEventArgs)
        Dim Options = New OptionsWindow()
        Options.ShowDialog()
    End Sub

    Private Sub Statistics_Click(sender As Object, e As RoutedEventArgs)
        Dim Stats = New Statistics
        Stats.ShowDialog()
    End Sub

    Private Sub VKH_PersonReport_Click(sender As Object, e As RoutedEventArgs)
        Dim VKH_Person = New VKH_Personen
        VKH_Person.Show()
    End Sub

    Private Sub Kirchenjahr_Click(sender As Object, e As RoutedEventArgs)
        Dim Kirchenjahr = New Kirchenjahr
        Kirchenjahr.Show()
    End Sub

End Class
