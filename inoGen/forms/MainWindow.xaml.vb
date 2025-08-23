Imports inoGenDLL
Imports System.IO
Imports Microsoft.Win32
Imports System.Collections.Specialized

Class MainWindow

    Public connectionString As String =
        String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", "D:\Daten\programierung neu\inoGen\Daten\Drews.accdb")


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
    End Sub

    Private Sub Familie_Click(sender As Object, e As RoutedEventArgs)
        MainContent.Content = New familien()
    End Sub

    Private Sub Konfession_Click(sender As Object, e As RoutedEventArgs)
        MainContent.Content = New konfession()
    End Sub

    Private Sub Quit_Click(sender As Object, e As RoutedEventArgs)
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

        MainContent.Content = New personen()

        AddToRecentFiles(strDB)
        RefreshRecentFilesMenu()

        DBText.Text = "Datenbank: " & Path.GetFileNameWithoutExtension(strDB)
        StatusText.Text = ""
    End Sub

    Public Sub ShowContent(ctrl As UserControl)
        MainContent.Content = ctrl
    End Sub

    Private Sub Ahnentafel_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New AhnenTafel()

        win.Show()
    End Sub

    Private Sub New_Click(sender As Object, e As RoutedEventArgs)
        Dim saveDialog As New SaveFileDialog()
        saveDialog.Title = "Neue Datenquelle anlegen"
        saveDialog.Filter = "Daten (*.inoGdb)|*.inoGdb"
        saveDialog.FileName = "Datenquelle.inoGdb"
        saveDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

        If saveDialog.ShowDialog() = True Then
            Dim filePath As String = saveDialog.FileName
            My.Settings.DBPath = filePath
            My.Settings.Save()
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
            My.Settings.DBPath = filePath
            My.Settings.Save()
            Start()
        End If
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
                                           My.Settings.Save()
                                           Start()
                                       End Sub
                RecentFilesMenu.Items.Add(item)
            Next
        End If
    End Sub

End Class
