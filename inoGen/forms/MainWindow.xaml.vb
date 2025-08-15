Imports inoGenDLL

Class MainWindow

    Public connectionString As String =
        String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", "D:\Daten\programierung neu\inoGen\Daten\Drews.accdb")


    Public Sub New()

        InitializeComponent()

        Start()

        MainContent.Content = New personen()

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

        'Dim uc As New konfession(connectionString)
        'MainContent.Content = uc
    End Sub

    Private Sub Quit_Click(sender As Object, e As RoutedEventArgs)

        Application.Current.Shutdown()
    End Sub

    Private Sub Start()
        Dim strDB As String
        Dim cDB As ClsDatabase
        Dim DbVersion As Long
        strDB = My.Settings.DBPath
        cDB = New ClsDatabase(strDB)
        DbVersion = cDB.CheckDBVersion
    End Sub

    Public Sub ShowContent(ctrl As UserControl)
        MainContent.Content = ctrl
    End Sub
End Class
