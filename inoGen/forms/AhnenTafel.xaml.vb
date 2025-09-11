Imports System.IO

Public Class AhnenTafel
    Private cAT As New inoGenDLL.clsAhnentafelDaten(My.Settings.DBPath)
    Private cGenDB As New inoGenDLL.ClsGenDB(My.Settings.DBPath)
    Dim PID As Integer = 1
    Public Sub New()
        ' Dieser Aufruf ist für den Designer erforderlich.
        InitializeComponent()
        btnCSV.IsEnabled = False
        btnOK.IsEnabled = False
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
        btnMap.IsEnabled = True
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New SuchePerson()
        AddHandler win.PersonSelected, Sub(id)

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
End Class
