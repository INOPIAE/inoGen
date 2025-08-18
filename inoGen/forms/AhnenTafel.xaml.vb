Imports System.IO

Public Class AhnenTafel
    Private cAT As New inoGenDLL.clsAhnentafelDaten(My.Settings.DBPath)
    Private cGenDB As New inoGenDLL.ClsGenDB(My.Settings.DBPath)
    Dim PID As Integer = 1
    Private Sub btnOK_Click(sender As Object, e As RoutedEventArgs)
        Dim mdFilePath As String = IO.Path.Combine(Application.MyAppFolder, "Ahnenbericht.md")
        cAT.RootPersonID = PID
        cAT.NewList()
        cAT.WriteTreeToFile(mdFilePath)
        Dim md As String = File.ReadAllText(mdFilePath)
        MdView.Markdown = md
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New SuchePerson()
        AddHandler win.PersonSelected, Sub(id)

                                           PID = id
                                           txtPerson.Text = cGenDB.PersonenDaten(id)
                                       End Sub

        win.Show()
    End Sub
End Class
