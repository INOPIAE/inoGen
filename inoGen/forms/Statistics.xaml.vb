Imports System.Data
Imports inoGenDLL

Public Class Statistics
    Private connectionString As String =
   String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)

    Private cGDB As New ClsGenDB(My.Settings.DBPath)
    Private Sub CloseButton_Click(sender As Object, e As RoutedEventArgs)
        Close()
    End Sub

    Private Sub Statistics_Initialized(sender As Object, e As EventArgs) Handles Me.Initialized
        CommonData()
        VKHData()
    End Sub

    Private Sub CommonData()
        A1.Content = "Vornamen gesamt: " & Format(cGDB.StatisicsVornamen, "#,##0")
        A2.Content = "Nachnamen gesamt: " & Format(cGDB.StatisicsNachnamen, "#,##0")
        A3.Content = "Orte gesamt: " & Format(cGDB.StatisicsOrte, "#,##0")
        A4.Content = "Personen gesamt: " & Format(cGDB.StatisicsPersonen, "#,##0")
        A5.Content = "Familien gesamt: " & Format(cGDB.StatisicsFamilien, "#,##0")
    End Sub

    Private Sub VKHData()
        Dim dt As DataTable = cGDB.StatisicsVKHeirat
        With dt.Rows(0)
            H1.Content = "Einträge gesamt: " & Format(.Item(0), "#,##0")
            H2.Content = "Bräutigame gesamt: " & Format(.Item(1), "#,##0")
            H3.Content = "Väter des Bräutigams gesamt: " & Format(.Item(2), "#,##0")
            H4.Content = "Mütter des Bräutigams gesamt: " & Format(.Item(3), "#,##0")
            H5.Content = "Bräute gesamt: " & Format(.Item(4), "#,##0")
            H6.Content = "Väter der Braut gesamt: " & Format(.Item(5), "#,##0")
            H7.Content = "Mütter des Braut gesamt: " & Format(.Item(6), "#,##0")
            H8.Content = "Personen gesamt: " & Format(.Item(1) + .Item(2) + .Item(3) + .Item(4) + .Item(5) + .Item(6), "#,##0")
            H9.Content = "Zeugen gesamt: " & Format(.Item(7) + .Item(8) + .Item(9) + .Item(10), "#,##0")
        End With
    End Sub
End Class
