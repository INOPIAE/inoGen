Imports inoGenDLL

Public Class VKH_Ausgabe
    Private connectionString As String =
        String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)

    Private cVKHD As New ClsVKHDaten(My.Settings.DBPath)
    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnOSMMap_Click(sender As Object, e As RoutedEventArgs) Handles btnOSMMap.Click
        cVKHD.ErstelleLocationList()

        Dim win As New OSMKarte(cVKHD.LocationList)
        win.Show()
    End Sub
End Class
