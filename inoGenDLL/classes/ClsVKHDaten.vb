Imports System.Data
Imports System.Data.OleDb

Public Class ClsVKHDaten
    Public DBPath As String = ""
    Public connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", "")

    Public cGenDB As inoGenDLL.ClsGenDB
    Public cOSM As New inoGenDLL.ClsOSMKarte

    Public LocationList As New List(Of ClsOSMKarte.marker)

    Public Sub New(dbFileString As String)
        DBPath = dbFileString
        connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", DBPath)
        cGenDB = New inoGenDLL.ClsGenDB(DBPath)
    End Sub

    Public Function GetGeoData(Location As String, Optional Total As Integer = 1) As ClsOSMKarte.marker

        Dim strSql As String = "SELECT
                Ort,
                Breite,
                Laenge
            FROM
                tblOrt
            WHERE
                ORT = ?"

        Dim marker As New ClsOSMKarte.marker

        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(strSql, conn)
                cmd.Parameters.AddWithValue("?", Location) 'Parameter einsetzen

                Using rdr As OleDbDataReader = cmd.ExecuteReader()
                    While rdr.Read()
                        If rdr("Breite") Is DBNull.Value Or rdr("Laenge") Is DBNull.Value Then
                            Continue While
                        End If

                        marker.lat = Convert.ToDouble(rdr("Breite"))
                        marker.lon = Convert.ToDouble(rdr("Laenge"))
                        marker.title = rdr("Ort")
                        marker.Count = Total
                    End While
                End Using
            End Using
        End Using
        Return marker
    End Function

    Public Sub ErstelleLocationList()
        LocationList = New List(Of ClsOSMKarte.marker)
        Dim dt As DataTable = cGenDB.StatisticsVKHLocations

        For Each r As DataRow In dt.Rows
            Dim newMarker As ClsOSMKarte.marker = GetGeoData(r.Item(0), r.Item(1))
            If IsNothing(newMarker.title) = False Then
                LocationList.Add(newMarker)
            End If
        Next
    End Sub

End Class
