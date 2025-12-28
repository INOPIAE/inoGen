Imports System.Text.Json
Imports System.Net.Http
Public Class ClsOSMKarte

    Public Structure marker
        Public lat As Double
        Public lon As Double
        Public title As String
        Public description As String
        Public icon As String
        Public Count As Integer
    End Structure

    Public Structure GeoCodeResult
        Public lat As String
        Public lon As String
        Public display_name As String
    End Structure

    Public Class NominatimResult
        Public Property lat As String
        Public Property lon As String
        Public Property display_name As String
    End Class

    Public Property Email As String = ""

    Public markers As New List(Of marker)

    Public Function FindGeoCode(searchText As String) As List(Of GeoCodeResult)
        Dim results As New List(Of GeoCodeResult)()
        Try
            Dim url As String = "https://nominatim.openstreetmap.org/search?format=json&q=" & Uri.EscapeDataString(searchText)
            Dim client As New HttpClient()
            client.DefaultRequestHeaders.Add("User-Agent", String.Format("inoGEN/1.0 ({0})", Email))

            Dim response As HttpResponseMessage = client.GetAsync(url).Result
            If response.IsSuccessStatusCode Then
                Dim jsonResponse As String = response.Content.ReadAsStringAsync().Result
                Dim options As New JsonSerializerOptions With {
                    .PropertyNameCaseInsensitive = True
                }

                Dim locations As List(Of NominatimResult) = JsonSerializer.Deserialize(Of List(Of NominatimResult))(jsonResponse, options)


                If locations IsNot Nothing Then
                        For Each l In locations
                        Dim geo As New GeoCodeResult()

                        geo.lat = l.lat
                            geo.lon = l.lon
                            geo.display_name = l.display_name
                            results.Add(geo)
                        Next
                    End If

                End If
        Catch ex As Exception
            ' Fehlerbehandlung
            results = New List(Of GeoCodeResult)()
        End Try
        Return results
    End Function

End Class
