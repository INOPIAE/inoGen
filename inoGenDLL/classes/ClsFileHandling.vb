Public Class ClsFileHandling
    Public Function OpenPdfFile(filePath As String) As String
        Try
            If IO.File.Exists(filePath) Then
                Process.Start(New ProcessStartInfo(filePath) With {
                    .UseShellExecute = True
                })
            Else
                Return "Datei nicht gefunden: " & filePath
            End If
        Catch ex As Exception
            Return "Fehler beim Öffnen der PDF: " & ex.Message
        End Try
        Return ""
    End Function

    Public Function OpenPngFile(filePath As String) As String
        Try

            If IO.File.Exists(filePath) Then

                Process.Start(
                    New ProcessStartInfo(filePath) With {
                        .UseShellExecute = True
                    })

            Else

                Return "Datei nicht gefunden: " & filePath

            End If

        Catch ex As Exception

            Return "Fehler beim Öffnen der PNG-Datei: " & ex.Message

        End Try

        Return ""

    End Function
End Class
