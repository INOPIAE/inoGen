Imports System.Data
Imports System.Data.OleDb

Public Class ClsGenDB

    Public connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", "")


    Public Sub New(dbFileString As String)
        connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";Persist Security Info=True", dbFileString)
        ' Constructor logic if needed
    End Sub

    Public Sub FillPerson(ByRef PD As clsAhnentafelDaten.PersonData)
        Dim strSQL As String = "SELECT
                tblPerson.*,
                tblNachname.Nachname,
                tblKonfession.Konfessionkurz
            FROM
                (
                    tblPerson
                    LEFT JOIN tblNachname ON tblPerson.tblNachnameID = tblNachname.tblNachnameID
                )
                LEFT JOIN tblKonfession ON tblPerson.tblKonfessionID = tblKonfession.tblKonfessionID
            WHERE tblPersonID = ?"
        Dim PNAme As String = ""
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("@tblPersonID", PD.ID)

                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        PD.Vorname = If(IsDBNull(reader("Vorname")), "", reader("Vorname").ToString())
                        PD.Nachname = If(IsDBNull(reader("Nachname")), "", reader("Nachname").ToString())
                        PD.Geschlecht = If(IsDBNull(reader("Sex")), "", reader("Sex").ToString())
                        PD.Konfession = If(IsDBNull(reader("Konfessionkurz")), "", reader("Konfessionkurz").ToString())
                        PD.FID = If(IsDBNull(reader("tblFamilieID")), 0, Convert.ToInt32(reader("tblFamilieID")))
                        PD.PS = If(IsDBNull(reader("PS")), "", reader("PS").ToString())
                        PD.FSID = If(IsDBNull(reader("FSID")), "", reader("FSID").ToString())
                    End If
                End Using
            End Using
        End Using
    End Sub

    Public Sub FillPersonEltern(ByRef PD As clsAhnentafelDaten.PersonData)
        Dim strSQL As String = "SELECT
                tblFamilie.*
            FROM
                tblFamilie
            WHERE tblFamilieID = ?"
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("@tblFamilieID", PD.FID)
                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        PD.V = If(IsDBNull(reader("tblPersonIDV")), 0, Convert.ToInt32(reader("tblPersonIDV")))
                        PD.M = If(IsDBNull(reader("tblPersonIDM")), 0, Convert.ToInt32(reader("tblPersonIDM")))
                    End If
                End Using
            End Using
        End Using
    End Sub

    Public Sub FillPersonData(ByRef PD As clsAhnentafelDaten.PersonData)
        FillPerson(PD)
        FillPersonEltern(PD)
        FillPersonDaten(PD)
        If PD.EID > 0 Then
            FillFamilieDaten(PD)
        End If

    End Sub

    Public Sub FillPersonDaten(ByRef PD As clsAhnentafelDaten.PersonData)
        Dim strSQL As String = "SELECT
                tblEreignis.tblEreignisID,
                tblEreignis.tblEreignisArtID,
                tblEreignisArt.EreignisArt AS Ereignis,
                tblEreignis.DatumText AS Datum,
                tblEreignis.Datum AS HDatum,
                tblEreignis.BisDatumText AS BDatum,
                tblEreignis.BisDatum AS BHDatum,
                IIf([tblKreis]![Kreis]<>"""",[tblOrt]![Ort] & "" ("" & [tblKreis]![Kreis] & "")"",[tblOrt]![Ort]) AS Ort,
                tblKonfession.Konfessionkurz AS Konfession,
                tblEreignis.Referenz,
                tblEreignis.FSID,
                tblEreignis.Info
            FROM
                (
                    (
                        (
                            tblEreignis
                            INNER JOIN tblEreignisArt ON tblEreignis.tblEreignisArtID = tblEreignisArt.tblEreignisArtID
                        )
                        INNER JOIN tblKonfession ON tblEreignis.tblKonfessionID = tblKonfession.tblKonfessionID
                    )
                    INNER JOIN tblOrt ON tblEreignis.tblOrtID = tblOrt.tblOrtID
                )
                LEFT JOIN tblKreis ON tblOrt.tblKreisID = tblKreis.tblKreisID
            WHERE tblPersonID = ? and tblFamilieID = 0
            ORDER BY
                tblEreignisArt.Reihenfolge,
                tblEreignis.Datum"
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("?", PD.ID) 'Parameter einsetzen

                Using rdr As OleDbDataReader = cmd.ExecuteReader()
                    While rdr.Read()
                        ' Werte in Variablen einlesen
                        Select Case rdr("tblEreignisArtID")
                            Case 1
                                PD.Geburtsdatum = GetDatum(rdr)
                                PD.Geburtsort = rdr("Ort").ToString()
                            Case 2
                                PD.Taufdatum = GetDatum(rdr)
                                PD.Taufort = rdr("Ort").ToString()
                            Case 6
                                PD.Sterbedatum = GetDatum(rdr)
                                PD.Sterbeort = rdr("Ort").ToString()
                            Case 7
                                PD.Begräbnisdatum = GetDatum(rdr)
                                PD.Begräbnisort = rdr("Ort").ToString()
                            Case >= 8
                                PD.Sonstige = True
                        End Select

                    End While
                End Using
            End Using
        End Using

    End Sub

    Private Shared Function GetDatum(rdr As OleDbDataReader) As String
        Dim strDatum As String = ""
        strDatum = rdr("Datum").ToString()
        If IsDBNull(rdr("BDatum")) = False Then
            If rdr("BDatum").ToString.Trim <> "" Then
                strDatum &= " - " & rdr("BDatum").ToString()
            End If
        End If
        Return strDatum
    End Function

    Public Sub FillFamilieDaten(ByRef PD As clsAhnentafelDaten.PersonData)
        Dim strSQL As String = "SELECT
                tblEreignis.tblEreignisID,
                tblEreignis.tblEreignisArtID,
                tblEreignisArt.EreignisArt AS Ereignis,
                tblEreignis.DatumText AS Datum,
                tblEreignis.Datum AS HDatum,
                tblEreignis.BisDatumText AS BDatum,
                tblEreignis.BisDatum AS BHDatum,
                IIf([tblKreis]![Kreis]<>"""",[tblOrt]![Ort] & "" ("" & [tblKreis]![Kreis] & "")"",[tblOrt]![Ort]) AS Ort,
                tblKonfession.Konfessionkurz AS Konfession,
                tblEreignis.Referenz,
                tblEreignis.FSID,
                tblEreignis.Info
            FROM
                (
                    (
                        (
                            tblEreignis
                            INNER JOIN tblEreignisArt ON tblEreignis.tblEreignisArtID = tblEreignisArt.tblEreignisArtID
                        )
                        INNER JOIN tblKonfession ON tblEreignis.tblKonfessionID = tblKonfession.tblKonfessionID
                    )
                    INNER JOIN tblOrt ON tblEreignis.tblOrtID = tblOrt.tblOrtID
                )
                LEFT JOIN tblKreis ON tblOrt.tblKreisID = tblKreis.tblKreisID
            WHERE tblPersonID = 0 and tblFamilieID = ?
            ORDER BY
                tblEreignisArt.Reihenfolge,
                tblEreignis.Datum"
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("?", PD.EID) 'Parameter einsetzen

                Using rdr As OleDbDataReader = cmd.ExecuteReader()
                    While rdr.Read()
                        ' Werte in Variablen einlesen
                        Select Case rdr("tblEreignisArtID")
                            Case 3
                                PD.Heiratdatum = GetDatum(rdr)
                                PD.Heiratort = rdr("Ort").ToString()
                            Case 4
                                PD.KHeiratdatum = GetDatum(rdr)
                                PD.KHeiratort = rdr("Ort").ToString()
                            Case 5
                                PD.Scheidungsdatum = GetDatum(rdr)
                                PD.Scheidungsort = rdr("Ort").ToString()
                            Case 8
                                PD.Verlobungsdatum = GetDatum(rdr)
                                PD.Verlobungsort = rdr("Ort").ToString()
                            Case >= 8
                                PD.Sonstige = True
                        End Select

                    End While
                End Using
            End Using
        End Using


    End Sub

    Public Function PersonenDaten(ID As Int16) As String
        Dim strSQL As String = "SELECT
                tblPerson.tblPersonID,
                tblPerson.PS,
                tblPerson.Vorname,
                tblNachname.Nachname
            FROM
                tblPerson
                LEFT JOIN tblNachname ON tblPerson.tblNachnameID = tblNachname.tblNachnameID
            WHERE tblPersonID = ?"
        Dim PNAme As String = ""
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("@tblPersonID", ID)

                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    If reader.Read() Then

                        PNAme = If(IsDBNull(reader("PS")), "", reader("PS").ToString()) & " " &
                        If(IsDBNull(reader("Vorname")), "", reader("Vorname").ToString()) & " " &
                        If(IsDBNull(reader("Nachname")), "", reader("Nachname").ToString().ToUpper)
                    End If
                End Using
            End Using
        End Using
        Return PNAme
    End Function
End Class
