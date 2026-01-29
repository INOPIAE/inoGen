Imports System.Data
Imports System.Data.OleDb
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports ADODB

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

    Public Function CalculateDatum(Datum As String) As Nullable(Of Date)
        Dim d As Nullable(Of Date) = Nothing
        Dim day As Integer
        Dim month As Integer
        Dim year As Integer
        Dim parts As String()
        If Datum IsNot Nothing Then
            If IsDate(Datum) Then
                d = CDate(Datum)
            Else
                Dim cleaned As String = Regex.Replace(Datum, "[^0-9. ]", "").Trim
                If IsDate(cleaned) Then
                    d = CDate(cleaned)
                Else
                    parts = cleaned.Split(New Char() {"."c, " "c}, StringSplitOptions.RemoveEmptyEntries)
                    If parts.Length = 2 Then
                        If parts(0).Length <= 2 AndAlso parts(1).Length > 2 Then
                            If Integer.TryParse(parts(0), month) AndAlso Integer.TryParse(parts(1), year) Then
                                If month >= 1 AndAlso month <= 12 AndAlso year >= 100 AndAlso year <= 9999 Then
                                    d = New Date(year, month, 1)
                                End If
                            End If
                        End If
                    End If
                    If parts.Length = 1 Then
                        If Integer.TryParse(parts(0), year) Then
                            If year >= 100 AndAlso year <= 9999 Then
                                d = New Date(year, 1, 1)
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Return d
    End Function

    Public Function GetPersonenAdditionalData(PID As Integer) As List(Of clsAhnentafelDaten.EventData)
        Dim strSQL As String = "SELECT
                tblEreignis.tblEreignisID,
                tblEreignis.tblEreignisArtID,
                tblEreignisArt.EreignisArt AS Ereignis,
                tblEreignis.DatumText AS Datum,
                tblEreignis.Datum AS HDatum,
                tblEreignis.BisDatumText AS BDatum,
                tblEreignis.BisDatum AS BHDatum,
                tblEreignis.Zusatz AS Zusatz,
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
            WHERE tblPersonID = ? AND tblFamilieID = 0 AND tblEreignis.tblEreignisArtID > 8
            ORDER BY
                tblEreignisArt.Reihenfolge,
                tblEreignis.Datum"
        Dim EDL As New List(Of clsAhnentafelDaten.EventData)
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("?", PID) 'Parameter einsetzen

                Using rdr As OleDbDataReader = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim ED As New clsAhnentafelDaten.EventData
                        ED.ID = PID
                        ED.Person = True
                        ED.EventID = rdr("tblEreignisArtID")
                        ED.EventDate = GetDatum(rdr)
                        ED.EventLocation = rdr("Ort").ToString()
                        ED.Eventname = rdr("Ereignis").ToString()
                        ED.EventTopic = rdr("Zusatz").ToString()

                        EDL.Add(ED)
                    End While
                End Using
            End Using
        End Using
        Return EDL
    End Function

    Public Function GetFamilies(PID As Integer) As List(Of clsAhnentafelDaten.FamilyData)
        Dim strSQL As String = "SELECT
                tblFamilie.tblFamilieID,
                tblFamilie.FS,
                tblFamilie.tblPersonIDV,
                tblFamilie.tblPersonIDM,
                Min(tblEreignis.Datum) AS Datum
            FROM
                tblFamilie
                LEFT JOIN tblEreignis ON tblFamilie.tblFamilieID = tblEreignis.tblFamilieID
            GROUP BY
                tblFamilie.tblFamilieID,
                tblFamilie.FS,
                tblFamilie.tblPersonIDV,
                tblFamilie.tblPersonIDM
            HAVING
                tblFamilie.tblPersonIDV = ? 
                OR tblFamilie.tblPersonIDM = ?
            ORDER BY
                Min(tblEreignis.Datum)"
        Dim FL As New List(Of clsAhnentafelDaten.FamilyData)
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("?", PID)
                cmd.Parameters.AddWithValue("?", PID)

                Using rdr As OleDbDataReader = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim FD As New clsAhnentafelDaten.FamilyData
                        FD.ID = rdr("tblFamilieID")
                        FD.VID = If(IsDBNull(rdr("tblPersonIDV")), Nothing, CType(rdr("tblPersonIDV"), Integer))
                        FD.MID = If(IsDBNull(rdr("tblPersonIDM")), Nothing, CType(rdr("tblPersonIDM"), Integer))
                        FL.Add(FD)
                    End While
                End Using
            End Using
        End Using
        Return FL
    End Function

    Public Function StatisicsVKHeirat() As DataTable
        Dim strSQL As String = "SELECT
                COUNT(tblVKHID) AS Total,
                COUNT(IIF(FN_BR IS NOT NULL AND FN_BR<>'',1,NULL)) AS TotalBR,
                COUNT(IIF(FN_VBR IS NOT NULL AND FN_VBR<>'',1,NULL)) AS TotalVBR,
                COUNT(IIF(FN_MBR IS NOT NULL AND FN_MBR<>'',1,NULL)) AS TotalMBR,
                COUNT(IIF(FN_BT IS NOT NULL AND FN_BT<>'',1,NULL)) AS TotalBT,
                COUNT(IIF(FN_VBT IS NOT NULL AND FN_VBT<>'',1,NULL)) AS TotalVBT,
                COUNT(IIF(FN_MBT IS NOT NULL AND FN_MBT<>'',1,NULL)) AS TotalMBT,
                COUNT(IIF(FN_HZ1 IS NOT NULL AND FN_HZ1<>'',1,NULL)) AS TotalZ1,
                COUNT(IIF(FN_HZ2 IS NOT NULL AND FN_HZ2<>'',1,NULL)) AS TotalZ2,
                COUNT(IIF(FN_HZ3 IS NOT NULL AND FN_HZ3<>'',1,NULL)) AS TotalZ3,
                COUNT(IIF(FN_HZ4 IS NOT NULL AND FN_HZ4<>'',1,NULL)) AS TotalZ4
            FROM tblVKH;"
        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt

    End Function

    Public Function StatisicsVornamen() As Integer
        Dim strSQL As String =
            "SELECT COUNT(*) FROM tblVorname"
        Dim count As Integer

        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                count = CInt(cmd.ExecuteScalar())
            End Using
        End Using
        Return count
    End Function

    Public Function StatisicsNachnamen() As Integer
        Dim strSQL As String =
            "SELECT COUNT(*) FROM tblNachname"
        Dim count As Integer

        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                count = CInt(cmd.ExecuteScalar())
            End Using
        End Using
        Return count
    End Function

    Public Function StatisicsOrte() As Integer
        Dim strSQL As String =
            "SELECT COUNT(*) FROM tblOrt"
        Dim count As Integer

        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                count = CInt(cmd.ExecuteScalar())
            End Using
        End Using
        Return count
    End Function

    Public Function StatisicsPersonen() As Integer
        Dim strSQL As String =
            "SELECT COUNT(*) FROM tblPerson"
        Dim count As Integer

        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                count = CInt(cmd.ExecuteScalar())
            End Using
        End Using
        Return count
    End Function

    Public Function StatisicsFamilien() As Integer
        Dim strSQL As String =
            "SELECT COUNT(*) FROM tblFamilie"
        Dim count As Integer

        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                count = CInt(cmd.ExecuteScalar())
            End Using
        End Using
        Return count
    End Function

    Public Function VKH_Personen() As DataTable
        Return VKH_Personen("", 1)
    End Function

    Public Function VKH_Personen(Filter As String) As DataTable
        Return VKH_Personen(Filter, 1)
    End Function

    Public Function VKH_Personen(Order As Integer) As DataTable
        Return VKH_Personen("", Order)
    End Function

    Public Function VKH_Personen(Filter As String, Order As Integer) As DataTable
        Dim SQLFilter As String = ""
        Dim SQLOrder As String = "ORDER BY Nachname, Vorname, SEITE_H, NR_H"
        If Filter.Trim <> "" Then
            SQLFilter = " WHERE Person = ? "
        End If
        Select Case Order
            Case 1
                SQLOrder = "ORDER BY Nachname, Vorname, SEITE_H, NR_H"
            Case 2
                SQLOrder = "ORDER BY Nachname, Vorname, NR_H"
            Case 3
                SQLOrder = "ORDER BY Nachname, NR_H"
        End Select
        Dim strSQL As String = String.Format(
            "SELECT * FROM (SELECT
                tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum,
                VN_BR AS Vorname,
                FN_BR AS Nachname,
                W_BR AS Wohnort,
                H_BR AS Heimatort,
                Z_BR AS Bemerkung,
                'Bräutigam' AS Person
            FROM
                tblVKH
            WHERE
                 Len(VN_BR & '') > 0
                 OR Len(FN_BR & '') > 0
                 OR Len(Z_BR & '') > 0
            UNION
            SELECT
                tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum,
                VN_VBR AS Vorname,
                FN_VBR AS Nachname,
                W_EBR AS Wohnort,
                '' AS Heimatort,
                Z_VBR AS Bemerkung,
                'Vater Bräutigam' AS Person
            FROM
                tblVKH
            WHERE
                 Len(VN_VBR & '') > 0
                 OR Len(FN_VBR & '') > 0
                 OR Len(Z_VBR & '') > 0
            UNION
            SELECT
                tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum,
                VN_MBR AS Vorname,
                FN_MBR AS Nachname,
                W_EBR AS Wohnort,
                '' AS Heimatort,
                Z_MBR AS Bemerkung,
                'Mutter Bräutigam' AS Person
            FROM
                tblVKH
            WHERE
                 Len(VN_MBR & '') > 0
                 OR Len(FN_MBR & '') > 0
                 OR Len(Z_MBR & '') > 0
            UNION
            SELECT
                tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum,
                VN_BT AS Vorname,
                FN_BT AS Nachname,
                W_BT AS Wohnort,
                H_BT AS Heimatort,
                Z_BT AS Bemerkung,
                'Braut' AS Person
            FROM
                tblVKH
            WHERE
                 Len(VN_BT & '') > 0
                 OR Len(FN_BT & '') > 0
                 OR Len(Z_BT & '') > 0
            UNION
            SELECT
                tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum,
                VN_VBT AS Vorname,
                FN_VBT AS Nachname,
                W_EBT AS Wohnort,
                '' AS Heimatort,
                Z_VBT AS Bemerkung,
                'Vater Braut' AS Person
            FROM
                tblVKH
            WHERE
                 Len(VN_VBT & '') > 0
                 OR Len(FN_VBT & '') > 0
                 OR Len(Z_VBT & '') > 0
            UNION
            SELECT
                tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum,
                VN_MBT AS Vorname,
                FN_MBT AS Nachname,
                W_EBT AS Wohnort,
                '' AS Heimatort,
                Z_MBT AS Bemerkung,
                'Mutter Braut' AS Person
            FROM
                tblVKH
            WHERE
                 Len(VN_MBT & '') > 0
                 OR Len(FN_MBT & '') > 0
                 OR Len(Z_MBT & '') > 0
            UNION
            SELECT tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum, VN_HZ1 AS Vorname, FN_HZ1 AS Nachname, '' AS Wohnort, '' AS Heimatort, Z_HZ1 AS Bemerkung, IIf(G_HZ1='m','Zeuge','Zeugin') AS Person
            FROM tblVKH
            WHERE
                 Len(VN_HZ1 & '') > 0
                 OR Len(FN_HZ1 & '') > 0
                 OR Len(Z_HZ1 & '') > 0
            UNION
            SELECT tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum, VN_HZ2 AS Vorname, FN_HZ2 AS Nachname, '' AS Wohnort, '' AS Heimatort, Z_HZ2 AS Bemerkung, IIf(G_HZ2='m','Zeuge','Zeugin') AS Person
            FROM tblVKH
            WHERE
                 Len(VN_HZ2 & '') > 0
                 OR Len(FN_HZ2 & '') > 0
                 OR Len(Z_HZ2 & '') > 0
            UNION
            SELECT tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum, VN_HZ3 AS Vorname, FN_HZ3 AS Nachname, '' AS Wohnort, '' AS Heimatort, Z_HZ3 AS Bemerkung, IIf(G_HZ3='m','Zeuge','Zeugin') AS Person
            FROM tblVKH
            WHERE
                 Len(VN_HZ3 & '') > 0
                 OR Len(FN_HZ3 & '') > 0
                 OR Len(Z_HZ3 & '') > 0
            UNION
            SELECT tblVKHID, BUCH_H, SEITE_H, NR_H, HDatum, VN_HZ4 AS Vorname, FN_HZ4 AS Nachname, '' AS Wohnort, '' AS Heimatort, Z_HZ4 AS Bemerkung, IIf(G_HZ4='m','Zeuge','Zeugin') AS Person
            FROM tblVKH
            WHERE
                 Len(VN_HZ4 & '') > 0
                 OR Len(FN_HZ4 & '') > 0
                 OR Len(Z_HZ4 & '') > 0)
            {0}
            {1}
            ;", SQLFilter, SQLOrder)

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                If Filter.Trim <> "" Then
                    cmd.Parameters.AddWithValue("@Person", Filter.Trim)
                End If

                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Public Function StatisticsVKHYears() As DataTable
        Dim strSQL As String =
            "SELECT Left(NR_H,4) AS Jahr, Count(NR_H) AS Anzahl, BUCH_H
                FROM tblVKH
                GROUP BY Left(NR_H,4), BUCH_H;"

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Public Function StatisticsVKHPages() As DataTable
        Dim strSQL As String =
            "SELECT SEITE_H, Count(tblVKHID) AS Anzahl, BUCH_H, ReferenceDetails
                FROM tblVKH
                GROUP BY SEITE_H, BUCH_H, ReferenceDetails;"

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Public Function StatisticsVKHLocations() As DataTable
        Dim strSQL As String =
            "SELECT Ort, SUM(Anzahl) AS Total FROM 
                (SELECT
                    W_BR AS Ort,
                    Count(W_BR) AS Anzahl
                FROM
                    tblVKH
                GROUP BY
                    W_BR
                HAVING
                    W_BR <> ''
                UNION
                SELECT
                    H_BR AS Ort,
                    Count(H_BR) AS Anzahl
                FROM
                    tblVKH
                GROUP BY
                    H_BR
                HAVING
                    H_BR <> ''
                UNION
                SELECT
                    W_EBR AS Ort,
                    Count(W_EBR) AS Anzahl
                FROM
                    tblVKH
                GROUP BY
                    W_EBR
                HAVING
                    W_EBR <> ''
                UNION
                SELECT
                    W_BT AS Ort,
                    Count(W_BT) AS Anzahl
                FROM
                    tblVKH
                GROUP BY
                    W_BT
                HAVING
                    W_BT <> ''
                UNION
                SELECT
                    H_BT AS Ort,
                    Count(H_BT) AS Anzahl
                FROM
                    tblVKH
                GROUP BY
                    H_BT
                HAVING
                    H_BT <> ''
                UNION
                SELECT
                    W_EBT AS Ort,
                    Count(W_EBT) AS Anzahl
                FROM
                    tblVKH
                GROUP BY
                    W_EBT
                HAVING
                    W_EBT <> ''
                ) 
                GROUP BY Ort;"

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function


    Public Function StatisticsVKHLocations(Ort As String, Buch As String) As DataTable
        Dim Filter As String = ""
        If Buch.Trim <> "" Then
            Filter = " AND BUCH_H = ? "
        End If
        Dim strSQL As String =
            String.Format("SELECT
                Ort, Count(*) as Anzahl
            FROM
                (
                SELECT
                    BUCH_H, NR_H, Ort, Count(*) AS Anzahl
                FROM
                (
                    SELECT BUCH_H, NR_H, W_BR  AS Ort FROM tblVKH WHERE W_BR  <> ''
                    UNION ALL
                    SELECT BUCH_H, NR_H, H_BR  AS Ort FROM tblVKH WHERE H_BR  <> ''
                    UNION ALL
                    SELECT BUCH_H, NR_H, W_EBR  AS Ort FROM tblVKH WHERE W_EBR  <> ''
                    UNION ALL
                    SELECT BUCH_H, NR_H, W_BT AS Ort FROM tblVKH WHERE W_BT <> ''
                    UNION ALL
                    SELECT BUCH_H, NR_H, H_BT AS Ort FROM tblVKH WHERE H_BT <> ''
                    UNION ALL
                    SELECT BUCH_H, NR_H, W_EBT AS Ort FROM tblVKH WHERE W_EBT <> ''
                    UNION ALL
                    SELECT BUCH_H, NR_H, K_Ort AS Ort FROM tblVKH WHERE K_Ort <> ''
                )
                GROUP BY BUCH_H, NR_H, Ort
                HAVING Ort <> ? {0}
                ORDER BY BUCH_H, NR_H, Ort
                )
            GROUP BY Ort
            ORDER BY
                Count(*) DESC,
                Ort;", Filter)

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)

                cmd.Parameters.AddWithValue("@Ort", Ort)
                If Filter.Trim <> "" Then
                    cmd.Parameters.AddWithValue("@BUCH_H", Buch)
                End If

                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Public Function StatisticsVKHLocationsExtern(Ort As String, Buch As String) As DataTable
        Dim Filter As String = ""
        If Buch.Trim <> "" Then
            Filter = " AND BUCH_H = ? "
        End If
        Dim strSQL As String =
            String.Format("
                SELECT BUCH_H, NR_H, K_Ort AS Ort 
                FROM tblVKH 
                WHERE K_Ort <> '' AND K_Ort <> ? {0}
                ORDER BY K_Ort, NR_H;", Filter)

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)

                cmd.Parameters.AddWithValue("@Ort", Ort)
                If Filter.Trim <> "" Then
                    cmd.Parameters.AddWithValue("@BUCH_H", Buch)
                End If

                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Public Function GetNachname() As DataTable
        Dim strSQL As String =
            "SELECT * FROM tblNachname ORDER BY Nachname"

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Public Function UpdateNachname(ID As Integer, Nachname As String) As DataTable
        Dim strSQL As String =
            "UPDATE tblNachname SET Nachname = ? WHERE tblNachnameID = ?"

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)

                cmd.Parameters.AddWithValue("@Nachname", Nachname)
                cmd.Parameters.AddWithValue("@ID", ID)
                cmd.ExecuteNonQuery()

            End Using
        End Using
        Return dt
    End Function

    Public Function CleanDate(ByRef dateString As String) As Boolean
        dateString = dateString.Trim.Replace(",", ".")
        dateString = dateString.Replace("-", ".")
        dateString = dateString.Replace("_", ".")
        dateString = dateString.Replace(";", ".")
        dateString = dateString.Replace(" ", "")
        If IsDate(dateString) Then
            Dim d As Date = CDate(dateString)
            dateString = d.ToString("dd.MM.yyyy")
            Return True
        End If
        If dateString = vbNullString Then
            Return True
        End If
        Return False
    End Function

    Public Function VKH_ReportData() As DataTable
        Return VKH_ReportData(False, "")
    End Function

    Public Function VKH_ReportData(Filter As Boolean) As DataTable
        Return VKH_ReportData(Filter, "")
    End Function

    Public Function VKH_ReportData(Book As String) As DataTable
        Return VKH_ReportData(False, Book)
    End Function

    Public Function VKH_ReportData(Filter As Boolean, book As String) As DataTable
        Dim SQLFilter As String = ""
        If Filter = True Then
            SQLFilter = " WHERE CheckNeeded = -1 "
        End If
        If book.Trim <> "" Then
            If SQLFilter.Trim <> "" Then
                SQLFilter &= " AND BUCH_H = ? "
            Else
                SQLFilter = " WHERE BUCH_H = ? "
            End If
        End If
        Dim strSQL As String = String.Format(
            "SELECT * FROM tblVKH
            {0}
            ORDER BY
                NR_H
            ;", SQLFilter)

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                If book.Trim <> "" Then
                    cmd.Parameters.AddWithValue("@BUCH_H", book.Trim)
                End If
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Public Function VKH_Orte() As DataTable
        Return VKH_Orte("", "")
    End Function

    Public Function VKH_Orte(Filter As String) As DataTable
        Return VKH_Orte(Filter, "")
    End Function
    Public Function VKH_Orte(Filter As String, Book As String) As DataTable
        Dim SQLFilter As String = ""
        If Filter.Trim <> "" Then
            SQLFilter = " WHERE Ort <> ? "
        End If
        If Book.Trim <> "" Then
            If SQLFilter.Trim <> "" Then
                SQLFilter &= " AND BUCH_H = ? "
            Else
                SQLFilter = " WHERE BUCH_H = ? "
            End If
        End If
        Dim strSQL As String = String.Format(
            "SELECT * FROM (SELECT
                SEITE_H,
                NR_H,
                W_BR AS Ort,
                BUCH_H
            FROM
                tblVKH
            WHERE
                W_BR IS NOT NULL AND W_BR <> ''
            UNION
            SELECT
                SEITE_H,
                NR_H,
                H_BR AS Ort,
                BUCH_H
            FROM
                tblVKH
            WHERE
                H_BR IS NOT NULL AND H_BR <> ''
            UNION
            SELECT
                SEITE_H,
                NR_H,
                W_EBR AS Ort,
                BUCH_H
            FROM
                tblVKH
            WHERE
                W_EBR IS NOT NULL AND W_EBR <> ''
            UNION
            SELECT
                SEITE_H,
                NR_H,
                W_BT AS Ort,
                BUCH_H
            FROM
                tblVKH
            WHERE
                W_BT IS NOT NULL AND W_BT <> ''
            UNION
            SELECT
                SEITE_H,
                NR_H,
                H_BT AS Ort,
                BUCH_H
            FROM
                tblVKH
            WHERE
                H_BT IS NOT NULL AND H_BT <> ''
            UNION
            SELECT
                SEITE_H,
                NR_H,
                W_EBT AS Ort,
                BUCH_H
            FROM
                tblVKH
            WHERE
                W_EBT IS NOT NULL AND W_EBT <> ''
            UNION
            SELECT
                SEITE_H,
                NR_H,
                K_Ort AS Ort,
                BUCH_H
            FROM
                tblVKH
            WHERE
                K_Ort IS NOT NULL AND K_Ort <> ''
            )
            {0}
            Order BY Ort, NR_H
            ;", SQLFilter)

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                If Filter.Trim <> "" Then
                    cmd.Parameters.AddWithValue("@Ort", Filter.Trim)
                End If
                If Book.Trim <> "" Then
                    cmd.Parameters.AddWithValue("@BUCH_H", Book.Trim)
                End If

                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Function ToTitleCase(text As String) As String

        text = Regex.Replace(text.Trim(), "\s+", " ")

        If String.IsNullOrWhiteSpace(text) Then Return text

        Dim culture As CultureInfo = CultureInfo.CurrentCulture
        Return culture.TextInfo.ToTitleCase(text.ToLower())
    End Function

    Function GetVKH_Books() As DataTable
        Dim strSQL As String =
            "SELECT DISTINCT BUCH_H FROM tblVKH ORDER BY BUCH_H"
        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Function GetSearchPersons() As DataTable
        Dim strSQL As String =
            "SELECT
                p.tblPersonID,
                p.PS,
                n.Nachname,
                p.Vorname,
                p.Sex,
                IIf([p].[tblFamilieID] > 0, ""X"", """") AS Kind,
                p.Info,
                p.FSID,
                e.Datum,
                e.DatumText,
                e.Referenz,
                e.EreignisArt
            FROM
                (tblPerson AS p
                    LEFT JOIN tblNachname AS n
                        ON p.tblNachnameID = n.tblNachnameID)
                LEFT JOIN
                    (
                        SELECT
                            tblEreignis.tblPersonID,
                            First(tblEreignis.Datum) AS Datum,
                            First(tblEreignis.DatumText) AS DatumText,
                            First(tblEreignis.Referenz) AS Referenz,
                            First(tblEreignisArt.EreignisArt) AS EreignisArt
                        FROM
                            tblEreignis
                            INNER JOIN tblEreignisArt
                                ON tblEreignis.tblEreignisArtID = tblEreignisArt.tblEreignisArtID
                        GROUP BY
                            tblEreignis.tblPersonID
                    ) AS e
                ON p.tblPersonID = e.tblPersonID
            ORDER BY
                p.PS;"

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function

    Function GetFamilies() As DataTable
        Dim strSQL As String =
            "SELECT
                tblFamilie.tblFamilieID,
                tblFamilie.FS,
                tblFamilie.tblPersonIDV,
                tblFamilie.tblPersonIDM,
                [qryPerson]![Vorname] & ' ' & UCase([qryPerson]![Nachname]) AS Vater, 
                m.Vorname & ' ' & UCase(m.Nachname) AS Mutter
            FROM qryPerson As m RIGHT Join (qryPerson RIGHT Join tblFamilie On qryPerson.tblPersonID = tblFamilie.tblPersonIDV) ON m.tblPersonID = tblFamilie.tblPersonIDM
            ORDER BY FS;"

        Dim dt As New DataTable()
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                Using adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dt)
                End Using
            End Using
        End Using
        Return dt
    End Function
End Class
