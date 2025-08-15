Imports System.Data
Imports System.Data.OleDb
Imports System.Security.Cryptography

Public Class personen
    Private connectionString As String =
   String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)

    Private dt As New DataTable()
    Private dtE As New DataTable()

    Private isNewRecord As Boolean = False
    Private ID As Integer? = Nothing
    Private FID As Integer
    Private VID As Integer
    Private MID As Integer

    Private NT As String = "____"
    Private VT As String = "____"
    Private DaT As String = "0000"

    Private pSQL As String = "SELECT
            tblPerson.tblPersonID,
            tblPerson.tblFamilieID,
            tblPerson.tblNachnameID,
            tblPerson.tblKonfessionID,
            tblPerson.PS,
            tblNachname.Nachname, 
            tblPerson.Vorname,
            tblPerson.Sex, 
            tblPerson.Info,
            tblPerson.FSID
        FROM
            tblPerson
            LEFT JOIN tblNachname ON tblPerson.tblNachnameID = tblNachname.tblNachnameID"

    Public Sub New()
        InitializeComponent()

        LoadData()
        LoadKonfessionListe()

        If My.Settings.LastPID > 0 Then
            ID = My.Settings.LastPID
            FillPerson(ID)
        Else
            btnNew_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        Dim rowView As DataRowView = CType(dgPersonen.SelectedItem, DataRowView)
        Dim VID As Int16 = VornameAnlegen(txtVorname.Text, ID)
        If VID = -1 Then Exit Sub

        Dim NID As Int16 = NachnamenID(txtNachname.Text)
        If NID = -1 Then Exit Sub
        If txtNachname.Text.Trim() <> "" Then
            NT = txtNachname.Text.Trim().Substring(0, 4)
        End If

        If rowView IsNot Nothing Or isNewRecord Then
            If ID > 0 Then
                FindYear()
            End If
            txtPS.Text = (NT & VT & DaT).ToUpper

            Try
                Using conn As New OleDbConnection(connectionString)
                    conn.Open()
                    If isNewRecord Then
                        Dim insertCmd As New OleDbCommand("INSERT INTO tblPerson (PS, Sex, FSID, tblFamilieID, tblNachnameID, tblKonfessionID, Vorname, Info) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", conn)
                        insertCmd.Parameters.AddWithValue("@PS", txtPS.Text.ToUpper)
                        insertCmd.Parameters.AddWithValue("@Sex", CType(cbSex.SelectedItem, ComboBoxItem).Content.ToString())
                        insertCmd.Parameters.AddWithValue("@FSID", "")
                        insertCmd.Parameters.AddWithValue("@tblFamilieID", 0)
                        insertCmd.Parameters.AddWithValue("@tblNachnameID", NID)
                        insertCmd.Parameters.AddWithValue("@tblKonfessionID", cbKonfession.SelectedValue)
                        insertCmd.Parameters.AddWithValue("@Vorname", txtVorname.Text)
                        insertCmd.Parameters.AddWithValue("@Info", txtInfo.Text)
                        insertCmd.ExecuteNonQuery()

                        Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                            ID = Convert.ToInt32(cmdId.ExecuteScalar())
                        End Using

                        VornameAnlegen(txtVorname.Text, ID)

                        MessageBox.Show("Neuer Datensatz gespeichert!")
                    ElseIf ID.HasValue Then
                        rowView("Vorname") = txtVorname.Text
                        rowView("tblKonfessionID") = cbKonfession.SelectedValue
                        rowView("Sex") = CType(cbSex.SelectedItem, ComboBoxItem).Content.ToString()
                        rowView("Info") = txtInfo.Text
                        Dim updateCmd As New OleDbCommand("UPDATE tblPerson SET PS = ?, Sex = ?, FSID=  ?, tblNachnameID = ?, tblKonfessionID = ?, Vorname = ?, Info = ? WHERE tblPersonID = ?", conn)
                        updateCmd.Parameters.AddWithValue("@PS", txtPS.Text)
                        updateCmd.Parameters.AddWithValue("@Sex", CType(cbSex.SelectedItem, ComboBoxItem).Content.ToString())
                        updateCmd.Parameters.AddWithValue("@FSID", "")
                        updateCmd.Parameters.AddWithValue("@tblNachnameID", NID)
                        updateCmd.Parameters.AddWithValue("@tblKonfessionID", cbKonfession.SelectedValue)
                        updateCmd.Parameters.AddWithValue("@Vorname", txtVorname.Text)
                        updateCmd.Parameters.AddWithValue("@Info", txtInfo.Text)
                        updateCmd.Parameters.AddWithValue("@ID", ID)
                        updateCmd.ExecuteNonQuery()

                        MessageBox.Show("Änderungen gespeichert!")
                    End If
                End Using

                LoadData() ' DataGrid aktualisieren
            Catch ex As Exception
                MessageBox.Show("Fehler beim Speichern: " & ex.Message)
            End Try
            My.Settings.LastPID = ID
            My.Settings.Save()
            btnNewEvent.IsEnabled = ID IsNot Nothing
        End If
    End Sub

    Private Sub LoadData()
        Dim strSQL As String = pSQL & " ORDER BY PS"


        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand(strSQL, conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                dt.Clear()
                adapter.Fill(dt)
            End Using

            dgPersonen.ItemsSource = dt.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler: " & ex.Message)
        End Try
    End Sub

    Private Sub btnNew_Click(sender As Object, e As RoutedEventArgs) Handles btnNew.Click
        txtVorname.Clear()
        txtNachname.Clear()
        cbKonfession.SelectedValue = 1
        cbSex.SelectedIndex = 0
        NT = "____"
        VT = "____"
        DaT = "0000"
        txtPS.Text = (NT & VT & DaT).ToUpper
        txtInfo.Text = ""
        txtFSID.Text = ""
        ID = Nothing
        isNewRecord = True
        txtVorname.Focus()
        LoadEventData()
        AdditionalContent.Content = Nothing
    End Sub


    Private Sub dgPersonen_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgPersonen.AutoGeneratedColumns
        For Each col In dgPersonen.Columns
            If col.Header IsNot Nothing Then
                Select Case col.Header.ToString()
                    Case "tblPersonID", "tblFamilieID", "tblNachnameID", "tblKonfessionID"
                        col.Visibility = Visibility.Collapsed
                End Select
            End If
        Next

        dgPersonen.IsReadOnly = True
        dgPersonen.CanUserAddRows = False
        dgPersonen.CanUserDeleteRows = False
    End Sub

    Private Sub dgPersonen_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim rowView As DataRowView = CType(dgPersonen.SelectedItem, DataRowView)
        If rowView IsNot Nothing Then
            ID = Convert.ToInt32(rowView("tblPersonID"))
            FillPerson(ID)
        End If

    End Sub

    Private Sub cbKonfession_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbKonfession.SelectionChanged
        If cbKonfession.SelectedValue IsNot Nothing Then
            Dim selectedID As Integer = CInt(cbKonfession.SelectedValue)
        End If
    End Sub

    Private Sub LoadKonfessionListe()
        Try
            Dim dt As New DataTable()
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT tblKonfessionID, Konfessionkurz FROM tblKonfession ORDER BY Konfessionkurz", conn)

                Dim adapter As New OleDbDataAdapter(cmd)
                adapter.Fill(dt)
            End Using

            cbKonfession.ItemsSource = dt.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler beim Laden der Konfession: " & ex.Message)
        End Try
    End Sub

    Private Function NachnamenID(Nachname As String) As Int16
        Dim id As Integer = -1
        Dim sqlSelect As String = "SELECT tblNachnameID FROM tblNachname WHERE Nachname = ?"
        Dim sqlInsert As String = "INSERT INTO tblNachname (Nachname) VALUES (?)"
        If Trim(Nachname) = "" Then
            Return 0
        End If
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(sqlSelect, conn)
                cmd.Parameters.AddWithValue("@Nachname", Nachname)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    id = Convert.ToInt32(result)
                    Return id
                End If
            End Using

            If MessageBox.Show(String.Format("Soll der Nachname '{0}' angelegt werden?", txtNachname.Text), "Nachname anlegen", MessageBoxButton.YesNo) = MessageBoxResult.No Then
                Return -1
            End If
            Using cmdInsert As New OleDbCommand(sqlInsert, conn)
                cmdInsert.Parameters.AddWithValue("@Nachname", Nachname)
                cmdInsert.ExecuteNonQuery()
            End Using

            Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                id = Convert.ToInt32(cmdId.ExecuteScalar())
            End Using
        End Using
        Return id
    End Function
    Private Function VornamenID(Vorname As String) As Int16
        Dim id As Integer = -1
        Dim sqlSelect As String = "SELECT tblVornameID FROM tblVorname WHERE Vorname = ?"
        Dim sqlInsert As String = "INSERT INTO tblVorname (Vorname) VALUES (?)"
        If Trim(Vorname) = "" Then
            Return 0
        End If
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(sqlSelect, conn)
                cmd.Parameters.AddWithValue("@Vorname", Vorname)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    id = Convert.ToInt32(result)
                    Return id
                End If
            End Using

            If MessageBox.Show(String.Format("Soll der Vorname '{0}' angelegt werden?", Vorname), "Vorname anlegen", MessageBoxButton.YesNo) = MessageBoxResult.No Then
                Return -1
            End If
            Using cmdInsert As New OleDbCommand(sqlInsert, conn)
                cmdInsert.Parameters.AddWithValue("@Vorname", Vorname)
                cmdInsert.ExecuteNonQuery()
            End Using

            Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                id = Convert.ToInt32(cmdId.ExecuteScalar())
            End Using
        End Using
        Return id
    End Function

    Private Function VornameAnlegen(Vornamen As String, PID As Int16?) As Int16
        'TODO Bindestrich prüfen
        If Trim(Vornamen) = "" Then
            Return 0
        End If

        Dim Vorname() As String = Vornamen.Trim.Split(" "c)
        Dim Reihenfolge As Int16 = 1
        For Each VN In Vorname
            Dim Zeichen As String = ""
            Dim VName As String = VN
            If Reihenfolge = 1 Then
                VT = VName.Substring(0, 4)
            End If
            If VName.StartsWith("*") Then
                Zeichen = "*"
                VName = VName.Substring(1)
                VT = VName.Substring(0, 4)
            End If
            Dim VID As Int16 = VornamenID(VName)
            If VID = -1 Then Return VID
            If PID > 0 Then
                VornameSpeichern(VID, PID, Reihenfolge, Zeichen)
            End If
            Reihenfolge += 1
        Next
        Return 1
    End Function

    Private Function VornameSpeichern(VornameID As Int16, PersonID As Int16, Reihenfolge As Int16, Zeichen As String) As Int16
        Dim sqlFind As String = "SELECT tblPVornameID FROM  tblPVorname  WHERE tblVornameID = ? AND  tblPersonID = ?"
        Dim sqlInsert As String = "INSERT INTO tblPVorname (tblVornameID, tblPersonID, Reihenfolge, Zeichen) VALUES (?, ?, ?, ?)"
        Dim sqlUpdate As String = "UPDATE tblPVorname SET tblVornameID = ?, tblPersonID = ?, Reihenfolge = ?, Zeichen = ? WHERE tblPVornameID = ?"
        Dim PVID As Int16 = 0
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(sqlFind, conn)
                cmd.Parameters.AddWithValue("@tblVornameID", VornameID)
                cmd.Parameters.AddWithValue("@tblPersonID", PersonID)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    PVID = Convert.ToInt32(result)
                End If
            End Using
            If PVID = 0 Then

                Using cmdInsert As New OleDbCommand(sqlInsert, conn)
                    cmdInsert.Parameters.AddWithValue("@tblVornameID", VornameID)
                    cmdInsert.Parameters.AddWithValue("@tblPersonID", PersonID)
                    cmdInsert.Parameters.AddWithValue("@Reihenfolge", Reihenfolge)
                    cmdInsert.Parameters.AddWithValue("@Zeichen", Zeichen)
                    cmdInsert.ExecuteNonQuery()
                End Using

                Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                    PVID = Convert.ToInt32(cmdId.ExecuteScalar())
                End Using
            Else
                Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                    cmdUpdate.Parameters.AddWithValue("@tblVornameID", VornameID)
                    cmdUpdate.Parameters.AddWithValue("@tblPersonID", PersonID)
                    cmdUpdate.Parameters.AddWithValue("@Reihenfolge", Reihenfolge)
                    cmdUpdate.Parameters.AddWithValue("@Zeichen", Zeichen)
                    cmdUpdate.Parameters.AddWithValue("@ID", PVID)
                    cmdUpdate.ExecuteNonQuery()
                End Using

            End If

        End Using

        Return PVID
    End Function

    Private Sub LoadEventData()
        Dim strSQL As String = "SELECT
                tblEreignis.tblEreignisID,
                tblEreignisArt.EreignisArt AS Ereignis,
                tblEreignis.DatumText AS Datum,
                tblEreignis.Datum AS HDatum,
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
                tblEreignis.Datum;"


        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("@PersonID", IIf(ID Is Nothing, 0, ID))
                Dim adapter As New OleDbDataAdapter(cmd)
                dtE.Clear()
                adapter.Fill(dtE)
            End Using

            dgEreignis.ItemsSource = dtE.DefaultView
            btnNewEvent.IsEnabled = ID IsNot Nothing
        Catch ex As Exception
            MessageBox.Show("Fehler: " & ex.Message)
        End Try
    End Sub

    Private Sub btnNewEvent_Click(sender As Object, e As RoutedEventArgs) Handles btnNewEvent.Click
        If ID Is Nothing Then
            MessageBox.Show("Der Datensatz muss zuerst gespeichert werden, bevor ein Ereignis angelegt werden kann.")
            Exit Sub
        End If
        Dim details = New ereignis()
        details.PersonId = ID
        details.FamilieId = 0
        AddHandler details.DataSaved, AddressOf OnDatenGespeichert
        AdditionalContent.Content = details
    End Sub

    Private Sub dgEreignis_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgEreignis.AutoGeneratedColumns
        For Each col In dgEreignis.Columns
            If col.Header IsNot Nothing Then
                Select Case col.Header.ToString()
                    Case "tblEreignisID", "HDatum"
                        col.Visibility = Visibility.Collapsed
                End Select
            End If
        Next
        dgEreignis.IsReadOnly = True
        dgEreignis.CanUserAddRows = False
        dgEreignis.CanUserDeleteRows = False
    End Sub

    Private Sub dgEreignis_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles dgEreignis.MouseDoubleClick
        Dim rowView As DataRowView = CType(dgEreignis.SelectedItem, DataRowView)
        If rowView IsNot Nothing Then

            Dim details = New ereignis()
            details.EintragId = Convert.ToInt32(rowView("tblEreignisID"))
            AddHandler details.DataSaved, AddressOf OnDatenGespeichert
            AdditionalContent.Content = details
        End If
    End Sub

    Private Sub OnDatenGespeichert(sender As Object, e As EventArgs)
        LoadEventData()
        PSSpeichern()
    End Sub

    Private Sub FindYear()
        Dim Jahr As Object = Nothing
        Dim strSQL As String = " SELECT
                Min(Year([Datum])) AS Jahr
            FROM
                tblEreignis
                INNER JOIN tblEreignisArt ON tblEreignis.tblEreignisArtID = tblEreignisArt.tblEreignisArtID
            WHERE
                tblEreignisArt.Reihenfolge < 10
            GROUP BY
                tblEreignis.tblPersonID
            HAVING
                tblEreignis.tblPersonID = ?"
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("?", ID) ' personID = deine gesuchte Person
                Jahr = cmd.ExecuteScalar()
            End Using
        End Using

        If Jahr IsNot Nothing AndAlso Not IsDBNull(Jahr) Then
            DaT = Jahr
        Else
            DaT = "0000"
        End If
    End Sub
    Private Sub PSSpeichern()
        Dim sqlUpdate As String = "UPDATE tblPerson SET PS = ? WHERE tblPersonID = ?"
        FindYear()
        txtPS.Text = (NT & VT & DaT).ToUpper
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                cmdUpdate.Parameters.AddWithValue("@PS", txtPS.Text)
                cmdUpdate.Parameters.AddWithValue("@tblPersonID", ID)
                cmdUpdate.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub btnFamile_Click(sender As Object, e As RoutedEventArgs)
        My.Settings.LastFID = FID
        My.Settings.Save()
        Dim mw = TryCast(Window.GetWindow(Me), MainWindow)
        If mw IsNot Nothing Then
            mw.ShowContent(New familien())
        End If
    End Sub

    Private Sub btnMutter_Click(sender As Object, e As RoutedEventArgs)
        FillPerson(MID)
    End Sub

    Private Sub btnVater_Click(sender As Object, e As RoutedEventArgs)
        FillPerson(VID)
    End Sub

    Private Sub FillPerson(PID As Integer)
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Dim sql As String = pSQL & " WHERE tblPersonID = ?"
            Using cmd As New OleDbCommand(sql, conn)
                cmd.Parameters.AddWithValue("@p1", PID)

                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        ' Werte auslesen und prüfen auf DBNull
                        If Not IsDBNull(reader("PS")) Then
                            txtPS.Text = reader("PS")
                        End If

                        If Not IsDBNull(reader("Vorname")) Then
                            txtVorname.Text = reader("Vorname")
                        End If

                        If Not IsDBNull(reader("Nachname")) Then
                            txtNachname.Text = reader("Nachname")
                        End If

                        If Not IsDBNull(reader("tblKonfessionID")) Then
                            cbKonfession.SelectedValue = CInt(reader("tblKonfessionID"))
                        End If

                        If Not IsDBNull(reader("Sex")) Then
                            'cbSex.SelectedValue = reader("Sex")
                            For Each item As ComboBoxItem In cbSex.Items
                                If item.Content.ToString() = reader("Sex") Then
                                    cbSex.SelectedItem = item
                                    Exit For
                                End If
                            Next
                        End If

                        If Not IsDBNull(reader("FSID")) Then
                            txtFSID.Text = reader("FSID")
                        End If

                        If Not IsDBNull(reader("Info")) Then
                            txtInfo.Text = reader("Info")
                        End If

                        If Not IsDBNull(reader("tblFamilieID")) Then
                            FID = CInt(reader("tblFamilieID"))
                        End If
                    Else
                        MessageBox.Show("Keine Person mit dieser ID gefunden.")
                    End If
                End Using
            End Using
        End Using

        isNewRecord = False
        LoadEventData()
        AdditionalContent.Content = Nothing

        If FID > 0 Then
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim sql As String = "SELECT tblPersonIDV, tblPersonIDM FROM tblFamilie WHERE tblFamilieID = ?"
                Using cmd As New OleDbCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@p1", FID)

                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            ' Werte auslesen und prüfen auf DBNull
                            If Not IsDBNull(reader("tblPersonIDV")) Then
                                VID = CInt(reader("tblPersonIDV"))
                            End If

                            If Not IsDBNull(reader("tblPersonIDM")) Then
                                MID = CInt(reader("tblPersonIDM"))
                            End If
                        Else
                            MessageBox.Show("Keine Familie mit dieser ID gefunden.")
                        End If
                    End Using
                End Using
            End Using
        Else
            VID = 0
            MID = 0
        End If

        btnFamile.IsEnabled = FID > 0
        btnVater.IsEnabled = VID > 0
        btnMutter.IsEnabled = MID > 0
        btnNewEvent.IsEnabled = ID IsNot Nothing
    End Sub

End Class
