Imports System.Data
Imports System.Data.OleDb
Imports System.Security.Cryptography
Imports inoGenDLL

Public Class familien
    Private cGenDB As New ClsGenDB(My.Settings.DBPath)
    Private connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)


    Private dt As New DataTable()
    Private dtE As New DataTable()
    Private dtP As New DataTable()

    Private isNewRecord As Boolean = False
    Private ID As Integer? = Nothing
    Private VID As Integer? = Nothing
    Private MID As Integer? = Nothing

    Private VT As String = "____"
    Private MT As String = "____"
    Private DaT As String = "0000"

    Public Sub New()
        InitializeComponent()

        isNewRecord = True
        If My.Settings.LastFID > 0 Then
            ID = My.Settings.LastFID
            FindFamilieByID(ID)
            LoadFamily()
        End If

        LoadData()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        Dim rowView As DataRowView = CType(dgFamilien.SelectedItem, DataRowView)

        If rowView IsNot Nothing Or isNewRecord Then
            If ID > 0 Then
                FindYear()
            End If
            txtFS.Text = (VT & MT & DaT).ToUpper

            Try
                Using conn As New OleDbConnection(connectionString)
                    conn.Open()
                    If isNewRecord Then
                        Dim insertCmd As New OleDbCommand("INSERT INTO tblFamilie (FS, tblPersonIDV, tblPersonIDM) VALUES (?, ?, ?)", conn)
                        insertCmd.Parameters.AddWithValue("@FS", txtFS.Text.ToUpper)
                        If IsNothing(VID) Then
                            insertCmd.Parameters.AddWithValue("@tblPersonIDV", DBNull.Value)
                        Else
                            insertCmd.Parameters.AddWithValue("@tblPersonIDV", VID)
                        End If
                        If IsNothing(MID) Then
                            insertCmd.Parameters.AddWithValue("@tblPersonIDM", DBNull.Value)
                        Else
                            insertCmd.Parameters.AddWithValue("@tblPersonIDM", MID)
                        End If

                        insertCmd.ExecuteNonQuery()


                        Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                            ID = Convert.ToInt32(cmdId.ExecuteScalar())
                        End Using


                        MessageBox.Show("Neuer Datensatz gespeichert!")
                    ElseIf ID.HasValue Then
                        Dim updateCmd As New OleDbCommand("UPDATE tblFamilie SET FS = ?, tblPersonIDV = ?, tblPersonIDM=  ? WHERE tblFamilieID = ?", conn)
                        updateCmd.Parameters.AddWithValue("@FS", txtFS.Text.ToUpper)
                        If IsNothing(VID) Then
                            updateCmd.Parameters.AddWithValue("@tblPersonIDV", DBNull.Value)
                        Else
                            updateCmd.Parameters.AddWithValue("@tblPersonIDV", VID)
                        End If
                        If IsNothing(MID) Then
                            updateCmd.Parameters.AddWithValue("@tblPersonIDM", DBNull.Value)
                        Else
                            updateCmd.Parameters.AddWithValue("@tblPersonIDM", MID)
                        End If
                        updateCmd.Parameters.AddWithValue("@ID", ID)
                        updateCmd.ExecuteNonQuery()

                        MessageBox.Show("Änderungen gespeichert!")
                    End If
                End Using

                LoadData()
            Catch ex As Exception
                MessageBox.Show("Fehler beim Speichern: " & ex.Message)
            End Try
            My.Settings.LastFID = ID
            My.Settings.Save()
        End If
    End Sub

    Private Sub LoadData()
        Dim strSQL As String = "SELECT 
                tblFamilie.tblFamilieID, 
                tblFamilie.FS, 
                tblFamilie.tblPersonIDV, 
                tblFamilie.tblPersonIDM, 
                [qryPerson]![Vorname] & ' ' & UCase([qryPerson]![Nachname]) AS Vater, 
                m.Vorname & ' ' & UCase(m.Nachname) AS Mutter
            FROM qryPerson AS m RIGHT JOIN (qryPerson RIGHT JOIN tblFamilie ON qryPerson.tblPersonID = tblFamilie.tblPersonIDV) ON m.tblPersonID = tblFamilie.tblPersonIDM
            ORDER BY FS;"


        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand(strSQL, conn)
                Dim adapter As New OleDbDataAdapter(cmd)
                dt.Clear()
                adapter.Fill(dt)
            End Using

            dgFamilien.ItemsSource = dt.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler: " & ex.Message)
        End Try

        If ID.HasValue Then
            For Each rowView As DataRowView In dgFamilien.Items
                If CInt(rowView("tblFamilieID")) = ID Then
                    ' Selektion setzen
                    dgFamilien.SelectedItem = rowView

                    ' Sichtbar machen
                    dgFamilien.ScrollIntoView(rowView)

                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub btnNew_Click(sender As Object, e As RoutedEventArgs) Handles btnNew.Click
        VT = "____"
        MT = "____"
        DaT = "0000"
        ID = Nothing
        VID = Nothing
        MID = Nothing
        isNewRecord = True
        txtVater.Focus()
        txtFS.Text = (VT & MT & DaT).ToUpper
        txtVater.Clear()
        txtMutter.Clear()
        LoadEventData()
        LoadChildData()
        AdditionalContent.Content = Nothing
    End Sub


    Private Sub dgFamilien_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgFamilien.AutoGeneratedColumns
        For Each col In dgFamilien.Columns
            If col.Header IsNot Nothing Then
                Select Case col.Header.ToString()
                    Case "tblPersonID", "tblFamilieID", "tblPersonIDV", "tblPersonIDM"
                        col.Visibility = Visibility.Collapsed
                End Select
            End If
        Next

        dgFamilien.IsReadOnly = True
        dgFamilien.CanUserAddRows = False
        dgFamilien.CanUserDeleteRows = False
    End Sub

    Private Sub dgFamilien_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim rowView As DataRowView = CType(dgFamilien.SelectedItem, DataRowView)
        If rowView IsNot Nothing Then
            ID = Convert.ToInt32(rowView("tblFamilieID"))
            If Not IsDBNull(rowView("tblPersonIDV")) Then
                VID = Convert.ToInt32(rowView("tblPersonIDv"))
            Else
                VID = Nothing
            End If
            If Not IsDBNull(rowView("tblPersonIDM")) Then
                MID = Convert.ToInt32(rowView("tblPersonIDM"))
            Else
                MID = Nothing
            End If
            If Not IsDBNull(rowView("FS")) Then
                txtFS.Text = rowView("FS")
            End If
        End If

        LoadFamily()

    End Sub

    Private Sub LoadFamily()
        isNewRecord = False
        LoadEventData()
        LoadChildData()
        AdditionalContent.Content = Nothing
        If VID > 0 Then
            txtVater.Text = cGenDB.PersonenDaten(VID)
            VT = txtVater.Text.Substring(0, 4)
        Else
            VT = "____"
            txtVater.Text = ""
        End If
        If MID > 0 Then
            txtMutter.Text = cGenDB.PersonenDaten(MID)
            MT = txtMutter.Text.Substring(0, 4)
        Else
            txtMutter.Text = ""
            MT = "____"
        End If
    End Sub

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
            WHERE tblFamilieID = ? and tblPersonID = 0
            ORDER BY
                tblEreignisArt.Reihenfolge,
                tblEreignis.Datum;"


        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("@FamilieID", IIf(ID Is Nothing, 0, ID))
                Dim adapter As New OleDbDataAdapter(cmd)
                dtE.Clear()
                adapter.Fill(dtE)
            End Using

            dgEreignis.ItemsSource = dtE.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler: " & ex.Message)
        End Try
    End Sub

    Private Sub btnNewEvent_Click(sender As Object, e As RoutedEventArgs) Handles btnNewEvent.Click
        If ID Is Nothing Then
            MessageBox.Show("Der Datensatz muss zuerst gespeichert werden, bevor ein Ereignis angelegt werden kann.")
            Exit Sub
        End If
        Dim details = New ereignis(False)
        details.FamilieId = ID
        details.PersonId = 0
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

            Dim details = New ereignis(False)
            details.EintragId = Convert.ToInt32(rowView("tblEreignisID"))
            AddHandler details.DataSaved, AddressOf OnDatenGespeichert
            AdditionalContent.Content = details
        End If
    End Sub

    Private Sub OnDatenGespeichert(sender As Object, e As EventArgs)
        LoadEventData()
        FSSpeichern()
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
                tblEreignis.tblFamilieID
            HAVING
                tblEreignis.tblFamilieID = ?"
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("?", ID)
                Jahr = cmd.ExecuteScalar()
            End Using
        End Using

        If Jahr IsNot Nothing AndAlso Not IsDBNull(Jahr) Then
            DaT = Jahr
        Else
            DaT = "0000"
        End If
    End Sub
    Private Sub FSSpeichern()
        Dim sqlUpdate As String = "UPDATE tblFamilie SET FS = ? WHERE tblFamilieID = ?"
        FindYear()
        txtFS.Text = (VT & MT & DaT).ToUpper
        Using conn As New OleDbConnection(connectionString)
            conn.Open()
            Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                cmdUpdate.Parameters.AddWithValue("@FS", txtFS.Text)
                cmdUpdate.Parameters.AddWithValue("@tblFamilieID", ID)
                cmdUpdate.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Private Sub btnV_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New SuchePerson(True)
        ' Event-Handler für Rückgabe setzen
        AddHandler win.PersonSelected, Sub(id, persontext)
                                           MessageBox.Show("Ausgewählte Person: " & persontext)
                                           ' Hier kannst du direkt Felder befüllen
                                           VID = id
                                           txtVater.Text = cGenDB.PersonenDaten(VID)
                                           VT = txtVater.Text.Substring(0, 4)
                                           txtFS.Text = (VT & MT & DaT).ToUpper
                                       End Sub

        win.Show()
    End Sub

    Private Sub btnM_Click(sender As Object, e As RoutedEventArgs)
        Dim win As New SuchePerson(False)
        AddHandler win.PersonSelected, Sub(id, persontext)
                                           MessageBox.Show("Ausgewählte Person: " & persontext)

                                           MID = id
                                           txtMutter.Text = cGenDB.PersonenDaten(MID)
                                           MT = txtMutter.Text.Substring(0, 4)
                                           txtFS.Text = (VT & MT & DaT).ToUpper
                                       End Sub

        win.Show()
    End Sub

    Private Sub btnNewChild_Click(sender As Object, e As RoutedEventArgs)
        If ID Is Nothing Then
            MessageBox.Show("Der Datensatz muss zuerst gespeichert werden, bevor ein Kind zugeordnet werden kann.")
            Exit Sub
        End If
        Dim win As New SuchePerson(VT)
        AddHandler win.PersonSelected, Sub(pid, persontext)
                                           If pid = VID Or pid = MID Then
                                               MessageBox.Show("Die Person ist bereits als Vater oder Mutter zugeordnet.")
                                               Exit Sub
                                           End If

                                           Using conn As New OleDbConnection(connectionString)
                                               conn.Open()

                                               Dim sqlCheck As String = "SELECT tblFamilieID FROM tblPerson WHERE tblPersonID = ? AND tblFamilieID > 0 "
                                               Using cmdCheck As New OleDbCommand(sqlCheck, conn)
                                                   cmdCheck.Parameters.AddWithValue("@p1", pid)

                                                   Dim result = cmdCheck.ExecuteScalar()

                                                   If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                                                       MessageBox.Show(String.Format("Diese Person '{0}' ist bereits einer Familie zugeordnet (FamilieID={1}}).", persontext, result.ToString()))
                                                   Else

                                                       MessageBox.Show("Ausgewählte Person: " & persontext)
                                                       Dim sqlUpdate As String = "UPDATE tblPerson SET tblFamilieID = ? WHERE tblPersonID = ?"
                                                       Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                                                           cmdUpdate.Parameters.AddWithValue("@p1", ID)
                                                           cmdUpdate.Parameters.AddWithValue("@p2", pid)
                                                           cmdUpdate.ExecuteNonQuery()
                                                       End Using

                                                       MessageBox.Show("Familie wurde erfolgreich zugeordnet (FamilieID=" & ID & ").")
                                                   End If
                                               End Using
                                           End Using
                                           LoadChildData()
                                       End Sub

        win.Show()
    End Sub

    Private Sub LoadChildData()
        Dim strSQL As String = "SELECT
                tblPersonID,
                PS,
                Sex,
                Vorname
            FROM
                tblPerson
            WHERE tblFamilieID = ?
            ORDER BY Right(PS, 4);"

        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("@FamilieID", IIf(ID Is Nothing, -1, ID))
                Dim adapter As New OleDbDataAdapter(cmd)
                dtP.Clear()
                adapter.Fill(dtP)
            End Using

            dgPersonen.ItemsSource = dtP.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler: " & ex.Message)
        End Try
    End Sub

    Private Sub dgPersonen_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgFamilien.AutoGeneratedColumns
        For Each col In dgPersonen.Columns
            If col.Header IsNot Nothing Then
                Select Case col.Header.ToString()
                    Case "tblPersonID"
                        col.Visibility = Visibility.Collapsed
                End Select
            End If
        Next

        dgFamilien.IsReadOnly = True
        dgFamilien.CanUserAddRows = False
        dgFamilien.CanUserDeleteRows = False
    End Sub

    Private Sub dgPersonen_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim rowView As DataRowView = CType(dgPersonen.SelectedItem, DataRowView)
        If rowView IsNot Nothing Then
            My.Settings.LastPID = Convert.ToInt32(rowView("tblPersonID"))
            My.Settings.Save()
            Dim mw = TryCast(Window.GetWindow(Me), MainWindow)
            If mw IsNot Nothing Then
                mw.ShowContent(New personen())
            End If
        End If
    End Sub

    Private Sub FindFamilieByID(familieID As Integer)
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Dim sql As String = "SELECT * FROM tblFamilie WHERE tblFamilieID = ?"
            Using cmd As New OleDbCommand(sql, conn)
                cmd.Parameters.AddWithValue("@p1", familieID)

                Using reader As OleDbDataReader = cmd.ExecuteReader()
                    If reader.Read() Then
                        If Not IsDBNull(reader("FS")) Then
                            txtFS.Text = reader("FS")
                        End If

                        If Not IsDBNull(reader("tblPersonIDV")) Then
                            VID = CInt(reader("tblPersonIDV"))
                        End If

                        If Not IsDBNull(reader("tblPersonIDM")) Then
                            MID = CInt(reader("tblPersonIDM"))
                        End If
                    Else
                        MessageBox.Show("Keine Person mit dieser ID gefunden.")
                    End If
                End Using
            End Using
        End Using
    End Sub

    Private Sub txtVater_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtVater.MouseDoubleClick
        If VID > 0 Then
            My.Settings.LastPID = VID
            My.Settings.Save()
            Dim mw = TryCast(Window.GetWindow(Me), MainWindow)
            If mw IsNot Nothing Then
                mw.ShowContent(New personen())
            End If
        End If
    End Sub

    Private Sub txtMutter_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtMutter.MouseDoubleClick
        If MID > 0 Then
            My.Settings.LastPID = MID
            My.Settings.Save()
            Dim mw = TryCast(Window.GetWindow(Me), MainWindow)
            If mw IsNot Nothing Then
                mw.ShowContent(New personen())
            End If
        End If
    End Sub

    Private Sub CMenuItemOpen_Click(sender As Object, e As RoutedEventArgs)
        Dim row As DataRowView = CType(dgPersonen.SelectedItem, DataRowView)
        If row IsNot Nothing Then
            My.Settings.LastPID = Convert.ToInt32(row("tblPersonID").ToString())
            My.Settings.Save()
            Dim mw = TryCast(Window.GetWindow(Me), MainWindow)
            If mw IsNot Nothing Then
                mw.ShowContent(New personen())
            End If
        End If
    End Sub

    Private Sub CMenuItemDelete_Click(sender As Object, e As RoutedEventArgs)
        Dim row As DataRowView = CType(dgPersonen.SelectedItem, DataRowView)
        If row IsNot Nothing Then
            If MessageBox.Show("Lösche Person mit PS " & row("PS").ToString(), "Löschen bestätigen", MessageBoxButton.YesNo, MessageBoxImage.Warning) = MessageBoxResult.Yes Then
                Dim personId As Integer = Convert.ToInt32(row("tblPersonID").ToString())
                Using conn As New OleDbConnection(connectionString)
                    conn.Open()
                    Dim sqlDelete As String = "UPDATE tblPerson SET tblFamilieID = 0 WHERE tblPersonID = ?"
                    Using cmdDelete As New OleDbCommand(sqlDelete, conn)
                        cmdDelete.Parameters.AddWithValue("@tblPersonID", personId)
                        cmdDelete.ExecuteNonQuery()
                    End Using
                End Using
                LoadChildData()
            End If
        End If
    End Sub
End Class
