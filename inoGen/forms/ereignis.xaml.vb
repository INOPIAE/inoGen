Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics.Metrics

Public Class ereignis
    Private connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)


    Private dtE As New DataTable()
    Private dtO As New DataTable()
    Private dtK As New DataTable()

    Private isNewRecord As Boolean = False
    Private ID As Integer? = Nothing
    Private PID As Integer = 1
    Private FID As Integer


    Public Property PersonId As Integer
        Get
            Return PID
        End Get
        Set(value As Integer)
            PID = value
            NewDataset()
        End Set
    End Property

    Public Property FamilieId As Integer
        Get
            Return FID
        End Get
        Set(value As Integer)
            FID = value
            NewDataset()
        End Set
    End Property

    Public Property EintragId As Integer
        Get
            Return ID
        End Get
        Set(value As Integer)
            ID = value
            LoadEvent(ID)
        End Set
    End Property

    Public Sub New()
        InitializeComponent()

        LoadOrtData()
        LoadKreisListe()
        LoadKonfessionListe()
    End Sub

    Public Event DataSaved(sender As Object, e As EventArgs)

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        Dim sqlFind As String = "SELECT tblEreignisID FROM  tblEreignis WHERE tblEreignisArtID = ? AND  tblPersonID = ?"
        Dim sqlFindF As String = "SELECT tblEreignisID FROM  tblEreignis  WHERE tblEreignisArtID = ? AND  tblFamilieID = ?"
        Dim sqlInsert As String = "INSERT INTO tblEreignis (tblEreignisArtID, tblPersonID, tblFamilieID, Datum, DatumText, BisDatum, BisDatumText, tblOrtID, tblKonfessionID, Referenz, FSID, Info) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Dim sqlUpdate As String = "UPDATE tblEreignis SET tblEreignisArtID = ?, tblPersonID = ?, tblFamilieID = ?, Datum = ?, DatumText = ?, BisDatum = ?, BisDatumText = ?, tblOrtID = ?, tblKonfessionID = ?, Referenz = ?, FSID = ?, Info = ? WHERE tblEreignisID = ?"

        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(sqlFind, conn)
                cmd.Parameters.AddWithValue("@tblEreignisArtID", cbEreignis.SelectedValue)
                cmd.Parameters.AddWithValue("@tblPersonID", PID)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    ID = Convert.ToInt32(result)
                End If
            End Using
            If ID = 0 Or ID Is Nothing Then



                Using cmdInsert As New OleDbCommand(sqlInsert, conn)
                    cmdInsert.Parameters.AddWithValue("@tblEreignisArtID", cbEreignis.SelectedValue)
                    cmdInsert.Parameters.AddWithValue("@tblPersonID", PID)
                    cmdInsert.Parameters.AddWithValue("@tblFamilieID", FID)
                    If IsDate(txtDatum.Text) Then
                        cmdInsert.Parameters.AddWithValue("@Datum", CDate(txtDatum.Text))
                    Else
                        cmdInsert.Parameters.AddWithValue("@Datum", DBNull.Value)
                    End If
                    cmdInsert.Parameters.AddWithValue("@DatumText", txtDatum.Text)
                    If IsDate(txtBisDatum.Text) Then
                        cmdInsert.Parameters.AddWithValue("@BDatum", CDate(txtDatum.Text))
                    Else
                        cmdInsert.Parameters.AddWithValue("@BDatum", DBNull.Value)
                    End If
                    cmdInsert.Parameters.AddWithValue("@BDatumText", txtBisDatum.Text)
                    cmdInsert.Parameters.AddWithValue("@tblOrtID", cbOrt.SelectedValue)
                cmdInsert.Parameters.AddWithValue("@tblKonfessionID", cbKonfession.SelectedValue)
                cmdInsert.Parameters.AddWithValue("@Referenz", txtReferenz.Text)
                cmdInsert.Parameters.AddWithValue("@FSID", txtFSID.Text)
                cmdInsert.Parameters.AddWithValue("@Info", txtInfo.Text)
                cmdInsert.ExecuteNonQuery()
                End Using

                Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                    ID = Convert.ToInt32(cmdId.ExecuteScalar())
                End Using
            Else
                Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                    cmdUpdate.Parameters.AddWithValue("@tblEreignisArtID", cbEreignis.SelectedValue)
                    cmdUpdate.Parameters.AddWithValue("@tblPersonID", PID)
                    cmdUpdate.Parameters.AddWithValue("@tblFamilieID", FID)
                    If IsDate(txtDatum.Text) Then
                        cmdUpdate.Parameters.AddWithValue("@Datum", CDate(txtDatum.Text))
                    Else
                        cmdUpdate.Parameters.AddWithValue("@Datum", DBNull.Value)
                    End If
                    cmdUpdate.Parameters.AddWithValue("@DatumText", txtDatum.Text)
                    If IsDate(txtBisDatum.Text) Then
                        cmdUpdate.Parameters.AddWithValue("@BDatum", CDate(txtDatum.Text))
                    Else
                        cmdUpdate.Parameters.AddWithValue("@BDatum", DBNull.Value)
                    End If
                    cmdUpdate.Parameters.AddWithValue("@BDatumText", txtBisDatum.Text)
                    cmdUpdate.Parameters.AddWithValue("@tblOrtID", cbOrt.SelectedValue)
                    cmdUpdate.Parameters.AddWithValue("@tblKonfessionID", cbKonfession.SelectedValue)
                    cmdUpdate.Parameters.AddWithValue("@Referenz", txtReferenz.Text)
                    cmdUpdate.Parameters.AddWithValue("@FSID", txtFSID.Text)
                    cmdUpdate.Parameters.AddWithValue("@Info", txtInfo.Text)
                    cmdUpdate.Parameters.AddWithValue("@ID", ID)
                    cmdUpdate.ExecuteNonQuery()
                End Using
                If PID > 0 Then
                    My.Settings.LastPID = PID
                    My.Settings.Save()
                End If
                If FID > 0 Then
                    My.Settings.LastFID = FID
                    My.Settings.Save()
                End If
            End If

        End Using
        RaiseEvent DataSaved(Me, EventArgs.Empty)
    End Sub

    Private Sub LoadOrtData()
        Try
            Dim dt As New DataTable()
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT tblOrt.tblOrtID, IIf([tblKreis]![Kreis]<>"""",[tblOrt]![Ort] & "" ("" & [tblKreis]![Kreis] & "")"",[tblOrt]![Ort]) AS Ort
                    FROM tblOrt LEFT JOIN tblKreis ON tblOrt.tblKreisID = tblKreis.tblKreisID ORDER BY Ort", conn)

                Dim adapter As New OleDbDataAdapter(cmd)
                adapter.Fill(dtO)
            End Using

            cbOrt.ItemsSource = dtO.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler beim Laden der Orte: " & ex.Message)
        End Try
    End Sub

    Private Sub btnNew_Click(sender As Object, e As RoutedEventArgs) Handles btnNew.Click
        NewDataset()
    End Sub

    Private Sub NewDataset()
        txtDatum.Clear()
        cbOrt.SelectedValue = 1
        cbEreignis.SelectedValue = 1
        cbKonfession.SelectedValue = 1
        txtReferenz.Clear()
        txtFSID.Clear()
        txtInfo.Clear()
        ID = Nothing
        isNewRecord = True
        cbEreignis.Focus()
    End Sub

    Private Sub cbOrt_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbOrt.SelectionChanged
        If cbOrt.SelectedValue IsNot Nothing Then
            Dim selectedID As Integer = CInt(cbOrt.SelectedValue)
        End If
    End Sub

    Private Sub LoadKreisListe()
        Try
            Dim dt As New DataTable()
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT tblEreignisArtID, EreignisArt FROM tblEreignisArt ORDER BY Reihenfolge", conn)

                Dim adapter As New OleDbDataAdapter(cmd)
                adapter.Fill(dtE)
            End Using

            cbEreignis.ItemsSource = dtE.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler beim Laden der Kreise: " & ex.Message)
        End Try
    End Sub

    Private Sub cbEreignis_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbEreignis.SelectionChanged
        If cbEreignis.SelectedValue IsNot Nothing Then
            Dim selectedID As Integer = CInt(cbEreignis.SelectedValue)
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
                adapter.Fill(dtK)
            End Using

            cbKonfession.ItemsSource = dtK.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler beim Laden der Konfession: " & ex.Message)
        End Try
    End Sub

    Private Sub LoadEvent(id As Int16)
        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Dim sql As String = "SELECT * FROM tblEreignis WHERE tblEreignisID = @id"
                Using cmd As New OleDbCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@id", id)

                    Using reader As OleDbDataReader = cmd.ExecuteReader()
                        If reader.Read() Then
                            ' Beispiel: Felder füllen

                            cbEreignis.SelectedValue = reader("tblEreignisArtID")
                            PID = reader("tblPersonID")
                            FID = reader("tblFamilieID")



                            txtDatum.Text = reader("DatumText").ToString()
                            cbOrt.SelectedValue = reader("tblOrtID")
                            cbKonfession.SelectedValue = reader("tblKonfessionID")
                            txtReferenz.Text = reader("Referenz").ToString()
                            txtFSID.Text = reader("FSID").ToString()
                            txtInfo.Text = reader("Info").ToString()
                        End If
                    End Using
                End Using
            End Using

        Catch ex As Exception
            MessageBox.Show("Fehler beim Laden: " & ex.Message)
        End Try
    End Sub
End Class
