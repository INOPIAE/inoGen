Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics.Metrics
Imports inoGenDLL

Public Class ereignis
    Implements IFormularClipboard

    Private connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)


    Private dtE As New DataTable()
    Private dtO As New DataTable()
    Private dtK As New DataTable()
    Private dtEventtag As New DataTable()

    Private isNewRecord As Boolean = False
    Private ID As Integer? = Nothing
    Private PID As Integer = 1
    Private FID As Integer
    Private VID As String
    Private MID As String
    Private isPers As Boolean = True
    Private EAID As Integer
    Private QuellZitatEreignisID As Integer
    Private QuellZitatEreignisVID As Integer
    Private QuellZitatEreignisMID As Integer

    Private cGenDB As New inoGenDLL.ClsGenDB(My.Settings.DBPath)
    Private cGH As New ClsGenHelper

    Private ReadOnly _main As MainWindow

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

    Public Property isPerson As Boolean
        Get
            Return isPers
        End Get
        Set(value As Boolean)
            isPers = value
        End Set
    End Property

    Public Sub New(isPerson As Boolean, main As MainWindow)
        InitializeComponent()
        _main = main
        isPers = isPerson
        LoadOrtData()
        LoadEventListe()
        LoadKonfessionListe()
        LoadEventTag()
        AddHandler Me.Loaded, AddressOf Ereignis_Loaded
        AddHandler Me.GotFocus, AddressOf Ereignis_GotFocus
    End Sub

    Public Event DataSaved(sender As Object, e As EventArgs)

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        SaveData()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As RoutedEventArgs) Handles btnDelete.Click
        If ID = 0 Or ID Is Nothing Then
            MessageBox.Show("Kein Datensatz zum Löschen ausgewählt.")
            Exit Sub
        End If
        If MessageBox.Show("Soll der Datensatz wirklich gelöscht werden?", "Datensatz löschen", MessageBoxButton.YesNo) = MessageBoxResult.No Then
            Exit Sub
        End If
        Select Case cGenDB.DeleteEreignis(ID)
            Case 0
                MessageBox.Show("Datensatz wurde nicht gelöscht.")
            Case 1
                MessageBox.Show("Datensatz konnte nicht gelöscht werden.")
            Case 2
                MessageBox.Show("Datensatz wurde gelöscht.")
                NewDataset()
        End Select
    End Sub
    Public Sub SaveData()
        If cbEreignis.SelectedValue > 8 And txtZusatz.Text.Trim <> "" Then
            Dim ZID As Int16 = ZusatzID(txtZusatz.Text, CInt(cbEreignis.SelectedValue))
            If ZID = -1 Then Exit Sub
        End If

        If IsNothing(cbEreignis.SelectedValue) Then
            MessageBox.Show("Bitte Ereignis auswählen.")
            cbEreignis.Focus()
            Exit Sub
        End If

        txtDatum.Text = cGH.CleanupDateString(txtDatum.Text)
        txtBisDatum.Text = cGH.CleanupDateString(txtBisDatum.Text)

        If cGH.IsValidDateString(txtDatum.Text) = False Then
            MessageBox.Show("Bitte ein gültiges Datum eingeben.")
            txtDatum.Focus()
            Exit Sub
        End If
        If cGH.IsValidDateString(txtBisDatum.Text) = False Then
            MessageBox.Show("Bitte ein gültiges Bis-Datum eingeben.")
            txtBisDatum.Focus()
            Exit Sub
        End If

        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            If ID = 0 Or ID Is Nothing Then
                ID = cGenDB.SetEreignis(If(cbEreignis.SelectedValue IsNot Nothing, CInt(cbEreignis.SelectedValue), 0), txtDatum.Text, cGenDB.CalculateDatum(txtDatum.Text), txtBisDatum.Text, cGenDB.CalculateDatum(txtBisDatum.Text), If(cbOrt.SelectedValue IsNot Nothing, CInt(cbOrt.SelectedValue), 0), If(cbKonfession.SelectedValue IsNot Nothing, CInt(cbKonfession.SelectedValue), 0), txtZusatz.Text, txtReferenz.Text, txtFSID.Text, txtInfo.Text, PID, FID)
            Else
                cGenDB.UpdateEreignis(ID, If(cbEreignis.SelectedValue IsNot Nothing, CInt(cbEreignis.SelectedValue), 0), txtDatum.Text, cGenDB.CalculateDatum(txtDatum.Text), txtBisDatum.Text, cGenDB.CalculateDatum(txtBisDatum.Text), If(cbOrt.SelectedValue IsNot Nothing, CInt(cbOrt.SelectedValue), 0), If(cbKonfession.SelectedValue IsNot Nothing, CInt(cbKonfession.SelectedValue), 0), txtZusatz.Text, txtReferenz.Text, txtFSID.Text, txtInfo.Text)
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
        txtBisDatum.Clear()
        cbOrt.SelectedValue = 1
        cbEreignis.SelectedValue = 1
        cbKonfession.SelectedValue = 1
        txtZusatz.Clear()
        txtReferenz.Clear()
        txtFSID.Clear()
        txtInfo.Clear()
        txtGeb.Clear()
        txtJahr.Clear()
        txtMonate.Clear()
        txtWochen.Clear()
        txtTage.Clear()
        ID = Nothing
        isNewRecord = True
        cbEreignis.Focus()
        QuellZitatEreignisID = 0
        QuellZitatEreignisMID = 0
        QuellZitatEreignisVID = 0
    End Sub

    Private Sub cbOrt_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbOrt.SelectionChanged
        If cbOrt.SelectedValue IsNot Nothing Then
            Dim selectedID As Integer = CInt(cbOrt.SelectedValue)
        End If
    End Sub

    Private Sub LoadEventListe()
        Try
            Dim dt As New DataTable()
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand("SELECT tblEreignisArtID, EreignisArt FROM tblEreignisArt WHERE PersonenEreignis = ? ORDER BY Reihenfolge", conn)
                cmd.Parameters.Add("PersonenEreignis", OleDbType.Boolean).Value = isPers
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

        Dim row As DataRowView = TryCast(cbEreignis.SelectedItem, DataRowView)
        If row IsNot Nothing Then
            Dim value = row("tblEreignisArtID")
            Dim text = row("EreignisArt").ToString()
            EAID = value
            If value > 8 Then
                lblZusatz.Text = text
            Else
                lblZusatz.Text = "Zusatz"
            End If
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
                            txtZusatz.Text = reader("Zusatz").ToString()
                            txtReferenz.Text = reader("Referenz").ToString()
                            txtFSID.Text = reader("FSID").ToString()
                            txtInfo.Text = reader("Info").ToString()
                            txtGeb.Clear()
                            txtJahr.Clear()
                            txtMonate.Clear()
                            txtWochen.Clear()
                            txtTage.Clear()
                        End If
                    End Using
                End Using
            End Using

            If isPers Then
                Dim dtEventPerson As DataTable = cGenDB.GetQuellzitatEvent(id, PID)
                If dtEventPerson.Rows.Count > 0 Then
                    Dim row As DataRow = dtEventPerson.Rows(0)
                    QuellZitatEreignisID = row("tblEreignisZitatID")
                    txtQuellZitat.Text = row("tblQuellZitatID").ToString()
                    If row("EventTag") IsNot DBNull.Value Then
                        cbEventTag.SelectedValue = row("EventTag")
                    Else
                        cbEventTag.SelectedValue = "-"
                    End If
                Else
                    QuellZitatEreignisID = 0
                    cbEventTag.SelectedValue = "-"
                    txtQuellZitat.Clear()
                End If
            Else
                VID = cGenDB.GetParentIDFromFamily(FID, True)
                MID = cGenDB.GetParentIDFromFamily(FID, False)
                If VID <> "" Then
                    Dim dtEventVater As DataTable = cGenDB.GetQuellzitatEvent(id, CInt(VID))
                    If dtEventVater.Rows.Count > 0 Then
                        Dim row As DataRow = dtEventVater.Rows(0)
                        QuellZitatEreignisVID = row("tblEreignisZitatID")
                        txtQuellZitatM.Text = row("tblQuellZitatID").ToString()
                        If row("EventTag") IsNot DBNull.Value Then
                            cbEventTagM.SelectedValue = row("EventTag")
                        Else
                            cbEventTagM.SelectedValue = "-"
                        End If
                    Else
                        QuellZitatEreignisVID = 0
                        cbEventTagM.SelectedValue = "-"
                        txtQuellZitatM.Clear()
                    End If
                End If
                If MID <> "" Then
                    Dim dtEventMutter As DataTable = cGenDB.GetQuellzitatEvent(id, CInt(MID))
                    If dtEventMutter.Rows.Count > 0 Then
                        Dim row As DataRow = dtEventMutter.Rows(0)
                        QuellZitatEreignisMID = row("tblEreignisZitatID")
                        txtQuellZitatV.Text = row("tblQuellZitatID").ToString()
                        If row("EventTag") IsNot DBNull.Value Then
                            cbEventTagV.SelectedValue = row("EventTag")
                        Else
                            cbEventTagV.SelectedValue = "-"
                        End If
                    Else
                        QuellZitatEreignisMID = 0
                        cbEventTagV.SelectedValue = "-"
                        txtQuellZitatV.Clear()
                    End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Fehler beim Laden: " & ex.Message)
        End Try
    End Sub

    Private Function ZusatzID(Zusatz As String, EreignisID As Integer) As Int16
        Dim id As Integer = -1
        Dim sqlSelect As String = "SELECT tblZusatzID FROM tblZusatz WHERE Zusatz = ? AND tblEreignisArtID = ?"
        Dim sqlInsert As String = "INSERT INTO tblZusatz (Zusatz, tblEreignisArtID) VALUES (?, ?)"
        If Trim(Zusatz) = "" Then
            Return 0
        End If
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(sqlSelect, conn)
                cmd.Parameters.AddWithValue("@Zusatz", Zusatz)
                cmd.Parameters.AddWithValue("@tblEreignisArtID", EreignisID)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    id = Convert.ToInt32(result)
                    Return id
                End If
            End Using

            If MessageBox.Show(String.Format("Soll der Eintrag '{0}' angelegt werden?", Zusatz), String.Format("{0} anlegen", lblZusatz.Text), MessageBoxButton.YesNo) = MessageBoxResult.No Then
                Return -1
            End If
            Using cmdInsert As New OleDbCommand(sqlInsert, conn)
                cmdInsert.Parameters.AddWithValue("@Zusatz", Zusatz)
                cmdInsert.Parameters.AddWithValue("@tblEreignisArtID", EreignisID)
                cmdInsert.ExecuteNonQuery()
            End Using

            Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                id = Convert.ToInt32(cmdId.ExecuteScalar())
            End Using
        End Using
        Return id
    End Function

    Public Sub InitNew(personId As Integer,
                   familieId As Integer,
                   Optional autoSave As Boolean = False)

        NewDataset()

        Me.PersonId = personId
        Me.FamilieId = familieId


        If autoSave Then
            SaveData()
        End If
    End Sub

    Private Sub txtZusatz_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtZusatz.LostFocus, txtInfo.LostFocus
        sender.text = cGenDB.ToTitleCase(sender.text)
        _main.CAutoCorrect.CheckAutoCorrection(sender)
    End Sub

    Private Sub btnCalcAge_Click(sender As Object, e As RoutedEventArgs) Handles btnCalcAge.Click
        Dim cA As New inoGenDLL.ClsAlter
        txtGeb.Text = cA.CalculateBirthday(txtDatum.Text, txtJahr.Text, txtMonate.Text, txtWochen.Text, txtTage.Text)
        If txtGeb.Text.StartsWith("err") Then
            Clipboard.SetText(txtGeb.Text)
        End If
    End Sub

    Private Sub cbEventTag_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbEventTag.SelectionChanged
        'If cbOrt.SelectedValue IsNot Nothing Then
        '    Dim selectedID As Integer = CInt(cbOrt.SelectedValue)
        'End If
    End Sub

    Private Sub cbEventTagV_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbEventTagV.SelectionChanged
        'If cbOrt.SelectedValue IsNot Nothing Then
        '    Dim selectedID As Integer = CInt(cbOrt.SelectedValue)
        'End If
    End Sub

    Private Sub cbEventTagM_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbEventTagM.SelectionChanged
        'If cbOrt.SelectedValue IsNot Nothing Then
        '    Dim selectedID As Integer = CInt(cbOrt.SelectedValue)
        'End If
    End Sub
    Private Sub LoadEventTag()
        dtEventtag.Clear()
        dtEventtag = cGenDB.GetEventTag()

        Dim row As DataRow = dtEventtag.NewRow()
        row("Tag") = "-"
        row("TagD") = "   "
        dtEventtag.Rows.InsertAt(row, 0)

        cbEventTag.ItemsSource = dtEventtag.DefaultView
        cbEventTag.DisplayMemberPath = "TagD"
        cbEventTag.SelectedValuePath = "Tag"
        cbEventTagV.ItemsSource = dtEventtag.DefaultView
        cbEventTagV.DisplayMemberPath = "TagD"
        cbEventTagV.SelectedValuePath = "Tag"
        cbEventTagM.ItemsSource = dtEventtag.DefaultView
        cbEventTagM.DisplayMemberPath = "TagD"
        cbEventTagM.SelectedValuePath = "Tag"

    End Sub

    Public Sub setQuellZitatID(QuellZitatID As Integer)
        If isPers Then
            txtQuellZitat.Text = QuellZitatID.ToString()
            If cbEreignis.SelectedValue < 8 Then
                cbEventTag.SelectedValue = "_PROB"
            End If
        Else
            If VID <> "" Then
                txtQuellZitatV.Text = QuellZitatID.ToString()
                If cbEreignis.SelectedValue < 5 Then
                    cbEventTagV.SelectedValue = "_BRIDEGROOM"
                End If
            End If
            If MID <> "" Then
                txtQuellZitatM.Text = QuellZitatID.ToString()
                If cbEreignis.SelectedValue < 5 Then
                    cbEventTagM.SelectedValue = "_BRIDE"
                End If
            End If
        End If
    End Sub

    Private Sub btnSaveSource_Click(sender As Object, e As RoutedEventArgs) Handles btnSaveSource.Click
        If IsNothing(ID) Then
            MessageBox.Show("Bitte zuerst den Datensatz speichern.")
            Exit Sub
        End If
        If isPers Then
            If IsNumeric(txtQuellZitat.Text) = False Then
                MessageBox.Show("Bitte eine gültige Quell-Zitat-ID eingeben.")
                Exit Sub
            End If
            If cbEventTag.SelectedValue = "-" Then
                MessageBox.Show("Bitte einen Event-Tag auswählen.")
                Exit Sub
            End If
            If QuellZitatEreignisID = 0 Then
                QuellZitatEreignisID = cGenDB.SetEreignisZitat(CInt(txtQuellZitat.Text), ID, PID, IIf(cbEventTag.SelectedValue.ToString() = "-", "", cbEventTag.SelectedValue))
            Else
                cGenDB.UpdateEreignisZitat(QuellZitatEreignisID, CInt(txtQuellZitat.Text), ID, PID, IIf(cbEventTag.SelectedValue.ToString() = "-", "", cbEventTag.SelectedValue))
            End If
            RaiseEvent DataSaved(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub btnSaveSource_ClickV(sender As Object, e As RoutedEventArgs) Handles btnSaveSourceV.Click
        If IsNothing(ID) Then
            MessageBox.Show("Bitte zuerst den Datensatz speichern.")
            Exit Sub
        End If
        If isPers = False And VID <> "" Then
            If IsNumeric(txtQuellZitatV.Text) = False Then
                MessageBox.Show("Bitte eine gültige Quell-Zitat-ID eingeben.")
                Exit Sub
            End If
            If cbEventTagV.SelectedValue = "-" Then
                MessageBox.Show("Bitte einen Event-Tag auswählen.")
                Exit Sub
            End If
            If QuellZitatEreignisVID = 0 Then
                QuellZitatEreignisVID = cGenDB.SetEreignisZitat(CInt(txtQuellZitatV.Text), ID, CInt(VID), IIf(cbEventTagV.SelectedValue.ToString() = "-", "", cbEventTagV.SelectedValue))
            Else
                cGenDB.UpdateEreignisZitat(QuellZitatEreignisVID, CInt(txtQuellZitatV.Text), ID, CInt(VID), IIf(cbEventTagV.SelectedValue.ToString() = "-", "", cbEventTagV.SelectedValue))
            End If
            RaiseEvent DataSaved(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub btnSaveSource_ClickM(sender As Object, e As RoutedEventArgs) Handles btnSaveSourceM.Click
        If IsNothing(ID) Then
            MessageBox.Show("Bitte zuerst den Datensatz speichern.")
            Exit Sub
        End If
        If isPers = False And MID <> "" Then
            If IsNumeric(txtQuellZitatM.Text) = False Then
                MessageBox.Show("Bitte eine gültige Quell-Zitat-ID eingeben.")
                Exit Sub
            End If
            If cbEventTagM.SelectedValue = "-" Then
                MessageBox.Show("Bitte einen Event-Tag auswählen.")
                Exit Sub
            End If
            If QuellZitatEreignisMID = 0 Then
                QuellZitatEreignisMID = cGenDB.SetEreignisZitat(CInt(txtQuellZitatM.Text), ID, CInt(MID), IIf(cbEventTagM.SelectedValue.ToString() = "-", "", cbEventTagM.SelectedValue))
            Else
                cGenDB.UpdateEreignisZitat(QuellZitatEreignisMID, CInt(txtQuellZitatM.Text), ID, CInt(MID), IIf(cbEventTagM.SelectedValue.ToString() = "-", "", cbEventTagM.SelectedValue))
            End If
            RaiseEvent DataSaved(Me, EventArgs.Empty)
        End If
    End Sub

    Public Function CopyFormData() As ClsFormularDatenCopy Implements IFormularClipboard.CopyFormData
        Dim daten As New ClsFormularDatenCopy()
        daten.Datum = txtDatum.Text
        daten.DatumBis = txtBisDatum.Text
        daten.OrtID = CInt(cbOrt.SelectedValue)

        Return daten
    End Function

    Public Sub PasteFormData(daten As ClsFormularDatenCopy) Implements IFormularClipboard.PasteFormData
        If daten Is Nothing Then
            Return
        End If
        txtBisDatum.Text = daten.DatumBis
        txtDatum.Text = daten.Datum
        cbOrt.SelectedValue = daten.OrtID
    End Sub

    Private Sub Ereignis_Loaded(sender As Object, e As RoutedEventArgs)
        ClsFormularClipboard.SetActiveForm(Me)
    End Sub

    Private Sub Ereignis_GotFocus(sender As Object, e As RoutedEventArgs)
        ClsFormularClipboard.SetActiveForm(Me)
    End Sub
End Class
