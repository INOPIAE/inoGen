Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics.Metrics
Imports System.Drawing.Text
Imports System.Security.Cryptography
Imports inoGenDLL

Public Class vkHeirat
    Private connectionString As String =
        String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)

    Private dt As New DataTable()

    Private isNewRecord As Boolean = False
    Private ID As Integer? = Nothing
    Private testDate As String
    Private VKHOrt As String

    Private cDB As New clsDB(My.Settings.DBPath)
    Private cGDB As New ClsGenDB(My.Settings.DBPath)
    Private cGH As New ClsGenHelper

    Public Event RequestResizeMainWindow(width As Double)

    Private ReadOnly _main As MainWindow

    Public Sub New(main As MainWindow)
        InitializeComponent()
        _main = main

        If My.Settings.LastVKHID > 0 Then
            ID = My.Settings.LastVKHID
            FillEntry(ID)
        Else
            btnNew_Click(Nothing, Nothing)
            btnSave.ClearValue(Button.BackgroundProperty)
        End If

        LoadData()

        If Not IsNothing(My.Settings.CurrentWork) Then
            Dim i As Integer = 1
            For Each entry In My.Settings.CurrentWork
                If entry.StartsWith("VKHOrt:") Then
                    VKHOrt = entry.Replace("VKHOrt:", "")
                    Exit For
                End If
            Next
        End If

        ckbAutoCorrect.IsChecked = True
        AddHandler Me.Loaded, AddressOf OnLoaded
    End Sub
    Private Sub btnNew_Click(sender As Object, e As RoutedEventArgs) Handles btnNew.Click
        NewEntry()
    End Sub

    Private Sub NewEntry()
        Dim Q As String = txtQuelle.Text
        Dim Seite As String = txtSeite.Text
        Dim Nr() As String = txtNr.Text.Split("/")
        Dim URL As String = txtURL.Text
        Dim QuelleSeite As String = txtQuelleSeite.Text
        ClearAllTextBoxes(Me)
        ID = Nothing
        isNewRecord = True
        txtQuelle.Text = Q
        txtSeite.Text = Seite
        txtURL.Text = URL
        txtQuelleSeite.Text = QuelleSeite
        txtNr.Text = ""

        If Nr.Length = 2 Then
            If IsNumeric(Nr(1)) Then
                Dim v As Integer = CInt(Nr(1)) + 1
                txtNr.Text = Nr(0) & "/" & v.ToString("000")
            ElseIf Nr(1).Length = 1 Then
                Dim nextChar As Char = cGH.NextChar(Nr(1))
                txtNr.Text = Nr(0) & "/" & nextChar
            Else
                txtNr.Text = Nr(0) & "/"
            End If

        End If

        txtKirchort.Text = VKHOrt

        txtHDatum.Focus()

        btnSave.Background = Brushes.LightGreen

    End Sub

    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click


        Dim VID As Int16 = cDB.VornameAnlegen(txtVBtg.Text, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVVtBtg.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVMtBtg.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVBt.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVBt.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVVtBt.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVMtBt.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVZ1.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVZ2.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVZ3.Text.Trim, 0)
        If VID = -1 Then Exit Sub
        VID = cDB.VornameAnlegen(txtVZ4.Text.Trim, 0)
        If VID = -1 Then Exit Sub

        Dim NID As Int16 = cDB.NachnamenID(txtNBtg.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNVtBtg.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNMtBtg.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNBt.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNVtBt.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNMtBt.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNZ1.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNZ2.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNZ3.Text.Trim)
        If NID = -1 Then Exit Sub
        NID = cDB.NachnamenID(txtNZ4.Text.Trim)
        If NID = -1 Then Exit Sub

        Dim OID As Int16 = cDB.OrtID(txtWOBtg.Text.Trim)
        OID = cDB.OrtID(txtKirchort.Text.Trim)
        If OID = -1 Then Exit Sub
        If OID = -1 Then Exit Sub
        OID = cDB.OrtID(txtHOBtg.Text.Trim)
        If OID = -1 Then Exit Sub
        OID = cDB.OrtID(txtWEBtg.Text.Trim)
        If OID = -1 Then Exit Sub
        OID = cDB.OrtID(txtWOBt.Text.Trim)
        If OID = -1 Then Exit Sub
        OID = cDB.OrtID(txtHOBt.Text.Trim)
        If OID = -1 Then Exit Sub
        OID = cDB.OrtID(txtWEBt.Text.Trim)
        If OID = -1 Then Exit Sub

        If txtVZ1.Text.Trim <> "" Or txtNZ1.Text.Trim <> "" Then
            If txtSexZ1.Text.Trim <> "m" And txtSexZ1.Text.Trim <> "w" Then
                MessageBox.Show("Geschlecht 1. Zeuge fehlt oder ist ungültig (m/w)!")
                txtSexZ1.Focus()
                Exit Sub
            End If
        End If

        If txtVZ2.Text.Trim <> "" Or txtNZ2.Text.Trim <> "" Then
            If txtSexZ2.Text.Trim <> "m" And txtSexZ2.Text.Trim <> "w" Then
                MessageBox.Show("Geschlecht 2. Zeuge fehlt oder ist ungültig (m/w)!")
                txtSexZ2.Focus()
                Exit Sub
            End If
        End If

        If txtVZ3.Text.Trim <> "" Or txtNZ3.Text.Trim <> "" Then
            If txtSexZ3.Text.Trim <> "m" And txtSexZ3.Text.Trim <> "w" Then
                MessageBox.Show("Geschlecht 3. Zeuge fehlt oder ist ungültig (m/w)!")
                txtSexZ3.Focus()
                Exit Sub
            End If
        End If

        If txtVZ4.Text.Trim <> "" Or txtNZ4.Text.Trim <> "" Then
            If txtSexZ4.Text.Trim <> "m" And txtSexZ4.Text.Trim <> "w" Then
                MessageBox.Show("Geschlecht 4. Zeuge fehlt oder ist ungültig (m/w)!")
                txtSexZ4.Focus()
                Exit Sub
            End If
        End If


        Dim strInsert As String = "INSERT INTO tblVKH (BUCH_H, SEITE_H, NR_H, HDatum, DimDatum, K_Ort,
            VN_BR, FN_BR, GebDatum_BR, W_BR, H_BR, Z_BR, VN_VBR, FN_VBR, Z_VBR, VN_MBR, FN_MBR, Z_MBR, W_EBR, 
            VN_BT, FN_BT, GebDatum_BT, W_BT, H_BT, Z_BT, VN_VBT, FN_VBT, Z_VBT, VN_MBT, FN_MBT, Z_MBT, W_EBT, 
            ANM_H, VN_HZ1, FN_HZ1, G_HZ1, Z_HZ1, VN_HZ2, FN_HZ2, G_HZ2, Z_HZ2, 
            VN_HZ3, FN_HZ3, G_HZ3, Z_HZ3, VN_HZ4, FN_HZ4, G_HZ4, Z_HZ4, CheckNeeded, OnlineReference, ReferenceDetails) 
            VALUES (?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 
            ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        Dim strUpdate As String = "UPDATE tblVKH SET BUCH_H = ?, SEITE_H = ?, NR_H = ?, HDatum = ?, DimDatum = ?, K_Ort = ?, VN_BR = ?, FN_BR = ?, GebDatum_BR = ?, W_BR = ?, H_BR = ?, Z_BR = ?, VN_VBR = ?, FN_VBR = ?, Z_VBR = ?, VN_MBR = ?, FN_MBR = ?, Z_MBR = ?, W_EBR = ?, VN_BT = ?, FN_BT = ?, GebDatum_BT = ?, W_BT = ?, H_BT = ?, Z_BT = ?, VN_VBT = ?, FN_VBT = ?, Z_VBT = ?, VN_MBT = ?, FN_MBT = ?, Z_MBT = ?, W_EBT = ?, ANM_H = ?, VN_HZ1 = ?, FN_HZ1 = ?, G_HZ1 = ?, Z_HZ1 = ?, VN_HZ2 = ?, FN_HZ2 = ?, G_HZ2 = ?, Z_HZ2 = ?, VN_HZ3 = ?, FN_HZ3 = ?, G_HZ3 = ?, Z_HZ3 = ?, VN_HZ4 = ?, FN_HZ4 = ?, G_HZ4 = ?, Z_HZ4 = ?, CheckNeeded = ?, OnlineReference = ?, ReferenceDetails =? WHERE tblVKHID = ?"

        Try
            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                Dim cmd As New OleDbCommand(If(isNewRecord, strInsert, strUpdate), conn)
                ' Add parameters in the same order as in the SQL statement
                cmd.Parameters.AddWithValue("BUCH_H", txtQuelle.Text.Trim)
                cmd.Parameters.AddWithValue("SEITE_H", txtSeite.Text.Trim)
                cmd.Parameters.AddWithValue("NR_H", txtNr.Text.Trim)
                If IsDate(txtHDatum.Text) Then
                    cmd.Parameters.AddWithValue("@HDatum", CDate(txtHDatum.Text))
                Else
                    cmd.Parameters.AddWithValue("@HDatum", DBNull.Value)
                End If
                If IsDate(txtDimDatum.Text) Then
                    cmd.Parameters.AddWithValue("@DimDatum", CDate(txtDimDatum.Text))
                Else
                    cmd.Parameters.AddWithValue("@DimDatum", DBNull.Value)
                End If
                cmd.Parameters.AddWithValue("K_Ort", txtWOBtg.Text.Trim)


                cmd.Parameters.AddWithValue("VN_BR", txtVBtg.Text.Trim)
                cmd.Parameters.AddWithValue("FN_BR", txtNBtg.Text.Trim)
                If IsDate(txtGebBtg.Text) Then
                    cmd.Parameters.AddWithValue("@GebDatum_BR", CDate(txtGebBtg.Text))
                Else
                    cmd.Parameters.AddWithValue("@GebDatum_BR", DBNull.Value)
                End If
                cmd.Parameters.AddWithValue("W_BR", txtWOBtg.Text.Trim)
                cmd.Parameters.AddWithValue("H_BR", txtHOBtg.Text.Trim)
                cmd.Parameters.AddWithValue("Z_BR", txtZuBtg.Text.Trim)
                cmd.Parameters.AddWithValue("VN_VBR", txtVVtBtg.Text.Trim)
                cmd.Parameters.AddWithValue("FN_VBR", txtNVtBtg.Text.Trim)
                cmd.Parameters.AddWithValue("Z_VBR", txtZuVtBtg.Text.Trim)
                cmd.Parameters.AddWithValue("VN_MBR", txtVMtBtg.Text.Trim)
                cmd.Parameters.AddWithValue("FN_MBR", txtNMtBtg.Text.Trim)
                cmd.Parameters.AddWithValue("Z_MBR", txtZuMtBtg.Text.Trim)
                cmd.Parameters.AddWithValue("W_EBR", txtWEBtg.Text.Trim)

                cmd.Parameters.AddWithValue("VN_BT", txtVBt.Text.Trim)
                cmd.Parameters.AddWithValue("FN_BT", txtNBt.Text.Trim)
                If IsDate(txtGebBt.Text) Then
                    cmd.Parameters.AddWithValue("@GebDatum_BT", CDate(txtGebBt.Text))
                Else
                    cmd.Parameters.AddWithValue("@GebDatum_BT", DBNull.Value)
                End If
                cmd.Parameters.AddWithValue("W_BT", txtWOBt.Text.Trim)
                cmd.Parameters.AddWithValue("H_BT", txtHOBt.Text.Trim)
                cmd.Parameters.AddWithValue("Z_BT", txtZuBt.Text.Trim)
                cmd.Parameters.AddWithValue("VN_VBT", txtVVtBt.Text.Trim)
                cmd.Parameters.AddWithValue("FN_VBT", txtNVtBt.Text.Trim)
                cmd.Parameters.AddWithValue("Z_VBT", txtZuVtBt.Text.Trim)
                cmd.Parameters.AddWithValue("VN_MBT", txtVMtBt.Text.Trim)
                cmd.Parameters.AddWithValue("FN_MBT", txtNMtBt.Text.Trim)
                cmd.Parameters.AddWithValue("Z_MBT", txtZuMtBt.Text.Trim)
                cmd.Parameters.AddWithValue("W_EBT", txtWEBt.Text.Trim)

                cmd.Parameters.AddWithValue("ANM_H", txtInfo.Text.Trim)
                cmd.Parameters.AddWithValue("VN_HZ1", txtVZ1.Text.Trim)
                cmd.Parameters.AddWithValue("FN_HZ1", txtNZ1.Text.Trim)
                cmd.Parameters.AddWithValue("G_HZ1", txtSexZ1.Text.Trim)
                cmd.Parameters.AddWithValue("Z_HZ1", txtZuZ1.Text.Trim)
                cmd.Parameters.AddWithValue("VN_HZ2", txtVZ2.Text.Trim)
                cmd.Parameters.AddWithValue("FN_HZ2", txtNZ2.Text.Trim)
                cmd.Parameters.AddWithValue("G_HZ2", txtSexZ2.Text.Trim)
                cmd.Parameters.AddWithValue("Z_HZ2", txtZuZ2.Text.Trim)

                cmd.Parameters.AddWithValue("VN_HZ3", txtVZ3.Text.Trim)
                cmd.Parameters.AddWithValue("FN_HZ3", txtNZ3.Text.Trim)
                cmd.Parameters.AddWithValue("G_HZ3", txtSexZ3.Text.Trim)
                cmd.Parameters.AddWithValue("Z_HZ3", txtZuZ3.Text.Trim)
                cmd.Parameters.AddWithValue("VN_HZ4", txtVZ4.Text.Trim)
                cmd.Parameters.AddWithValue("FN_HZ4", txtNZ4.Text.Trim)
                cmd.Parameters.AddWithValue("G_HZ4", txtSexZ4.Text.Trim)
                cmd.Parameters.AddWithValue("Z_HZ4", txtZuZ4.Text.Trim)

                cmd.Parameters.AddWithValue("CheckNeeded", If(ckbCheck.IsChecked.HasValue AndAlso ckbCheck.IsChecked.Value, True, False))
                cmd.Parameters.AddWithValue("OnlineReference", txtURL.Text.Trim)
                cmd.Parameters.AddWithValue("ReferenceDetails", txtQuelleSeite.Text.Trim)

                If Not isNewRecord AndAlso ID.HasValue Then
                    cmd.Parameters.AddWithValue("tblVKHID", ID.Value)
                End If
                cmd.ExecuteNonQuery()

                If isNewRecord Then
                    Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                        ID = Convert.ToInt32(cmdId.ExecuteScalar())
                    End Using
                End If

                conn.Close()
                MessageBox.Show("Datensatz erfolgreich " & If(isNewRecord, "erstellt", "aktualisiert") & ".")
                isNewRecord = False
                LoadData()
                My.Settings.LastVKHID = ID
                My.Settings.Save()
            End Using

            btnSave.ClearValue(Button.BackgroundProperty)

        Catch ex As Exception
            MessageBox.Show("Fehler beim Speichern: " & ex.Message)
        End Try

    End Sub

    Private Sub LoadData()
        dt = cGDB.GetVKH_Table()

        dgEintrag.ItemsSource = dt.DefaultView

        If ID.HasValue Then
            For Each rowView As DataRowView In dgEintrag.Items
                If CInt(rowView("tblVKHID")) = ID Then
                    dgEintrag.SelectedItem = rowView
                    dgEintrag.ScrollIntoView(rowView)
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub dgEintrag_AutoGeneratedColumns(sender As Object, e As EventArgs) Handles dgEintrag.AutoGeneratedColumns
        For Each col In dgEintrag.Columns
            If col.Header IsNot Nothing Then
                Select Case col.Header.ToString()
                    Case "SEITE_H", "NR_H", "VN_BR", "FN_BR", "VN_BT", "FN_BT"
                    Case "HDatum"
                        Dim textCol = TryCast(col, DataGridTextColumn)
                        If textCol IsNot Nothing Then
                            Dim oldBinding = TryCast(textCol.Binding, Binding)
                            If oldBinding IsNot Nothing Then
                                textCol.Binding = New Binding(oldBinding.Path.Path) With {
                                    .StringFormat = "dd.MM.yyyy"
                                }
                            End If
                        End If
                    Case Else
                        col.Visibility = Visibility.Collapsed
                End Select
            End If
        Next

        dgEintrag.IsReadOnly = True
        dgEintrag.CanUserAddRows = False
        dgEintrag.CanUserDeleteRows = False
    End Sub

    Private Sub dgEintrag_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs)
        Dim rowView As DataRowView = CType(dgEintrag.SelectedItem, DataRowView)
        If rowView IsNot Nothing Then
            ID = Convert.ToInt32(rowView("tblVKHID"))
            FillEntry(ID)
        End If

    End Sub

    Public Sub ClearAllTextBoxes(parent As DependencyObject)
        For i As Integer = 0 To VisualTreeHelper.GetChildrenCount(parent) - 1
            Dim child As DependencyObject = VisualTreeHelper.GetChild(parent, i)

            If TypeOf child Is TextBox Then
                DirectCast(child, TextBox).Clear()
            Else
                ClearAllTextBoxes(child)
            End If
        Next
        ckbCheck.IsChecked = False
    End Sub

    Private Sub OnLoaded(sender As Object, e As RoutedEventArgs)
        RaiseEvent RequestResizeMainWindow(1300)
    End Sub

    Private Sub FillEntry(EID As Integer)

        ClearAllTextBoxes(Me)
        Dim dtE As DataTable = cGDB.GetVKH_TableEntry(EID)

        If dtE.Rows.Count = 0 Then
            MessageBox.Show("Kein Eintrag mit dieser ID gefunden.")
            Exit Sub
        End If
        With dtE.Rows(0)
            If Not IsDBNull(.Item("BUCH_H")) Then
                txtQuelle.Text = .Item("BUCH_H")
            End If

            If Not IsDBNull(.Item("SEITE_H")) Then
                txtSeite.Text = .Item("SEITE_H")
            End If

            If Not IsDBNull(.Item("NR_H")) Then
                txtNr.Text = .Item("NR_H")
            End If

            If Not IsDBNull(.Item("HDatum")) Then
                txtHDatum.Text = .Item("HDatum")
            End If

            If Not IsDBNull(.Item("DimDatum")) Then
                txtDimDatum.Text = .Item("DimDatum")
            End If

            If Not IsDBNull(.Item("K_Ort")) Then
                txtKirchort.Text = .Item("K_Ort")
            End If

            If Not IsDBNull(.Item("VN_BR")) Then
                txtVBtg.Text = .Item("VN_BR")
            End If

            If Not IsDBNull(.Item("FN_BR")) Then
                txtNBtg.Text = .Item("FN_BR")
            End If
            If Not IsDBNull(.Item("GebDatum_BR")) Then
                txtGebBtg.Text = .Item("GebDatum_BR")
            End If
            If Not IsDBNull(.Item("W_BR")) Then
                txtWOBtg.Text = .Item("W_BR")
            End If
            If Not IsDBNull(.Item("H_BR")) Then
                txtHOBtg.Text = .Item("H_BR")
            End If
            If Not IsDBNull(.Item("Z_BR")) Then
                txtZuBtg.Text = .Item("Z_BR")
            End If
            If Not IsDBNull(.Item("VN_VBR")) Then
                txtVVtBtg.Text = .Item("VN_VBR")
            End If
            If Not IsDBNull(.Item("FN_VBR")) Then
                txtNVtBtg.Text = .Item("FN_VBR")
            End If
            If Not IsDBNull(.Item("Z_VBR")) Then
                txtZuVtBtg.Text = .Item("Z_VBR")
            End If
            If Not IsDBNull(.Item("VN_MBR")) Then
                txtVMtBtg.Text = .Item("VN_MBR")
            End If
            If Not IsDBNull(.Item("FN_MBR")) Then
                txtNMtBtg.Text = .Item("FN_MBR")
            End If
            If Not IsDBNull(.Item("Z_MBR")) Then
                txtZuMtBtg.Text = .Item("Z_MBR")
            End If
            If Not IsDBNull(.Item("W_EBR")) Then
                txtWEBtg.Text = .Item("W_EBR")
            End If
            If Not IsDBNull(.Item("VN_BT")) Then
                txtVBt.Text = .Item("VN_BT")
            End If
            If Not IsDBNull(.Item("FN_BT")) Then
                txtNBt.Text = .Item("FN_BT")
            End If
            If Not IsDBNull(.Item("GebDatum_BT")) Then
                txtGebBt.Text = .Item("GebDatum_BT")
            End If
            If Not IsDBNull(.Item("W_BT")) Then
                txtWOBt.Text = .Item("W_BT")
            End If
            If Not IsDBNull(.Item("H_BT")) Then
                txtHOBt.Text = .Item("H_BT")
            End If
            If Not IsDBNull(.Item("Z_BT")) Then
                txtZuBt.Text = .Item("Z_BT")
            End If
            If Not IsDBNull(.Item("VN_VBT")) Then
                txtVVtBt.Text = .Item("VN_VBT")
            End If
            If Not IsDBNull(.Item("FN_VBT")) Then
                txtNVtBt.Text = .Item("FN_VBT")
            End If
            If Not IsDBNull(.Item("Z_VBT")) Then
                txtZuVtBt.Text = .Item("Z_VBT")
            End If
            If Not IsDBNull(.Item("VN_MBT")) Then
                txtVMtBt.Text = .Item("VN_MBT")
            End If
            If Not IsDBNull(.Item("FN_MBT")) Then
                txtNMtBt.Text = .Item("FN_MBT")
            End If
            If Not IsDBNull(.Item("Z_MBT")) Then
                txtZuMtBt.Text = .Item("Z_MBT")
            End If
            If Not IsDBNull(.Item("W_EBT")) Then
                txtWEBt.Text = .Item("W_EBT")
            End If
            If Not IsDBNull(.Item("ANM_H")) Then
                txtInfo.Text = .Item("ANM_H")
            End If
            If Not IsDBNull(.Item("VN_HZ1")) Then
                txtVZ1.Text = .Item("VN_HZ1")
            End If
            If Not IsDBNull(.Item("FN_HZ1")) Then
                txtNZ1.Text = .Item("FN_HZ1")
            End If
            If Not IsDBNull(.Item("G_HZ1")) Then
                txtSexZ1.Text = .Item("G_HZ1")
            End If
            If Not IsDBNull(.Item("Z_HZ1")) Then
                txtZuZ1.Text = .Item("Z_HZ1")
            End If
            If Not IsDBNull(.Item("VN_HZ2")) Then
                txtVZ2.Text = .Item("VN_HZ2")
            End If
            If Not IsDBNull(.Item("FN_HZ2")) Then
                txtNZ2.Text = .Item("FN_HZ2")
            End If
            If Not IsDBNull(.Item("G_HZ2")) Then
                txtSexZ2.Text = .Item("G_HZ2")
            End If
            If Not IsDBNull(.Item("Z_HZ2")) Then
                txtZuZ2.Text = .Item("Z_HZ2")
            End If
            If Not IsDBNull(.Item("VN_HZ3")) Then
                txtVZ3.Text = .Item("VN_HZ3")
            End If
            If Not IsDBNull(.Item("FN_HZ3")) Then
                txtNZ3.Text = .Item("FN_HZ3")
            End If
            If Not IsDBNull(.Item("G_HZ3")) Then
                txtSexZ3.Text = .Item("G_HZ3")
            End If
            If Not IsDBNull(.Item("Z_HZ3")) Then
                txtZuZ3.Text = .Item("Z_HZ3")
            End If
            If Not IsDBNull(.Item("VN_HZ4")) Then
                txtVZ4.Text = .Item("VN_HZ4")
            End If
            If Not IsDBNull(.Item("FN_HZ4")) Then
                txtNZ4.Text = .Item("FN_HZ4")
            End If
            If Not IsDBNull(.Item("G_HZ4")) Then
                txtSexZ4.Text = .Item("G_HZ4")
            End If
            If Not IsDBNull(.Item("Z_HZ4")) Then
                txtZuZ4.Text = .Item("Z_HZ4")
            End If

            ckbCheck.IsChecked = Not IsDBNull(.Item("CheckNeeded")) AndAlso Convert.ToBoolean(.Item("CheckNeeded"))

            If Not IsDBNull(.Item("OnlineReference")) Then
                txtURL.Text = .Item("OnlineReference")
            End If
            If Not IsDBNull(.Item("ReferenceDetails")) Then
                txtQuelleSeite.Text = .Item("ReferenceDetails")
            End If


            ID = EID
            isNewRecord = False
        End With

        CalculateAge()
    End Sub

    Private Sub txtURL_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles txtURL.MouseDoubleClick
        If txtURL.Text <> "" Then
            Dim url As String = txtURL.Text

            If MainWindow.fsWindow Is Nothing OrElse Not MainWindow.fsWindow.IsLoaded Then
                MainWindow.fsWindow = New FamilySearchWeb(url)
                MainWindow.fsWindow.Show()
            Else
                MainWindow.fsWindow.Focus()
                MainWindow.fsWindow.NavigateTo(url)
            End If
        End If
    End Sub

    Private Sub txtNVtBt_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtNVtBt.GotFocus
        If txtNVtBt.Text.Trim() = "" And txtVVtBt.Text.Trim() <> "" Then
            txtNVtBt.Text = txtNBt.Text.Trim()
        End If
    End Sub

    Private Sub txtNVtBtg_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtNVtBtg.GotFocus
        If txtNVtBtg.Text.Trim() = "" And txtVVtBtg.Text.Trim() <> "" Then
            txtNVtBtg.Text = txtNBtg.Text.Trim()
        End If
    End Sub

    Private Sub txtWEBt_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtWEBt.GotFocus
        If txtWEBt.Text.Trim() = "" And txtWOBt.Text.Trim() <> "" And txtNVtBt.Text.Trim() <> "" Then
            txtWEBt.Text = txtWOBt.Text.Trim()
        End If
    End Sub

    Private Sub txtWEBtg_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtWEBtg.GotFocus
        If txtWEBtg.Text.Trim() = "" And txtWOBtg.Text.Trim() <> "" And txtNVtBtg.Text.Trim() <> "" Then
            txtWEBtg.Text = txtWOBtg.Text.Trim()
        End If
    End Sub

    Private Sub txtHDatum_GotFocus(sender As Object, e As RoutedEventArgs) Handles txtHDatum.GotFocus, txtDimDatum.GotFocus, txtGebBtg.GotFocus, txtGebBt.GotFocus
        testDate = sender.Text
    End Sub

    Private Sub txtHDatum_PreviewLostKeyboardFocus(sender As Object, e As KeyboardFocusChangedEventArgs) _
    Handles txtHDatum.PreviewLostKeyboardFocus, txtGebBt.PreviewLostKeyboardFocus,
            txtDimDatum.PreviewLostKeyboardFocus, txtGebBtg.PreviewLostKeyboardFocus

        Dim tb = DirectCast(sender, TextBox)

        If Not cGDB.CleanDate(tb.Text) Then
            MessageBox.Show("Ungültiges Datum!")
            sender.Text = tb.Text
            e.Handled = True   ' Fokus bleibt im Feld
        Else
            sender.Text = tb.Text
            CalculateAge()
        End If
    End Sub

    Private Sub CalculateAge()
        Dim eventDate As Date
        If IsDate(txtHDatum.Text) Then
            eventDate = CDate(txtHDatum.Text)
        End If
        If IsDate(txtGebBtg.Text) AndAlso IsDate(txtHDatum.Text) Then
            Dim birthDate As Date = CDate(txtGebBtg.Text)
            Dim age As Integer = eventDate.Year - birthDate.Year
            If (eventDate.Month < birthDate.Month) Or (eventDate.Month = birthDate.Month And eventDate.Day < birthDate.Day) Then
                age -= 1
            End If
            lblAlterBtg.Text = age.ToString()
        Else
            lblAlterBtg.Text = ""
        End If
        If IsDate(txtGebBt.Text) AndAlso IsDate(txtHDatum.Text) Then
            Dim birthDate As Date = CDate(txtGebBt.Text)
            Dim age As Integer = eventDate.Year - birthDate.Year
            If (eventDate.Month < birthDate.Month) Or (eventDate.Month = birthDate.Month And eventDate.Day < birthDate.Day) Then
                age -= 1
            End If
            lblAlterBt.Text = age.ToString()
        Else
            lblAlterBt.Text = ""
        End If
    End Sub

    Private Sub txtWOBt_LostFocus(sender As Object, e As RoutedEventArgs) Handles _
            txtWOBt.LostFocus, txtWOBtg.LostFocus, txtHOBt.LostFocus, txtHOBtg.LostFocus,
            txtWEBt.LostFocus, txtWEBtg.LostFocus,
            txtNBt.LostFocus, txtNBtg.LostFocus, txtVBt.LostFocus, txtVBtg.LostFocus,
            txtNVtBt.LostFocus, txtNVtBtg.LostFocus, txtVVtBt.LostFocus, txtVVtBtg.LostFocus,
            txtNMtBt.LostFocus, txtNMtBtg.LostFocus, txtVMtBt.LostFocus, txtVMtBtg.LostFocus,
            txtNZ1.LostFocus, txtNZ2.LostFocus, txtNZ3.LostFocus, txtNZ4.LostFocus,
            txtVZ1.LostFocus, txtVZ2.LostFocus, txtVZ3.LostFocus, txtVZ4.LostFocus,
            txtKirchort.LostFocus

        If ckbAutoCorrect.IsChecked = True Then
            sender.text = cGDB.ToTitleCase(sender.text)
            _main.CAutoCorrect.CheckAutoCorrection(sender)
        End If

    End Sub

    Private Sub btnNewP_Click(sender As Object, e As RoutedEventArgs) Handles btnNewP.Click
        NewEntry()
        Dim Nr() As String = txtQuelleSeite.Text.Split({" "c}, StringSplitOptions.RemoveEmptyEntries)

        If IsNumeric(txtSeite.Text) Then
            txtSeite.Text = CStr(CInt(txtSeite.Text) + 1)
        End If

        If IsNumeric(Nr(Nr.Length - 1)) Then
            Nr(Nr.Length - 1) = CInt(Nr(Nr.Length - 1)) + 1
            txtQuelleSeite.Text = String.Join(" ", Nr)
        End If
    End Sub

    Private Sub txtInfo_LostFocus(sender As Object, e As RoutedEventArgs) Handles txtInfo.LostFocus, txtZuBtg.LostFocus, txtZuBt.LostFocus,
            txtZuVtBtg.LostFocus, txtZuVtBt.LostFocus,
            txtZuMtBtg.LostFocus, txtZuMtBt.LostFocus,
            txtZuZ1.LostFocus, txtZuZ2.LostFocus, txtZuZ3.LostFocus, txtZuZ4.LostFocus

        If ckbAutoCorrect.IsChecked = True Then
            _main.CAutoCorrect.CheckAutoCorrection(sender)
        End If
    End Sub
End Class
