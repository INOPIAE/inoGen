Imports System.Data
Imports System.Data.OleDb
Imports inoGenDLL

Class QuellenZitate
    Implements IFormularClipboard

    Private connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", My.Settings.DBPath)

    Private dt As New DataTable()
    Private dtQ As New DataTable()
    Private dtE As New DataTable()

    Private isNewRecord As Boolean = False
    Private ID As Integer? = Nothing

    Private isUCLoaded As Boolean = False

    Private cGenDB As New ClsGenDB(My.Settings.DBPath)

    Private ReadOnly _navigation As INavigationService

    Public Property EreignisControl As ereignis

    Private Sub btnFilter_Click(sender As Object, e As RoutedEventArgs)
        LoadData()
        My.Settings.QZQuelle = cbQuelle.SelectedValue
        My.Settings.QZEreignis = cbEreignisArt.SelectedValue
        My.Settings.QZJahr = txtJahr.Text
        My.Settings.QZBd = txtBd.Text
        My.Settings.QZSeite = txtSeite.Text
        My.Settings.QZNummer = txtNummer.Text
        My.Settings.Save()
    End Sub

    Private Sub btnCurrent_Click(sender As Object, e As RoutedEventArgs)
        If dgQuellZitate.SelectedItem Is Nothing Then
            MessageBox.Show("Bitte eine Zeile auswählen.")
            Return
        End If
        If EreignisControl IsNot Nothing Then
            EreignisControl.setQuellZitatID(CType(dgQuellZitate.SelectedItem, DataRowView)("tblQuellZitatID"))
            My.Settings.LastQuellZitat = CType(dgQuellZitate.SelectedItem, DataRowView)("tblQuellZitatID")
            My.Settings.Save()
        End If

    End Sub

    Private Sub btnLastCitation_Click(sender As Object, e As RoutedEventArgs)
        If EreignisControl IsNot Nothing Then
            EreignisControl.setQuellZitatID(My.Settings.LastQuellZitat)
        End If
    End Sub

    Private Sub btnUsage_Click(sender As Object, e As RoutedEventArgs)
        If dgQuellZitate.SelectedItem Is Nothing Then
            MessageBox.Show("Bitte eine Zeile auswählen.")
            Return
        End If

        Dim wnd As New allgemeinesFenster(
            New QuellenZitatEreignis(CType(dgQuellZitate.SelectedItem, DataRowView)("tblQuellZitatID"), _navigation), _navigation, "Verwendung")
        wnd.ShowDialog()

    End Sub

    Private Sub LoadData()
        Dim filter As New Dictionary(Of String, Object)

        If cbQuelle.SelectedValue IsNot Nothing AndAlso CInt(cbQuelle.SelectedValue) <> 0 Then
            filter.Add("tblQuelleID", CInt(cbQuelle.SelectedValue))
        End If

        If cbEreignisArt.SelectedValue IsNot Nothing AndAlso CInt(cbEreignisArt.SelectedValue) <> 0 Then
            filter.Add("tblEreignisArtID", CInt(cbEreignisArt.SelectedValue))
        End If

        If txtJahr.Text.Trim() <> "" Then
            Dim jahr As Integer
            If Integer.TryParse(txtJahr.Text.Trim(), jahr) Then
                filter.Add("Jahr", jahr)
            Else
                MessageBox.Show("Bitte geben Sie eine gültige Jahreszahl ein.")
                Return
            End If
        End If
        If txtSeite.Text.Trim() <> "" Then
            filter.Add("Seite", txtSeite.Text.Trim())
        End If
        If txtBd.Text.Trim() <> "" Then
            filter.Add("Bd", txtBd.Text.Trim())
        End If
        If txtNummer.Text.Trim() <> "" Then
            filter.Add("Nummer", txtNummer.Text.Trim())
        End If



        Try
            dt.Clear()
            dt = cGenDB.GetQuellZitateF(filter)

            dgQuellZitate.ItemsSource = dt.DefaultView

        Catch ex As Exception
            MessageBox.Show("Fehler: " & ex.Message)
        End Try

        If ID.HasValue Then
            For Each rowView As DataRowView In dgQuellZitate.Items
                If CInt(rowView("tblQuellZitateID")) = ID Then
                    ' Selektion setzen
                    dgQuellZitate.SelectedItem = rowView

                    ' Sichtbar machen
                    dgQuellZitate.ScrollIntoView(rowView)

                    Exit For
                End If
            Next
        End If

    End Sub

    Public Sub New(navigation As INavigationService)
        InitializeComponent()

        _navigation = navigation
        LoadData()
        LoadDataQuellen()
        LoadDataEreignis()
        LoadFilter()
        LoadData()

        AddHandler Me.Loaded, AddressOf QuellenZitate_Loaded
        AddHandler Me.GotFocus, AddressOf QuellenZitate_GotFocus
    End Sub

    Private Sub LoadDataQuellen()
        Try
            dtQ.Clear()
            dtQ = cGenDB.GetQuellen

            Dim row As DataRow = dtQ.NewRow()
            row("tblQuelleID") = 0
            row("QuelleKurz") = "Alle"
            dtQ.Rows.InsertAt(row, 0)

            cbQuelle.ItemsSource = dtQ.DefaultView
            cbQuelle.DisplayMemberPath = "QuelleKurz"
            cbQuelle.SelectedValuePath = "tblQuelleID"

            Dim colQuelle As New DataGridComboBoxColumn()

            colQuelle.Header = "Quelle"
            colQuelle.ItemsSource = dtQ.DefaultView
            colQuelle.DisplayMemberPath = "QuelleKurz"
            colQuelle.SelectedValuePath = "tblQuelleID"

            ' Bindung an tblQuelleID der DataGrid-Zeile
            colQuelle.SelectedValueBinding =
                New Binding("tblQuelleID") With {
                    .Mode = BindingMode.TwoWay,
                    .UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                }

            dgQuellZitate.Columns.Add(colQuelle)

        Catch ex As Exception
            MessageBox.Show("Fehler: " & ex.Message)
        End Try


    End Sub

    Private Sub LoadDataEreignis()
        Try
            dtE.Clear()
            Try

                Using conn As New OleDbConnection(connectionString)
                    conn.Open()
                    Dim cmd As New OleDbCommand("SELECT tblEreignisArtID, EreignisArt FROM tblEreignisArt ORDER BY Reihenfolge", conn)
                    Dim adapter As New OleDbDataAdapter(cmd)
                    adapter.Fill(dtE)
                End Using


            Catch ex As Exception
                MessageBox.Show("Fehler beim Laden der Kreise: " & ex.Message)
            End Try

            Dim row As DataRow = dtE.NewRow()
            row("tblEreignisArtID") = 0
            row("EreignisArt") = "Alle"
            dtE.Rows.InsertAt(row, 0)

            cbEreignisArt.ItemsSource = dtE.DefaultView
            cbEreignisArt.DisplayMemberPath = "EreignisArt"
            cbEreignisArt.SelectedValuePath = "tblEreignisArtID"
            Dim colEreignis As New DataGridComboBoxColumn()

            colEreignis.Header = "Ereignis"
            colEreignis.ItemsSource = dtE.DefaultView
            colEreignis.DisplayMemberPath = "EreignisArt"
            colEreignis.SelectedValuePath = "tblEreignisArtID"

            ' Bindung an tblEreignisID der DataGrid-Zeile
            colEreignis.SelectedValueBinding =
                New Binding("tblEreignisArtID") With {
                    .Mode = BindingMode.TwoWay,
                    .UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                }

            dgQuellZitate.Columns.Add(colEreignis)

        Catch ex As Exception
            MessageBox.Show("Fehler: " & ex.Message)
        End Try


    End Sub

    Private Sub dgQuellZitate_AutoGeneratedColumns(
    sender As Object,
    e As EventArgs
) Handles dgQuellZitate.AutoGeneratedColumns

        Dim jahrConverter As New ClsJahrBackgroundConverter()



        For i As Integer = dgQuellZitate.Columns.Count - 1 To 0 Step -1

            Dim col As DataGridColumn = dgQuellZitate.Columns(i)

            If col.Header Is Nothing Then Continue For

            Dim header As String = col.Header.ToString()

            ' =========================================================
            ' ID-SPALTEN AUSBLENDEN
            ' =========================================================
            'If String.Equals(header, "tblQuellZitatID", StringComparison.OrdinalIgnoreCase) _
            If String.Equals(header, "tblQuelleID", StringComparison.OrdinalIgnoreCase) _
            OrElse String.Equals(header, "tblEreignisArtID", StringComparison.OrdinalIgnoreCase) _
            OrElse String.Equals(header, "Anzahl", StringComparison.OrdinalIgnoreCase) _
            OrElse String.Equals(header, "active", StringComparison.OrdinalIgnoreCase) Then

                col.Visibility = Visibility.Collapsed


                ' =========================================================
                ' JAHR
                ' =========================================================
            ElseIf String.Equals(header, "tblQuellZitatID", StringComparison.OrdinalIgnoreCase) And isUCLoaded = False Then

                Dim index As Integer = i

                dgQuellZitate.Columns.RemoveAt(i)

                Dim jahrColumn As New DataGridTemplateColumn()
                jahrColumn.Header = "ID"

                ' -----------------------------------------------------
                ' Border für Hintergrundfarbe
                ' -----------------------------------------------------
                Dim borderFactory As New FrameworkElementFactory(
                GetType(Border))

                borderFactory.SetBinding(
                Border.BackgroundProperty,
                New Binding("Anzahl") With {
                    .Converter = jahrConverter
                })

                ' -----------------------------------------------------
                ' Jahr anzeigen
                ' -----------------------------------------------------
                Dim textFactory As New FrameworkElementFactory(
                GetType(TextBlock))

                textFactory.SetBinding(
                TextBlock.TextProperty,
                New Binding("tblQuellZitatID"))

                textFactory.SetValue(
                TextBlock.VerticalAlignmentProperty,
                VerticalAlignment.Center)

                textFactory.SetValue(
                TextBlock.MarginProperty,
                New Thickness(4, 0, 4, 0))

                borderFactory.AppendChild(textFactory)

                Dim jahrTemplate As New DataTemplate()
                jahrTemplate.VisualTree = borderFactory

                jahrColumn.CellTemplate = jahrTemplate

                dgQuellZitate.Columns.Insert(index, jahrColumn)


                ' =========================================================
                ' DATUM
                ' =========================================================
            ElseIf String.Equals(header, "Datum", StringComparison.OrdinalIgnoreCase) Then
                If isUCLoaded = False Then
                    Dim index As Integer = i

                    dgQuellZitate.Columns.RemoveAt(i)

                    Dim datumColumn As New DataGridTemplateColumn()
                    datumColumn.Header = "Datum "

                    ' -----------------------------------------------------
                    ' Anzeige
                    ' -----------------------------------------------------
                    Dim textFactory As New FrameworkElementFactory(
                GetType(TextBlock))

                    textFactory.SetBinding(
                TextBlock.TextProperty,
                New Binding("Datum") With {
                    .StringFormat = "dd.MM.yyyy"
                })

                    textFactory.SetValue(
                TextBlock.VerticalAlignmentProperty,
                VerticalAlignment.Center)

                    textFactory.SetValue(
                TextBlock.MarginProperty,
                New Thickness(4, 0, 4, 0))

                    Dim displayTemplate As New DataTemplate()
                    displayTemplate.VisualTree = textFactory

                    datumColumn.CellTemplate = displayTemplate

                    ' -----------------------------------------------------
                    ' Bearbeiten → DatePicker
                    ' -----------------------------------------------------
                    Dim datePickerFactory As New FrameworkElementFactory(
                GetType(DatePicker))

                    datePickerFactory.SetBinding(
                DatePicker.SelectedDateProperty,
                New Binding("Datum") With {
                    .Mode = BindingMode.TwoWay,
                    .UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged
                })

                    datePickerFactory.SetValue(
                DatePicker.VerticalAlignmentProperty,
                VerticalAlignment.Center)

                    datePickerFactory.SetValue(
                DatePicker.MarginProperty,
                New Thickness(0))

                    Dim editTemplate As New DataTemplate()
                    editTemplate.VisualTree = datePickerFactory

                    datumColumn.CellEditingTemplate = editTemplate

                    dgQuellZitate.Columns.Insert(index, datumColumn)
                Else
                    col.Visibility = Visibility.Collapsed
                End If


                ' =========================================================
                ' INTERNETADRESSE
                ' =========================================================
            ElseIf String.Equals(header, "InternetAdresse", StringComparison.OrdinalIgnoreCase) And isUCLoaded = False Then

                Dim index As Integer = i

                dgQuellZitate.Columns.RemoveAt(i)

                Dim linkColumn As New DataGridTemplateColumn()
                linkColumn.Header = "Internetadresse"

                ' -----------------------------------------------------
                ' Hyperlink
                ' -----------------------------------------------------
                Dim textBlockFactory As New FrameworkElementFactory(
                GetType(TextBlock))

                Dim hyperlinkFactory As New FrameworkElementFactory(
                GetType(Hyperlink))

                hyperlinkFactory.SetBinding(
                Hyperlink.NavigateUriProperty,
                New Binding("InternetAdresse"))

                hyperlinkFactory.SetBinding(
                Hyperlink.CommandProperty,
                New Binding("DataContext.OpenUrlCommand") With {
                    .RelativeSource = New RelativeSource(
                        RelativeSourceMode.FindAncestor,
                        GetType(DataGrid),
                        1)
                })

                Dim linkTextFactory As New FrameworkElementFactory(
                GetType(Run))

                linkTextFactory.SetBinding(
                Run.TextProperty,
                New Binding("InternetAdresse"))

                hyperlinkFactory.AppendChild(linkTextFactory)
                textBlockFactory.AppendChild(hyperlinkFactory)

                textBlockFactory.SetValue(
                TextBlock.VerticalAlignmentProperty,
                VerticalAlignment.Center)

                textBlockFactory.SetValue(
                TextBlock.MarginProperty,
                New Thickness(4, 0, 4, 0))

                Dim linkTemplate As New DataTemplate()
                linkTemplate.VisualTree = textBlockFactory

                linkColumn.CellTemplate = linkTemplate

                dgQuellZitate.Columns.Insert(index, linkColumn)

            End If

        Next

        dgQuellZitate.CanUserAddRows = False
        dgQuellZitate.CanUserDeleteRows = False

        isUCLoaded = True

    End Sub

    Private Sub btnNeu_Click(sender As Object, e As RoutedEventArgs)
        Dim neueZeile As DataRow = dt.NewRow()

        neueZeile("tblQuelleID") = 0
        neueZeile("tblEreignisArtID") = 0
        neueZeile("Jahr") = DBNull.Value
        neueZeile("Bd") = ""
        neueZeile("Seite") = ""
        neueZeile("Nummer") = ""
        neueZeile("Datum") = DBNull.Value
        neueZeile("InternetAdresse") = ""
        neueZeile("URLBeschreibung") = ""
        neueZeile("ZitatBeschreibung") = ""
        neueZeile("active") = True
        neueZeile("Anzahl") = 0

        dt.Rows.Add(neueZeile)

        ' Neue Zeile im DataGrid auswählen
        Dim neueZeileView As DataRowView = dt.DefaultView(dt.DefaultView.Count - 1)

        dgQuellZitate.SelectedItem = neueZeileView
        dgQuellZitate.ScrollIntoView(neueZeileView)
    End Sub

    Private Sub btnSpeichern_Click(sender As Object, e As RoutedEventArgs)
        If dgQuellZitate.SelectedItem Is Nothing Then
            MessageBox.Show("Bitte eine Zeile auswählen.")
            Return
        End If

        Dim rowView As DataRowView =
        TryCast(dgQuellZitate.SelectedItem, DataRowView)

        If rowView Is Nothing Then Return



        Dim row As DataRow = rowView.Row

        If row("tblQuelleID") = 0 Then
            MessageBox.Show("Bitte eine Quelle auswählen.")
            Return
        End If

        If row("tblEreignisArtID") = 0 Then
            MessageBox.Show("Bitte eine Ereignisart auswählen.")
            Return
        End If

        Try

            ' Neue Zeile
            If IsDBNull(row("tblQuellZitatID")) OrElse
           Convert.ToInt32(row("tblQuellZitatID")) = 0 Then

                Dim neueID As Integer = cGenDB.SetQuellZitat(row("tblQuelleID"), row("tblEreignisArtID"),
                                        IIf(IsDBNull(row("Jahr")), Nothing, row("Jahr")),
                                        IIf(IsDBNull(row("Seite")), "", row("Seite")),
                                        IIf(IsDBNull(row("Bd")), "", row("Bd")),
                                        IIf(IsDBNull(row("Nummer")), "", row("Nummer")),
                                        IIf(IsDBNull(row("Datum")), Nothing, row("Datum")),
                                        IIf(IsDBNull(row("InternetAdresse")), "", row("InternetAdresse")),
                                        IIf(IsDBNull(row("URLBeschreibung")), "", row("URLBeschreibung")),
                                        IIf(IsDBNull(row("ZitatBeschreibung")), "", row("ZitatBeschreibung")))
                row("tblQuellZitatID") = neueID
            Else

                ' Bestehende Zeile
                cGenDB.UpdateQuellZitat(row("tblQuellZitatID"), row("tblQuelleID"), row("tblEreignisArtID"),
                                        IIf(IsDBNull(row("Jahr")), Nothing, row("Jahr")),
                                        IIf(IsDBNull(row("Seite")), "", row("Seite")),
                                        IIf(IsDBNull(row("Bd")), "", row("Bd")),
                                        IIf(IsDBNull(row("Nummer")), "", row("Nummer")),
                                        IIf(IsDBNull(row("Datum")), Nothing, row("Datum")),
                                        IIf(IsDBNull(row("InternetAdresse")), "", row("InternetAdresse")),
                                        IIf(IsDBNull(row("URLBeschreibung")), "", row("URLBeschreibung")),
                                        IIf(IsDBNull(row("ZitatBeschreibung")), "", row("ZitatBeschreibung")))

            End If

            MessageBox.Show("Gespeichert.")

            ' Daten neu laden
            ' LoadData()

        Catch ex As Exception

            MessageBox.Show(
            "Fehler beim Speichern:" & vbCrLf &
            ex.Message)

        End Try

    End Sub

    Private Sub btnKopieren_Click(sender As Object, e As RoutedEventArgs)
        ' Aktuell ausgewählte Zeile ermitteln
        Dim alteZeileView As DataRowView =
            TryCast(dgQuellZitate.SelectedItem, DataRowView)

        If alteZeileView Is Nothing Then
            MessageBox.Show("Bitte zuerst eine Zeile auswählen.")
            Return
        End If

        Dim alteZeile As DataRow = alteZeileView.Row

        ' Neue Zeile erzeugen
        Dim neueZeile As DataRow = dt.NewRow()

        ' ============================================================
        ' Werte aus vorheriger Zeile übernehmen
        ' ============================================================

        neueZeile("tblQuelleID") = alteZeile("tblQuelleID")
        neueZeile("tblEreignisArtID") = alteZeile("tblEreignisArtID")

        neueZeile("Jahr") = alteZeile("Jahr")
        neueZeile("Bd") = alteZeile("Bd")
        neueZeile("Seite") = alteZeile("Seite")
        Dim nummer As String = If(IsDBNull(alteZeile("Nummer")), "", alteZeile("Nummer"))
        If IsNumeric(nummer) Then
            Dim länge As Long = nummer.Length
            nummer = String.Format("{0:D" & länge & "}", (CInt(nummer) + 1))
        End If


        neueZeile("Nummer") = nummer

        'neueZeile("Datum") = alteZeile("Datum")

        neueZeile("InternetAdresse") = alteZeile("InternetAdresse")
        neueZeile("URLBeschreibung") = alteZeile("URLBeschreibung")
        'neueZeile("ZitatBeschreibung") = alteZeile("ZitatBeschreibung")

        neueZeile("active") = True

        ' Anzahl gehört nicht zur Tabelle tblQuellZitat,
        ' sondern kommt aus der SQL-Abfrage
        neueZeile("Anzahl") = 0

        ' ============================================================
        ' ID zurücksetzen → neue Datenbankzeile
        ' ============================================================

        neueZeile("tblQuellZitatID") = 0

        ' ============================================================
        ' Neue Zeile hinzufügen
        ' ============================================================

        dt.Rows.Add(neueZeile)

        ' ============================================================
        ' Neue Zeile im DataGrid auswählen
        ' ============================================================

        Dim neueZeileView As DataRowView =
            dt.DefaultView(dt.DefaultView.Count - 1)

        dgQuellZitate.SelectedItem = neueZeileView

        dgQuellZitate.ScrollIntoView(neueZeileView)
    End Sub

    Private Sub LoadFilter()

        cbQuelle.SelectedValue = My.Settings.QZQuelle
        cbEreignisArt.SelectedValue = My.Settings.QZEreignis
        txtJahr.Text = My.Settings.QZJahr
        txtBd.Text = My.Settings.QZBd
        txtSeite.Text = My.Settings.QZSeite
        txtNummer.Text = My.Settings.QZNummer

    End Sub

    Public Function CopyFormData() As ClsFormularDatenCopy Implements IFormularClipboard.CopyFormData
        If dgQuellZitate.SelectedItem Is Nothing Then
            MessageBox.Show("Bitte eine Zeile auswählen.")
            Exit Function
        End If
        If EreignisControl IsNot Nothing Then
            EreignisControl.setQuellZitatID(CType(dgQuellZitate.SelectedItem, DataRowView)("tblQuellZitatID"))
            My.Settings.LastQuellZitat = CType(dgQuellZitate.SelectedItem, DataRowView)("tblQuellZitatID")
            My.Settings.Save()
        End If

        Dim daten As New ClsFormularDatenCopy()
        daten.Datum = CType(dgQuellZitate.SelectedItem, DataRowView)("Datum")
        daten.DatumBis = ""
        daten.OrtID = Nothing

        Return daten
    End Function

    Public Sub PasteFormData(daten As ClsFormularDatenCopy) Implements IFormularClipboard.PasteFormData
        If daten Is Nothing Then
            Return
        End If
        If dgQuellZitate.SelectedItem Is Nothing Then
            MessageBox.Show("Bitte eine Zeile auswählen.")
            Return
        End If

        CType(dgQuellZitate.SelectedItem, DataRowView)("Datum") = daten.Datum

    End Sub

    Private Sub QuellenZitate_Loaded(sender As Object, e As RoutedEventArgs)
        ClsFormularClipboard.SetActiveForm(Me)
    End Sub

    Private Sub QuellenZitate_GotFocus(sender As Object, e As RoutedEventArgs)
        ClsFormularClipboard.SetActiveForm(Me)
    End Sub
End Class
