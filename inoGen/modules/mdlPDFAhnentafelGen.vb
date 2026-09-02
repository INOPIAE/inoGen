Imports System.IO
Imports System.Windows.Forms
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports inoGen.MdlPdfAhnentafel
Imports inoGenDLL
Imports inoGenDLL.clsAhnentafelDaten
Imports inoGenDLL.ClsAhnentafelPDF
Imports iText.IO.Codec.Brotli
Imports iText.IO.Font
Imports iText.IO.Font.Constants
Imports iText.Kernel.Colors
Imports iText.Kernel.Font
Imports iText.Kernel.Geom
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Action
Imports iText.Kernel.Pdf.Annot
Imports iText.Kernel.Pdf.Canvas
Imports iText.Kernel.Pdf.Event
Imports iText.Kernel.Pdf.Navigation
Imports iText.Layout
Imports iText.Layout.Borders
Imports iText.Layout.Element
Imports iText.Layout.Properties
Module mdlPDFAhnentafelGen
    Private cAP As New ClsAhnentafelPDF
    Private dinFormat As DinFormat
    Public Sub PrintAhnentafelBlankoAll()
        ' Layout für A4 erstellen
        Dim layout As New ClsAhnentafelLayout(PdfPageSizes.sizes(DinFormat.A1), True)
        ' Dim layout As New ClsAhnentafelLayout((2384, 1684), false)
        'layout.PrintLayout()

        ' PDF erstellen
        Using writer As New PdfWriter("D:\test\ahnentafel.pdf")
            Using pdfDoc As New PdfDocument(writer)
                pdfDoc.SetDefaultPageSize(New PageSize(layout.PageWidth, layout.PageHeight))

                Dim Canvas As New PdfCanvas(pdfDoc.AddNewPage())
                Dim font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)

                ' Titelblock zeichnen
                Dim titleBlock = layout.TitleBlockArea
                Canvas.SetStrokeColor(ColorConstants.BLUE)
                Canvas.Rectangle(titleBlock.GetX(), titleBlock.GetY(),
                                   titleBlock.GetWidth(), titleBlock.GetHeight())
                Canvas.Stroke()

                ' Alle Kästchen zeichnen (15x15 Raster)
                For row As Integer = 0 To 14
                    For col As Integer = 0 To 14
                        Dim box = layout.GetBox(row, col)

                        ' Zeile 8 und Spalte 8 hervorheben
                        If row = 7 OrElse col = 7 Then
                            Canvas.SetStrokeColor(ColorConstants.RED)
                            Canvas.SetLineWidth(2)
                        Else
                            Canvas.SetStrokeColor(ColorConstants.BLACK)
                            Canvas.SetLineWidth(1)
                        End If

                        Canvas.Rectangle(box.GetX(), box.GetY(),
                                           box.GetWidth(), box.GetHeight())
                        Canvas.Stroke()

                        ' Position als Text schreiben
                        Canvas.BeginText()
                        Canvas.SetFontAndSize(font, 6)
                        Canvas.MoveText(box.GetX() + 2, box.GetY() + box.GetHeight() / 2)
                        Canvas.ShowText($"{row + 1},{col + 1}")
                        Canvas.EndText()
                    Next
                Next

                ' Referenzlinien zeichnen (zur Kontrolle)
                Canvas.SetStrokeColor(ColorConstants.LIGHT_GRAY)
                Canvas.SetLineWidth(0.5)

                ' Vertikale Linien (Spalten)
                For Each lineX In layout.LinePositions
                    Canvas.MoveTo(lineX, layout.MarginBottom)
                    Canvas.LineTo(lineX, layout.PageHeight - layout.MarginTop - layout.TitleBlockHeight)
                    Canvas.Stroke()
                Next

                ' Horizontale Linien (Zeilen)
                For Each lineY In layout.RowPositions
                    Canvas.MoveTo(layout.MarginLeft, lineY)
                    Canvas.LineTo(layout.PageWidth - layout.MarginRight, lineY)
                    Canvas.Stroke()
                Next

                ' Mittelpunkte markieren
                Canvas.SetFillColor(ColorConstants.RED)
                Canvas.Circle(layout.PageWidth / 2,
                                 layout.MarginBottom + layout.UsableHeight / 2, 3)
                Canvas.Fill()

            End Using
        End Using

        Console.WriteLine("PDF erstellt: D:\test\ahnentafel.pdf")
        MessageBox.Show("fertig")
    End Sub

    Public Sub PrintAhnentafelBlanko()
        ' Layout für A4 erstellen
        '   Dim layout As New ClsAhnentafelLayout(ClsAhnentafelPDF.PdfPageSizes.A1, True)
        Dim layout As New ClsAhnentafelLayout((2384, 1684), False)

        'layout.PrintLayout()

        Dim fileName As String = "D:\test\testLayout.log"
        Try
            Kill(fileName)
        Catch ex As Exception

        End Try

        Dim AhnenBox As Dictionary(Of Integer, Koordinaten) = cAP.CreateAhnentafelDictionaryGen7()

        ' PDF erstellen
        Using writer As New PdfWriter("D:\test\ahnentafelblanko.pdf")
            Using pdfDoc As New PdfDocument(writer)
                pdfDoc.SetDefaultPageSize(New PageSize(layout.PageWidth, layout.PageHeight))

                Dim Canvas As New PdfCanvas(pdfDoc.AddNewPage())
                Dim font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)

                ' Titelblock zeichnen
                Dim titleBlock = layout.TitleBlockArea
                Canvas.SetStrokeColor(ColorConstants.BLUE)
                Canvas.Rectangle(titleBlock.GetX(), titleBlock.GetY(),
                                   titleBlock.GetWidth(), titleBlock.GetHeight())
                Canvas.Stroke()


                For Each kvp In AhnenBox
                    Dim personNr As Integer = kvp.Key
                    Dim row As Integer = kvp.Value.KZeile - 1
                    Dim col As Integer = kvp.Value.KSpalte - 1

                    ' Box-Position ermitteln
                    Dim box = layout.GetBox(row, col)

                    ' Rechteck zeichnen
                    Canvas.SetStrokeColor(ColorConstants.BLACK)
                    Canvas.Rectangle(box.GetX(), box.GetY(), box.GetWidth(), box.GetHeight())
                    Canvas.Stroke()

                    ' Nummer eintragen
                    Canvas.BeginText()
                    Canvas.SetFontAndSize(font, 8)
                    Canvas.MoveText(box.GetX() + 2, box.GetY() + box.GetHeight() - 10)
                    Canvas.ShowText(personNr.ToString())
                    Canvas.EndText()

                    If personNr > 1 Then
                        Dim VG As Long = personNr \ 2
                        Dim VGR As Boolean = (personNr Mod 2) = 0

                        File.AppendAllText(fileName, "Persnr " & personNr & " VG " & VG & " " & VGR & Environment.NewLine)

                        ' Vorgänger-Koordinaten holen
                        If AhnenBox.ContainsKey(VG) Then
                            Dim koordVG = AhnenBox(VG)
                            Dim rowVG As Integer = koordVG.KZeile - 1
                            Dim colVG As Integer = koordVG.KSpalte - 1

                            ' Linie vom aktuellen zum Vorgänger zeichnen
                            ' WICHTIG: col = X-Position (Spalte), row = Y-Position (Zeile)
                            Canvas.SetStrokeColor(ColorConstants.BLACK)
                            Canvas.MoveTo(layout.LinePositions(col), layout.RowPositions(row))
                            Canvas.LineTo(layout.LinePositions(colVG), layout.RowPositions(rowVG))
                            Canvas.Stroke()
                        End If
                    End If
                Next


                ' Referenzlinien zeichnen (zur Kontrolle)
                Canvas.SetStrokeColor(ColorConstants.LIGHT_GRAY)
                Canvas.SetLineWidth(0.5)

                '' Vertikale Linien (Spalten)
                'For Each lineX In layout.LinePositions
                '    Canvas.MoveTo(lineX, layout.MarginBottom)
                '    Canvas.LineTo(lineX, layout.PageHeight - layout.MarginTop - layout.TitleBlockHeight)
                '    Canvas.Stroke()
                'Next

                '' Horizontale Linien (Zeilen)
                'For Each lineY In layout.RowPositions
                '    Canvas.MoveTo(layout.MarginLeft, lineY)
                '    Canvas.LineTo(layout.PageWidth - layout.MarginRight, lineY)
                '    Canvas.Stroke()
                'Next

                '' Mittelpunkte markieren
                'Canvas.SetFillColor(ColorConstants.RED)
                'Canvas.Circle(layout.PageWidth / 2,
                '                 layout.MarginBottom + layout.UsableHeight / 2, 3)
                'Canvas.Fill()

            End Using
        End Using

        Console.WriteLine("PDF erstellt: D:\test\ahnentafel.pdf")
        MessageBox.Show("fertig")
    End Sub

    Public Function PrintAhnentafelGen7(Persons As List(Of clsAhnentafelDaten.PersonData), pdfFilename As String) As Boolean


        Select Case My.Settings.LastGenPapersize
            Case "A0"
                dinFormat = DinFormat.A0
            Case "A1"
                dinFormat = DinFormat.A1
            Case "A2"
                dinFormat = DinFormat.A2
            Case Else
                dinFormat = DinFormat.A3
                MessageBox.Show("A3 wird nicht unterstützt, bitte A2 oder größer wählen.", "Hinweis", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
        End Select

        Dim layout As New ClsAhnentafelLayout(PdfPageSizes.sizes(dinFormat), True)
        'Dim layout As New ClsAhnentafelLayout((2384, 1684))

        layout.PrintLayout()

        Dim AhnenBox As Dictionary(Of Integer, Koordinaten) = cAP.CreateAhnentafelDictionaryGen7()

        Dim filename As String = "D:\test\testLayout.log"
        Try
            Kill(filename)
        Catch ex As Exception

        End Try


        ' PDF erstellen
        Using writer As New PdfWriter(pdfFilename)
            Using pdfDoc As New PdfDocument(writer)
                pdfDoc.SetDefaultPageSize(New PageSize(layout.PageWidth, layout.PageHeight))

                Dim Canvas As New PdfCanvas(pdfDoc.AddNewPage())
                Dim font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)

                ' Titelblock zeichnen
                Dim titleBlock = layout.TitleBlockArea
                Dim titleCanvas As New Canvas(Canvas, titleBlock)

                ' Haupttitel
                'titleCanvas.Add(New Paragraph("Ahnentafel").
                '    SetFont(font).
                '    SetFontSize(18).
                '    SetFontColor(ColorConstants.BLACK).
                '    SetTextAlignment(TextAlignment.CENTER).
                '    SetMarginTop(10).
                '    SetMarginBottom(2))

                '' Name
                'titleCanvas.Add(New Paragraph("für " & Persons(0).Vorname & " " & Persons(0).Nachname).
                '    SetFont(font).
                '    SetFontSize(14).
                '    SetFontColor(ColorConstants.BLACK).
                '    SetTextAlignment(TextAlignment.CENTER).
                '    SetMarginTop(0).
                '    SetMarginBottom(2))

                titleCanvas.Add(New Paragraph("Ahnentafel für " & Persons(0).Vorname & " " & Persons(0).Nachname).
                    SetFont(font).
                    SetFontSize(50).
                    SetFontColor(ColorConstants.BLACK).
                    SetTextAlignment(TextAlignment.CENTER).
                    SetMarginTop(0).
                    SetMarginBottom(2))

                ' Datum (optional)
                titleCanvas.Add(New Paragraph($"Erstellt am {DateTime.Now:dd.MM.yyyy}").
                    SetFont(font).
                    SetFontSize(8).
                    SetFontColor(ColorConstants.GRAY).
                    SetTextAlignment(TextAlignment.CENTER).
                    SetMarginTop(5))

                titleCanvas.Close()

                Canvas.SaveState()
                Canvas.SetStrokeColor(ColorConstants.BLACK)
                Canvas.SetLineWidth(0.5)

                For Each personData In Persons
                    Dim kekule As Long = personData.Pos
                    If kekule > 1 And kekule < 128 Then

                        Dim koordP = AhnenBox(kekule)
                        Dim koordVG = AhnenBox(kekule \ 2)

                        Canvas.MoveTo(layout.LinePositions(koordP.KSpalte - 1), layout.RowPositions(koordP.KZeile - 1))
                        Canvas.LineTo(layout.LinePositions(koordVG.KSpalte - 1), layout.RowPositions(koordVG.KZeile - 1))
                        Canvas.Stroke()
                    End If
                Next

                Canvas.RestoreState()


                For Each personData In Persons
                    Dim kekule As Long = personData.Pos
                    If kekule < 128 AndAlso AhnenBox.ContainsKey(kekule) Then
                        Dim koord = AhnenBox(kekule)
                        Dim box = layout.GetBox(koord.KZeile - 1, koord.KSpalte - 1)

                        '            File.AppendAllText(filename, kekule & Environment.NewLine)
                        ' Person in Box zeichnen
                        DrawPerson(Canvas, font, personData,
                           box.GetX(), box.GetY(),
                           box.GetWidth(), box.GetHeight())
                        '            File.AppendAllText(filename, personData.Vorname & " " & personData.Nachname & Environment.NewLine)

                    End If
                Next
            End Using
        End Using

        Return True
    End Function

    Private Function DrawPerson(canvas As PdfCanvas, font As PdfFont, person As clsAhnentafelDaten.PersonData, x As Single, y As Single, w As Single, h As Single) As Rectangle
        ' Füllfarbe nach Geschlecht

        Dim isInZweig4 As Boolean = IsInAhnenZweig(person.Pos, 5)

        Dim fillColor As DeviceRgb

        Select Case My.Settings.LastGenColortype
            Case "Geschlecht"
                fillColor = GetGenderColor(person.Geschlecht)
            Case "Zweig"
                Dim generation As Integer = GetGeneration(person.Pos)
                If IsInAhnenZweig(person.Pos, 4) Then
                    fillColor = GetZweigColor(generation, 4)
                ElseIf IsInAhnenZweig(person.Pos, 5) Then
                    fillColor = GetZweigColor(generation, 5)
                ElseIf IsInAhnenZweig(person.Pos, 6) Then
                    fillColor = GetZweigColor(generation, 6)
                ElseIf IsInAhnenZweig(person.Pos, 7) Then
                    fillColor = GetZweigColor(generation, 7)
                Else
                    fillColor = GetBasisColor(person.Pos)
                End If
            Case Else
                fillColor = ColorConstants.WHITE
        End Select

        ' Box mit Füllung zeichnen
        canvas.SetLineWidth(1)
        canvas.SetStrokeColor(ColorConstants.BLACK)
        canvas.SetFillColor(fillColor)
        canvas.Rectangle(x, y, w, h)
        canvas.FillStroke()

        If person.Pos >= 64 And person.FID > 0 Then
            canvas.SetLineWidth(0.5)
            canvas.Rectangle(x + 1, y + 1, w - 2, h - 2)
            canvas.Stroke()
            canvas.SetLineWidth(1)
        End If

        ' Canvas für Text-Layout
        Dim box As New Rectangle(x, y, w, h)
        Dim docCanvas As New Canvas(canvas, box)

        Dim DinFaktor As Single = 1.0F
        Select Case dinFormat
            Case DinFormat.A0
                DinFaktor = 1.3F
            Case DinFormat.A1
                DinFaktor = 1.1F
            Case DinFormat.A2
                DinFaktor = 1.0F
        End Select
        ' Name
        AddCenteredText(docCanvas, font, person.Vorname, 9 * DinFaktor, 2)
        AddCenteredText(docCanvas, font, person.Nachname.ToUpper, 10 * DinFaktor, 0)

        ' Lebensdaten
        AddLifeEvent(docCanvas, font, "*", person.Geburtsdatum, person.Geburtsort)
        AddLifeEvent(docCanvas, font, "~", person.Taufdatum, person.Taufort)
        AddLifeEvent(docCanvas, font, "+", person.Sterbedatum, person.Sterbeort)
        AddLifeEvent(docCanvas, font, "✝", person.Begräbnisdatum, person.Begräbnisort)

        docCanvas.Close()
        Return box
    End Function

    Private Function GetGenderColor(geschlecht As String) As DeviceRgb
        Select Case geschlecht.ToUpper
            Case "M"
                Return New DeviceRgb(173.0F / 255.0F, 216.0F / 255.0F, 230.0F / 255.0F)  ' Hellblau
            Case "W"
                Return New DeviceRgb(255.0F / 255.0F, 182.0F / 255.0F, 193.0F / 255.0F)  ' Rosa
            Case Else
                Return ColorConstants.WHITE
        End Select
    End Function

    Private Sub AddCenteredText(canvas As Canvas, font As PdfFont, text As String, fontSize As Single, marginTop As Single)
        If Not String.IsNullOrEmpty(text) Then
            canvas.Add(New Paragraph(text).
                SetFont(font).
                SetFontSize(fontSize).
                SetFontColor(ColorConstants.BLACK).
                SetTextAlignment(TextAlignment.CENTER).
                SetMarginTop(marginTop).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If
    End Sub

    Private Sub AddLifeEvent(canvas As Canvas, font As PdfFont, symbol As String, datum As String, ort As String)
        If Not String.IsNullOrEmpty(datum) Then
            canvas.Add(New Paragraph(symbol & " " & datum & " " & ort).
                SetFont(font).
                SetFontSize(6).
                SetFontColor(ColorConstants.BLACK).
                SetTextAlignment(TextAlignment.LEFT).
                SetMarginTop(0).
                SetMarginLeft(3).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If
    End Sub

    ' Prüft ob eine Kekule-Nummer im Ahnenzweig einer bestimmten Person liegt
    Private Function IsInAhnenZweig(kekuleNr As Long, stammNr As Long) As Boolean
        If kekuleNr = stammNr Then Return True

        ' Gehe den Stammbaum nach oben (zu den Vorfahren)
        Dim current As Long = kekuleNr
        While current > stammNr
            current = current \ 2  ' Elternteil
            If current = stammNr Then Return True
        End While

        Return False
    End Function

    ' Ermittelt die Generation einer Kekule-Nummer (1=Gen1, 2-3=Gen2, 4-7=Gen3, etc.)
    Private Function GetGeneration(kekuleNr As Long) As Integer
        If kekuleNr = 0 Then Return 0
        Return CInt(Math.Floor(Math.Log(kekuleNr, 2))) + 1
    End Function


    Private Function GetZweigColor(generation As Integer, stammNr As Long) As DeviceRgb
        ' Basis: Dunkles Farbe für Generation 1
        ' Je höher die Generation, desto heller

        Dim colorStart As System.Drawing.Color = My.Settings.Gen74S
        Dim colorEnd As System.Drawing.Color = My.Settings.Gen74E

        Select Case stammNr
            Case 4
                colorStart = My.Settings.Gen74S
                colorEnd = My.Settings.Gen74E
            Case 5
                colorStart = My.Settings.Gen75S
                colorEnd = My.Settings.Gen75E
            Case 6
                colorStart = My.Settings.Gen76S
                colorEnd = My.Settings.Gen76E
            Case 7
                colorStart = My.Settings.Gen77S
                colorEnd = My.Settings.Gen77E
        End Select


        Dim baseR As Single = colorStart.R
        Dim baseG As Single = colorStart.G
        Dim baseB As Single = colorStart.B

        Dim targetR As Single = colorEnd.R
        Dim targetG As Single = colorEnd.G
        Dim targetB As Single = colorEnd.B

        ' Aufhellung pro Generation (max 7 Generationen)
        Dim maxGen As Integer = 7
        Dim factor As Single = Math.Min(1.0F, (generation - 1) / CSng(maxGen - 1))

        Dim r As Single = (baseR + (targetR - baseR) * factor) / 255.0F
        Dim g As Single = (baseG + (targetG - baseG) * factor) / 255.0F
        Dim b As Single = (baseB + (targetB - baseB) * factor) / 255.0F

        Return New DeviceRgb(r, g, b)
    End Function
    Private Function GetBasisColor(kekule As Long) As DeviceRgb
        Dim colorStart As System.Drawing.Color
        Select Case kekule
            Case 1
                colorStart = My.Settings.Gen71
            Case 2
                colorStart = My.Settings.Gen72
            Case 3
                colorStart = My.Settings.Gen73
        End Select

        Return New DeviceRgb(colorStart.R, colorStart.G, colorStart.B)
    End Function
End Module
