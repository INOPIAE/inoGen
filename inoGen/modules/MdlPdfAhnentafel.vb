Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports inoGenDLL
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




Module MdlPdfAhnentafel
    Dim cAT As New clsAhnentafelDaten(My.Settings.DBPath)
    Public Class Person
        Public Property Vorname As String
        Public Property Nachname As String
        Public Property Geburt As String
        Public Property Taufe As String
        Public Property Tod As String
        Public Property Begräbnis As String
    End Class

    Public Structure PPos
        Public Pos As Integer
        Public Level As Integer
        Public PosFaktor As Integer
        Public VonFaktor As Integer
    End Structure

    Public PPositionen As New List(Of PPos)

    Public Structure PdfCanvasData
        Public Canvas As PdfCanvas
        Public RootPos As Integer
        Public PageNummer As Integer
    End Structure

    Public MyCanvas As New List(Of PdfCanvasData)

    Sub AT(Persons As List(Of clsAhnentafelDaten.PersonData), Datei As String)


        InitPositionen()

        Using writer As New PdfWriter(Datei)
            Using pdf As New PdfDocument(writer)
                ' Dim page = pdf.AddNewPage(PageSize.A4) ' Hochformat
                Dim doc As New Document(pdf)

                InitCanvas(pdf, 1)

                Dim font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)
                Dim canvas = MyCanvas(0).Canvas
                ' Kästchengröße
                Dim w As Single = 135
                Dim h As Single = 48
                Dim a As Single = 50

                Dim e As Integer = 0

                Dim rect As Rectangle
                Dim PageNummer As Integer = 1

                For Each personData In Persons
                    Dim pos As Int128 = personData.Pos - 1
                    If personData.Gen > 3 Then
                        canvas = GetCanvas(pdf, cAT.CalculateChild(personData), PageNummer)
                        pos = cAT.CalculateChildPosChart(personData)
                    End If
                    rect = DrawPerson(canvas, font, personData, 10 + (w + 10) * PPositionen(pos).Level, PPositionen(pos).PosFaktor * a, w, h)
                    If PPositionen(pos).Level > 0 Then
                        Connect(canvas, (w + 10) * PPositionen(pos).Level - 1, PPositionen(pos).VonFaktor * a, 10 + (w + 10) * PPositionen(pos).Level, PPositionen(pos).PosFaktor * a, w, h)
                    End If
                    If (personData.Heiratdatum <> "" Or personData.KHeiratdatum <> "" Or personData.Scheidungsdatum <> "") And personData.Pos Mod 2 = 0 Then
                        DrawHeirat(canvas, font, personData, 10 + (w + 10) * PPositionen(pos).Level, PPositionen((pos + 1) / 2 - 1).PosFaktor * a, w, h)
                    End If
                    If personData.Gen Mod 3 = 1 And pos >= 7 Then
                        If personData.FID > 0 Then
                            Dim fromCanvas As PdfCanvas = canvas
                            Dim fromRect As Rectangle = rect
                            Dim ToPage As Integer
                            InitCanvas(pdf, personData.Pos)
                            canvas = GetCanvas(pdf, personData.Pos, ToPage)
                            rect = DrawPerson(canvas, font, personData, 10 + (w + 10) * PPositionen(0).Level, PPositionen(0).PosFaktor * a, w, h)
                            CreateLinks(pdf, PageNummer, fromRect, ToPage, rect)
                        End If
                    End If
                Next
                doc.Close()
            End Using
        End Using
    End Sub


    Private Function DrawPerson(canvas As PdfCanvas, font As PdfFont, person As clsAhnentafelDaten.PersonData, x As Single, y As Single, w As Single, h As Single) As Rectangle
        canvas.SetStrokeColor(ColorConstants.BLACK)
        If person.FID > 0 And (person.Gen = 4 Or person.Gen = 7 Or person.Gen = 10) Then
            canvas.SetLineDash(3, 3)
        End If
        canvas.Rectangle(x, y, w, h)
        canvas.Stroke()

        Dim box As New Rectangle(x, y, w, h)
        Dim docCanvas As New Canvas(canvas, box)
        docCanvas.Add(New Paragraph(person.Vorname).
            SetFont(font).
            SetFontSize(8).
            SetTextAlignment(TextAlignment.CENTER).
            SetMarginTop(1).
            SetMarginBottom(0).
            SetMultipliedLeading(1))

        docCanvas.Add(New Paragraph(person.Nachname.ToUpper).
            SetFont(font).
            SetFontSize(9).
            SetTextAlignment(TextAlignment.CENTER).
            SetMarginTop(0).
            SetMarginBottom(0).
            SetMultipliedLeading(1))

        ' Geburt
        If person.Geburtsdatum <> "" Then
            docCanvas.Add(New Paragraph("* " & person.Geburtsdatum & " " & person.Geburtsort).
                SetFont(font).
                SetFontSize(6).
                SetTextAlignment(TextAlignment.LEFT).
                SetMarginTop(0).
                SetMarginLeft(3).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If
        If person.Taufdatum <> "" Then
            docCanvas.Add(New Paragraph("~ " & person.Taufdatum & " " & person.Taufort).
                SetFont(font).
                SetFontSize(6).
                SetTextAlignment(TextAlignment.LEFT).
                SetMarginTop(0).
                SetMarginLeft(3).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If

        ' Tod
        If person.Sterbedatum <> "" Then
            docCanvas.Add(New Paragraph("+ " & person.Sterbedatum & " " & person.Sterbeort).
                SetFont(font).
                SetFontSize(6).
                SetTextAlignment(TextAlignment.LEFT).
                SetMarginTop(0).
                SetMarginLeft(3).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If

        If person.Begräbnisdatum <> "" Then
            docCanvas.Add(New Paragraph("✝ " & person.Begräbnisdatum & " " & person.Begräbnisort).
                SetFont(font).
                SetFontSize(6).
                SetTextAlignment(TextAlignment.LEFT).
                SetMarginTop(0).
                SetMarginLeft(3).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If
        docCanvas.Close()
        canvas.SetLineDash(0)
        Return box
    End Function

    ' Linie von Eltern-Kasten unten zur Mitte von Kind-Kasten oben
    Private Sub Connect(canvas As PdfCanvas, x1 As Single, y1 As Single, x2 As Single, y2 As Single, w As Single, h As Single)
        Dim cx1 = x1 + 1
        Dim cy1 = y1 + h / 2
        Dim cx2 = x1 + 5
        Dim cy2 = y1 + h / 2
        Dim cx3 = x1 + 5
        Dim cy3 = y2 + h / 2
        Dim cx4 = x2
        Dim cy4 = y2 + h / 2
        canvas.MoveTo(cx1, cy1).LineTo(cx2, cy2).Stroke()
        canvas.MoveTo(cx2, cy2).LineTo(cx3, cy3).Stroke()
        canvas.MoveTo(cx3, cy3).LineTo(cx4, cy4).Stroke()
    End Sub

    Private Sub DrawHeirat(canvas As PdfCanvas, font As PdfFont, person As clsAhnentafelDaten.PersonData, x As Single, y As Single, w As Single, h As Single)


        Dim box As New Rectangle(x, y - h / 2, w, h)
        Dim docCanvas As New Canvas(canvas, box)

        If person.Heiratdatum <> "" Then
            docCanvas.Add(New Paragraph("oo " & person.Heiratdatum & " " & person.Heiratort).
                SetFont(font).
                SetFontSize(6).
                SetTextAlignment(TextAlignment.LEFT).
                SetMarginTop(0).
                SetMarginLeft(3).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If
        If person.KHeiratdatum <> "" Then
            docCanvas.Add(New Paragraph("oo+ " & person.KHeiratdatum & " " & person.KHeiratort).
                SetFont(font).
                SetFontSize(6).
                SetTextAlignment(TextAlignment.LEFT).
                SetMarginTop(0).
                SetMarginLeft(3).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If
        If person.Scheidungsdatum <> "" Then
            docCanvas.Add(New Paragraph("✝ " & person.Scheidungsdatum & " " & person.Scheidungsort).
                SetFont(font).
                SetFontSize(6).
                SetTextAlignment(TextAlignment.LEFT).
                SetMarginTop(0).
                SetMarginLeft(3).
                SetMarginBottom(0).
                SetMultipliedLeading(1))
        End If

        docCanvas.Close()
        canvas.SetLineDash(0)
    End Sub
    Private Sub InitPositionen()
        Dim p As New PPos
        p.Pos = 1
        p.Level = 0
        p.PosFaktor = 8
        p.VonFaktor = 0
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 2
        p.Level = 1
        p.PosFaktor = 12
        p.VonFaktor = 8
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 3
        p.Level = 1
        p.PosFaktor = 4
        p.VonFaktor = 8
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 4
        p.Level = 2
        p.PosFaktor = 14
        p.VonFaktor = 12
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 5
        p.Level = 2
        p.PosFaktor = 10
        p.VonFaktor = 12
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 6
        p.Level = 2
        p.PosFaktor = 6
        p.VonFaktor = 4
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 7
        p.Level = 2
        p.PosFaktor = 2
        p.VonFaktor = 4
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 8
        p.Level = 3
        p.PosFaktor = 15
        p.VonFaktor = 14
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 9
        p.Level = 3
        p.PosFaktor = 13
        p.VonFaktor = 14
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 10
        p.Level = 3
        p.PosFaktor = 11
        p.VonFaktor = 10
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 11
        p.Level = 3
        p.PosFaktor = 9
        p.VonFaktor = 10
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 12
        p.Level = 3
        p.PosFaktor = 7
        p.VonFaktor = 6
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 13
        p.Level = 3
        p.PosFaktor = 5
        p.VonFaktor = 6
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 14
        p.Level = 3
        p.PosFaktor = 3
        p.VonFaktor = 2
        PPositionen.Add(p)

        p = New PPos
        p.Pos = 15
        p.Level = 3
        p.PosFaktor = 1
        p.VonFaktor = 2
        PPositionen.Add(p)

    End Sub

    Private Sub InitCanvas(pdf As PdfDocument, Root As Integer)
        Dim page = pdf.AddNewPage(PageSize.A4) ' Hochformat
        Dim canvas As New PdfCanvas(page)


        canvas.SetLineWidth(0.8F)
        Dim pcd As New PdfCanvasData
        pcd.Canvas = canvas
        pcd.RootPos = Root
        pcd.PageNummer = pdf.GetNumberOfPages()
        MyCanvas.Add(pcd)
    End Sub

    Private Function GetCanvas(pdf As PdfDocument, Root As Integer, ByRef PageNumber As Integer) As PdfCanvas
        For Each c In MyCanvas
            If c.RootPos = Root Then
                PageNumber = c.PageNummer
                Return c.Canvas
            End If
        Next
        'InitCanvas(pdf, Root)

        'Return MyCanvas(MyCanvas.Count - 1).Canvas
        Return Nothing
    End Function

    Private Sub CreateLinks(pdf As PdfDocument, fromP As Integer, fromR As Rectangle, toP As Integer, toR As Rectangle)

        Dim page1 As PdfPage = pdf.GetPage(fromP)

        ' --- Seite 2 ---
        Dim page2 As PdfPage = pdf.GetPage(toP)
        ' Ziele definieren
        Dim destPage1 As PdfDestination = PdfExplicitDestination.CreateXYZ(page1, fromR.GetLeft, fromR.GetTop, 1)
        Dim destPage2 As PdfDestination = PdfExplicitDestination.CreateXYZ(page2, toR.GetLeft, toR.GetTop, 1)

        ' Link von Seite 1 → Seite 2 (unsichtbar)
        Dim link1 As New PdfLinkAnnotation(fromR)
        link1.SetAction(PdfAction.CreateGoTo(destPage2))
        link1.SetBorder(New PdfArray({0, 0, 0})) ' Kein Rahmen
        page1.AddAnnotation(link1)

        ' Link von Seite 2 → Seite 1 (unsichtbar)
        Dim link2 As New PdfLinkAnnotation(toR)
        link2.SetAction(PdfAction.CreateGoTo(destPage1))
        link2.SetBorder(New PdfArray({0, 0, 0})) ' Kein Rahmen
        page2.AddAnnotation(link2)
    End Sub
End Module
