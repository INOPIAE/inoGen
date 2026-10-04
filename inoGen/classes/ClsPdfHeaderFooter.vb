Imports iText.IO.Font
Imports iText.Kernel.Font
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas
Imports iText.Kernel.Pdf.Event
Imports iText.Kernel.Geom
Imports iText.Layout
Imports iText.Layout.Element
Imports iText.Layout.Properties
Imports iText.Kernel.Pdf.Xobject

Public Class ClsPdfHeaderFooter
        Inherits AbstractPdfDocumentEventHandler


        Private ReadOnly _font As PdfFont

        Private ReadOnly _headerLeft As String
        Private ReadOnly _headerCenter As String
        Private ReadOnly _headerRight As String

        Private ReadOnly _footerLeft As String
        Private ReadOnly _footerCenter As String


        '==============================================================
        ' Platzhalter für Gesamtseitenzahl
        '==============================================================

        Private ReadOnly _totalPagesPlaceholder As PdfFormXObject


        Public Sub New(
        pdfDoc As PdfDocument,
        Optional headerLeft As String = "",
        Optional headerCenter As String = "",
        Optional headerRight As String = "",
        Optional footerLeft As String = "",
        Optional footerCenter As String = "")


            _headerLeft = headerLeft
            _headerCenter = headerCenter
            _headerRight = headerRight

            _footerLeft = footerLeft
            _footerCenter = footerCenter


            '==========================================================
            ' Schrift
            '==========================================================

            _font =
            PdfFontFactory.CreateFont(
                "C:\Windows\Fonts\segoeui.ttf",
                PdfEncodings.IDENTITY_H,
                PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED)


            '==========================================================
            ' Platzhalter für Gesamtseitenzahl
            '==========================================================

            _totalPagesPlaceholder =
            New PdfFormXObject(
                New Rectangle(0, 0, 25, 12))

        End Sub


        '==============================================================
        ' Seitenende
        '==============================================================

        Protected Overrides Sub OnAcceptedEvent(
        pdfEvent As AbstractPdfDocumentEvent)


            Dim documentEvent As PdfDocumentEvent =
            CType(pdfEvent, PdfDocumentEvent)

            Dim pdfDoc As PdfDocument =
            documentEvent.GetDocument()

            Dim page As PdfPage =
            documentEvent.GetPage()


            If page Is Nothing Then
                Return
            End If


            Dim pageSize As Rectangle =
            page.GetPageSize()


            Dim pageNumber As Integer =
            pdfDoc.GetPageNumber(page)


            '==========================================================
            ' Canvas
            '==========================================================

            Dim pdfCanvas As New PdfCanvas(
            page.NewContentStreamAfter(),
            page.GetResources(),
            pdfDoc)

            Dim canvas As New Canvas(
            pdfCanvas,
            pageSize)


            '==========================================================
            ' Kopfzeile
            '==========================================================

            If _headerLeft <> "" Then

                AddText(
                canvas,
                _headerLeft,
                pageSize.GetLeft() + 40,
                pageSize.GetTop() - 25,
                TextAlignment.LEFT)

            End If


            If _headerCenter <> "" Then

                AddText(
                canvas,
                _headerCenter,
                pageSize.GetWidth() / 2,
                pageSize.GetTop() - 25,
                TextAlignment.CENTER)

            End If


            If _headerRight <> "" Then

                AddText(
                canvas,
                _headerRight,
                pageSize.GetRight() - 40,
                pageSize.GetTop() - 25,
                TextAlignment.RIGHT)

            End If


            '==========================================================
            ' Fußzeile links
            '==========================================================

            If _footerLeft <> "" Then

                AddText(
                canvas,
                _footerLeft,
                pageSize.GetLeft() + 40,
                pageSize.GetBottom() + 20,
                TextAlignment.LEFT)

            End If


            '==========================================================
            ' Fußzeile Mitte
            '==========================================================

            If _footerCenter <> "" Then

                AddText(
                canvas,
                _footerCenter,
                pageSize.GetWidth() / 2,
                pageSize.GetBottom() + 20,
                TextAlignment.CENTER)

            End If


            '==========================================================
            ' Seite X von Y
            '==========================================================

            Dim pageText As New Paragraph(
            "Seite " & pageNumber.ToString() & " von")

            pageText.SetFont(_font)
            pageText.SetFontSize(8)
            pageText.SetMargin(0)


            Dim x As Single =
            pageSize.GetRight() - 40

            Dim y As Single =
            pageSize.GetBottom() + 20


            ' "Seite X von" rechtsbündig ausgeben
            canvas.ShowTextAligned(
            pageText,
            x - 25,
            y,
            TextAlignment.RIGHT)


            '==========================================================
            ' Platzhalter für Y
            '==========================================================

            pdfCanvas.AddXObjectAt(
            _totalPagesPlaceholder,
            x - 20,
            y - 3)


            canvas.Close()

        End Sub


        '==============================================================
        ' Gesamtseitenzahl in den Platzhalter schreiben
        '==============================================================

        Public Sub WriteTotalPages(
        pdfDoc As PdfDocument)


            Dim canvas As New Canvas(
            _totalPagesPlaceholder,
            pdfDoc)


            Dim totalPages As String =
            pdfDoc.GetNumberOfPages().ToString()


            Dim p As New Paragraph(totalPages)

            p.SetFont(_font)
            p.SetFontSize(8)
            p.SetMargin(0)


            canvas.ShowTextAligned(
            p,
            0,
            3,
            TextAlignment.LEFT)


            canvas.Close()

        End Sub


        '==============================================================
        ' Text ausgeben
        '==============================================================

        Private Sub AddText(
        canvas As Canvas,
        text As String,
        x As Single,
        y As Single,
        alignment As TextAlignment)


            Dim p As New Paragraph(text)

            p.SetFont(_font)
            p.SetFontSize(8)
            p.SetMargin(0)


            canvas.ShowTextAligned(
            p,
            x,
            y,
            alignment)

        End Sub

    End Class