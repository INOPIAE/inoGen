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


        Dim pdfCanvas As New PdfCanvas(
        page.GetLastContentStream(),
        page.GetResources(),
        pdfDoc)

        Dim canvas As New Canvas(
        pdfCanvas,
        pageSize)


        ' ... Kopfzeile ...


        Dim p As New Paragraph(
        "Seite " & pageNumber.ToString() & " von")

        p.SetFont(_font)
        p.SetFontSize(8)
        p.SetMargin(0)

        canvas.ShowTextAligned(
        p,
        pageSize.GetRight() - 45,
        pageSize.GetBottom() + 20,
        TextAlignment.RIGHT)


        '==========================================================
        ' Placeholder für Gesamtseitenzahl
        '==========================================================

        pdfCanvas.AddXObjectAt(
        _totalPagesPlaceholder,
        pageSize.GetRight() - 40,
        pageSize.GetBottom() + 17)


        canvas.Close()

    End Sub

    '==============================================================
    ' Gesamtseitenzahl in den Platzhalter schreiben
    '==============================================================

    Public Sub WriteTotalPages(pdfDoc As PdfDocument)

        If pdfDoc Is Nothing Then
            Return
        End If

        If _totalPagesPlaceholder Is Nothing Then
            Return
        End If

        Dim p As Integer =
        pdfDoc.GetNumberOfPages()

        If p <= 0 Then
            Return
        End If

        Dim canvas As New Canvas(
        _totalPagesPlaceholder,
        pdfDoc)

        canvas.ShowTextAligned(
        p.ToString(),
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