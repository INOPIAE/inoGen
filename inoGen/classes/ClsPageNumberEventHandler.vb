Imports iText.IO.Font.Constants
Imports iText.Kernel.Events
Imports iText.Kernel.Font
Imports iText.Kernel.Geom
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Canvas
Imports iText.Kernel.Pdf.Event
Imports iText.Layout
Imports iText.Layout.Element
Imports iText.Layout.Properties

Public Class ClsPageNumberEventHandler
    Inherits AbstractPdfDocumentEventHandler

    ' Diese Methode muss überschrieben werden
    Protected Overrides Sub OnAcceptedEvent(e As AbstractPdfDocumentEvent)
        ' Wir wissen, dass es ein PdfDocumentEvent ist
        Dim pdfEvent As PdfDocumentEvent = CType(e, PdfDocumentEvent)
        Dim pdfDoc As PdfDocument = pdfEvent.GetDocument()
        Dim page As PdfPage = pdfEvent.GetPage()
        Dim pageNumber As Integer = pdfDoc.GetPageNumber(page)
        Dim pageSize As Rectangle = page.GetPageSize()

        ' Canvas für Footer
        Dim canvas As New PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), pdfDoc)

        ' Paragraph für Seitenzahl
        Dim pdfFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)
        Dim p As New Paragraph("Seite " & pageNumber) '.SetFont(pdfFont).SetFontSize(9)

        ' Layout-Objekt zum Zeichnen
        Dim doc As New Document(pdfDoc)
        doc.ShowTextAligned(p, pageSize.GetWidth() / 2, 20, TextAlignment.CENTER)
        doc.Flush()
    End Sub
End Class
