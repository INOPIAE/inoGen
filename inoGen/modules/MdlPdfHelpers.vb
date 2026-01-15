Imports iText.Kernel.Pdf
Imports iText.Layout
Imports iText.Layout.Element
Imports iText.Layout.Properties
Imports iText.Layout.Renderer
Imports iText.Layout.Layout

Public Module MdlPdfHelpers

    ''' <summary>
    ''' Fügt einen Absatz hinzu und prüft vorher, ob noch genügend Platz auf der Seite ist.
    ''' Bei zu wenig Platz wird ein Seitenumbruch eingefügt.
    ''' </summary>
    ''' <param name="document">iText Document</param>
    ''' <param name="minDistanceCm">Minimaler Abstand zum unteren Rand in cm</param>
    Public Sub AddPageBreakIfNeeded(document As Document, minDistanceCm As Single)

        Dim paragraph As New Paragraph("Test")
        ' Renderer für Absatz erzeugen
        Dim pdfDoc = document.GetPdfDocument()
        Dim page As PdfPage = pdfDoc.GetLastPage()
        Dim pageSize = page.GetPageSize()

        ' Abstand vom unteren Rand in Punkt (1 cm ≈ 28,35 pt)
        Dim minDistancePt As Single = minDistanceCm * 28.35F

        ' Y-Position des unteren Rands der letzten Seite
        Dim yBottom As Single = pageSize.GetBottom() + minDistancePt

        ' Aktuelle Position auf der Seite schätzen:
        ' Höhe des Dokuments bisher auf der Seite
        Dim renderer = document.GetRenderer()
        Dim occupiedHeight As Single = 0
        If renderer.GetCurrentArea() IsNot Nothing Then
            occupiedHeight = renderer.GetCurrentArea().GetBBox().GetHeight()
        End If

        Dim yCurrent As Single = pageSize.GetTop() - occupiedHeight

        ' Absatz-Höhe schätzen (FontSize * Zeilenanzahl)
        Dim paragraphHeight As Single = 10 * 1.2F  ' grobe Schätzung für eine Zeile
        'paragraphHeight *= paragraph.GetChildren().Count   ' multipliziere mit Zeilenanzahl

        ' Prüfen, ob Seitenumbruch nötig
        If yCurrent - paragraphHeight < yBottom Then
            document.Add(New AreaBreak(AreaBreakType.NEXT_PAGE))
        End If

        ' Absatz hinzufügen
        'document.Add(paragraph)
    End Sub

End Module
