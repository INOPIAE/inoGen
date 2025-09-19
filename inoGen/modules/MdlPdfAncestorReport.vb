Imports System.IO
Imports System.Text.RegularExpressions
Imports iText.IO.Font
Imports iText.IO.Font.Constants
Imports iText.Kernel.Events
Imports iText.Kernel.Font
Imports iText.Kernel.Geom
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Action
Imports iText.Kernel.Pdf.Annot
Imports iText.Kernel.Pdf.Canvas
Imports iText.Kernel.Pdf.Navigation
Imports iText.Layout
Imports iText.Layout.Element
Module MdlPdfAncestorReport

    Dim projectRoot As String = Directory.GetParent(AppDomain.CurrentDomain.BaseDirectory).Parent.Parent.FullName

    ' Fonts-Ordner



    '' EventHandler für Kopf- und Fußzeilen
    'Public Class HeaderFooterHandler
    '    Implements IEventHandler

    '    Private ReadOnly _headerText As String
    '    Private ReadOnly _font As iText.Kernel.Font.PdfFont

    '    Public Sub New(headerText As String)
    '        _headerText = headerText
    '        _font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)
    '    End Sub

    '    Public Sub HandleEvent([event] As [Event]) Implements IEventHandler.HandleEvent
    '        Dim docEvent As PdfDocumentEvent = CType([event], PdfDocumentEvent)
    '        Dim pdfDoc As PdfDocument = docEvent.GetDocument()
    '        Dim page As PdfPage = docEvent.GetPage()
    '        Dim pageNumber As Integer = pdfDoc.GetPageNumber(page)
    '        Dim pageSize As Rectangle = page.GetPageSize()

    '        Dim canvas As New PdfCanvas(page.NewContentStreamAfter(), page.GetResources(), pdfDoc)

    '        ' Kopfzeile
    '        canvas.BeginText() _
    '              .SetFontAndSize(_font, 10) _
    '              .MoveText(pageSize.GetWidth() / 2 - 30, pageSize.GetTop() - 20) _
    '              .ShowText(_headerText) _
    '              .EndText()

    '        ' Fußzeile mit Seitennummer
    '        canvas.BeginText() _
    '              .SetFontAndSize(_font, 10) _
    '              .MoveText(pageSize.GetWidth() / 2 - 15, pageSize.GetBottom() + 20) _
    '              .ShowText("Seite " & pageNumber) _
    '              .EndText()

    '        canvas.Release()
    '    End Sub
    'End Class

    Sub GenerateReport(src As String, dest As String)
        'Dim src As String = "input.txt"
        'Dim dest As String = "output.pdf"
        'PdfFonts.InitFonts()

        Using writer As New PdfWriter(dest)
            Using pdfDoc As New PdfDocument(writer)
                Using document As New Document(pdfDoc)

                    ' Kopf-/Fußzeilen aktivieren
                    'pdfDoc.AddEventHandler(PdfDocumentEvent.END_PAGE, New HeaderFooterHandler("Allgemeine Kopfzeile"))

                    '' Schriftarten definieren
                    Dim normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)
                    Dim boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD)
                    Dim italicFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_OBLIQUE)
                    'Dim fontDir As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets", "liberation-fonts-ttf-2.1.5")

                    'Dim normalFont As PdfFont = PdfFontFactory.CreateFont(IO.Path.Combine(fontDir, "LiberationSans-Regular.ttf"), PdfEncodings.IDENTITY_H, True)
                    'Dim boldFont As PdfFont = PdfFontFactory.CreateFont(IO.Path.Combine(fontDir, "LiberationSans-Bold.ttf"), PdfEncodings.IDENTITY_H, True)
                    'Dim italicFont As PdfFont = PdfFontFactory.CreateFont(IO.Path.Combine(fontDir, "LiberationSans-Italic.ttf"), PdfEncodings.IDENTITY_H, True)

                    Dim linkPattern As String = "\[(.*?)\]\((.*?)\)"

                    ' Outline-Root für Bookmarks
                    Dim rootOutline As PdfOutline = pdfDoc.GetOutlines(False)

                    For Each line As String In File.ReadAllLines(src)

                        If String.IsNullOrWhiteSpace(line) Then
                            'document.Add(New Paragraph(" ")) ' Leerzeile beibehalten
                            Continue For
                        End If

                        ' === Überschrift mit # ===
                        If line.Trim().StartsWith("#") Then
                            Dim headingText As String = line.TrimStart("#"c, " "c)
                            Dim countHashtag As Integer = line.Count(Function(c) c = "#"c)
                            Dim fSize As Integer = 14
                            If countHashtag > 1 Then fSize = Math.Max(10, 14 - (countHashtag - 1) * 2)

                            Dim matches = Regex.Matches(headingText, linkPattern)

                            Dim beforeText As String = ""

                            Dim para As New Paragraph()
                            para.SetFontSize(fSize)
                            If matches.Count = 0 Then
                                ' Kein Link → nur Kursiv-Markup verarbeiten
                                AddItalicAndNormalParts(para, headingText, normalFont, italicFont)
                                beforeText = headingText
                            Else
                                Dim lastIndex As Integer = 0
                                For Each m As Match In matches
                                    ' Text vor Link (mit Kursiv-Erkennung)
                                    If m.Index > lastIndex Then
                                        beforeText = headingText.Substring(lastIndex, m.Index - lastIndex)
                                        AddItalicAndNormalParts(para, beforeText, normalFont, italicFont)
                                    End If

                                    ' Linktext evtl. mit Kursiv-Markup
                                    Dim linkText As String = m.Groups(1).Value
                                    Dim url As String = m.Groups(2).Value

                                    Dim linkParts = linkText.Split("*"c)
                                    For i As Integer = 0 To linkParts.Length - 1
                                        If linkParts(i).Length > 0 Then
                                            Dim pdfLinkAnnot As New PdfLinkAnnotation(New iText.Kernel.Geom.Rectangle(0, 0, 0, 0))
                                            pdfLinkAnnot.SetAction(PdfAction.CreateURI(url))

                                            ' Border-Array [0 0 0] entfernt den sichtbaren Rahmen der Annotation
                                            Dim borderArr As New iText.Kernel.Pdf.PdfArray()
                                            borderArr.Add(New iText.Kernel.Pdf.PdfNumber(0))
                                            borderArr.Add(New iText.Kernel.Pdf.PdfNumber(0))
                                            borderArr.Add(New iText.Kernel.Pdf.PdfNumber(0))
                                            pdfLinkAnnot.SetBorder(borderArr)

                                            ' Erzeuge das Link-Element mit der Annotation
                                            Dim linkElem As New Link(linkParts(i), pdfLinkAnnot)

                                            ' Schriftart (kursiv oder normal)
                                            If i Mod 2 = 1 Then
                                                linkElem.SetFont(italicFont)
                                            Else
                                                linkElem.SetFont(normalFont)
                                            End If

                                            ' Einheitliche Größe (optional, oder Paragraph hat die Größe)
                                            linkElem.SetFontSize(8)

                                            ' Farbe auf schwarz setzen, damit kein blau/unterstrichenes Aussehen
                                            linkElem.SetFontColor(iText.Kernel.Colors.ColorConstants.BLACK)

                                            para.Add(linkElem)

                                        End If
                                    Next

                                    lastIndex = m.Index + m.Length + 2
                                Next

                                ' Rest nach letztem Link
                                If lastIndex < line.Length Then
                                    Dim afterText As String = line.Substring(lastIndex)
                                    AddItalicAndNormalParts(para, afterText, normalFont, italicFont)
                                End If
                            End If

                            document.Add(para)

                            If countHashtag = 1 Then
                                Dim page = pdfDoc.GetPage(pdfDoc.GetNumberOfPages())
                                Dim outlineTitle As String = Regex.Replace(beforeText, "\*(.*?)\*", "*$1")
                                rootOutline.AddOutline(outlineTitle).AddDestination(PdfExplicitDestination.CreateFit(page))
                            End If
                        Else
                            Dim para As New Paragraph()

                            Dim parts = line.Split("*"c)
                            For i As Integer = 0 To parts.Length - 1
                                Dim txt As New Text(parts(i))
                                If i Mod 2 = 1 Then
                                    txt = txt.SetFont(italicFont) ' Kursiv
                                Else

                                    txt = txt.SetFont(normalFont) ' Normal
                                End If
                                para.Add(txt)
                            Next

                            para.SetFontSize(10)
                            document.Add(para)
                        End If
                    Next
                End Using
            End Using
        End Using
        Console.WriteLine("PDF erstellt: " & dest)
    End Sub


    Private Sub AddItalicAndNormalParts(para As Paragraph, text As String, normalFont As iText.Kernel.Font.PdfFont, italicFont As iText.Kernel.Font.PdfFont)
        Dim parts = text.Split("*"c)
        For i As Integer = 0 To parts.Length - 1
            Dim txt As New Text(parts(i))
            If i Mod 2 = 1 Then
                txt = txt.SetFont(italicFont)
            Else
                txt = txt.SetFont(normalFont)
            End If
            para.Add(txt)
        Next
    End Sub
End Module

