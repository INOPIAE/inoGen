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
Imports iText.Kernel.Pdf.Event
Imports iText.Kernel.Pdf.Navigation
Imports iText.Layout
Imports iText.Layout.Element
Module MdlPdfAncestorReport

    Private cPdfText As New ClsPdfTextFormatter()

    Sub GenerateReport(src As String, dest As String, person As String)
        Dim fontSizeHeader As Integer = 14
        Dim fontSizeNormal As Integer = 10
        Dim fontSizeSource As Integer = 8

        Using writer As New PdfWriter(dest)
            Using pdfDoc As New PdfDocument(writer)
                Using document As New Document(pdfDoc)

                    '==========================================================
                    ' Kopf-/Fußzeile
                    '==========================================================

                    Dim headerFooter As New ClsPdfHeaderFooter(
                        pdfDoc,
                        headerLeft:=$"Ahnenliste für {person}",
                        headerCenter:="",
                        headerRight:=$"Stand: {DateTime.Now:dd.MM.yyyy}",
                        footerLeft:="© inoGen",
                        footerCenter:="")


                    pdfDoc.AddEventHandler(PdfDocumentEvent.END_PAGE, headerFooter)


                    '==========================================================
                    ' Dokument
                    '==========================================================


                    Dim linkPattern As String = "\[(.*?)\]\((.*?)\)"

                    ' Outline-Root für Bookmarks
                    Dim rootOutline As PdfOutline = pdfDoc.GetOutlines(False)

                    For Each line As String In File.ReadAllLines(src)

                        If String.IsNullOrWhiteSpace(line) Then
                            'document.Add(New Paragraph(" ")) ' Leerzeile beibehalten
                            Continue For
                        End If

                        '=== Überschrift mit # ===

                        If line.Trim().StartsWith("#") Then

                            Dim headingText As String =
                                line.TrimStart("#"c, " "c)

                            Dim countHashtag As Integer =
                                line.Count(Function(c) c = "#"c)

                            Dim fSize As Integer = fontSizeHeader

                            If countHashtag > 1 Then
                                fSize =
                                    Math.Max(
                                        10,
                                        fontSizeHeader - (countHashtag - 1) * 2)
                            End If


                            '==========================================================
                            ' Überschrift mit cPdfTextFormatter erzeugen
                            '==========================================================

                            Dim para As Paragraph =
                                cPdfText.CreateParagraphWithLinks(
                                    headingText,
                                    fSize, 8 - (countHashtag - 1), 4)


                            document.Add(para)


                            '==========================================================
                            ' Outline nur für Überschriften der ersten Ebene
                            '==========================================================

                            If countHashtag = 1 Then

                                Dim page =
                                    pdfDoc.GetPage(
                                        pdfDoc.GetNumberOfPages())


                                ' Markdown-Kursiv-Markup aus dem Outline-Titel entfernen
                                Dim outlineTitle As String =
                                    Regex.Replace(
                                        headingText,
                                        "\*(.*?)\*",
                                        "$1")


                                rootOutline.
                                    AddOutline(outlineTitle).
                                    AddDestination(
                                        PdfExplicitDestination.CreateFit(page))

                            End If
                        ElseIf line.Trim().StartsWith("[") Then

                            document.Add(cPdfText.CreateParagraphWithLinks(line, fontSizeSource, 0, 1))

                        Else

                            document.Add(cPdfText.CreateParagraphWithLinks(line, fontSizeNormal, 0, 0))
                        End If
                    Next

                    'Hinweise
                    document.Add(cPdfText.CreateParagraphWithLinks("Hinweise", fontSizeNormal, 8, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("Es werden diese genealogischen Zeichen für die Ereignisse verwendet:", fontSizeSource, 0, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("* - Geburt", fontSizeSource, 0, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("* - Taufe", fontSizeSource, 0, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("† - Tod", fontSizeSource, 0, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("⚰ - Begräbnis", fontSizeSource, 0, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("⚭ - Heirat", fontSizeSource, 0, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("♁⚭ - kirchliche Heirat", fontSizeSource, 0, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("⚬ - Verlobung", fontSizeSource, 0, 0))
                    document.Add(cPdfText.CreateParagraphWithLinks("⚮ - Scheidung", fontSizeSource, 0, 0))

                    document.Add(cPdfText.CreateParagraphWithLinks("Ist hinter dem Personennamen ein Klammereintrag, so ist dies die Familysearch ID-Nummer. Der Link kann nur mit einem kostenfreien Konto auf [https://familysearch.org](https://familysearch.org) geöffnet werden.", fontSizeSource, 4, 0))

                    document.Add(cPdfText.CreateParagraphWithLinks("Für die Links bei den Quellzitaten wird zum Teil ein ggf. kostenpflichtiges Konto benötigt.", fontSizeSource, 4, 0))

                    'Gesamtzahl
                    headerFooter.WriteTotalPages(pdfDoc)
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

    Public Sub TestFont()

        Dim fontPath As String = "C:\Windows\Fonts\segoeui.ttf" ' "C:\Windows\Fonts\calibri.ttc,0"  '

        Console.WriteLine(File.Exists(fontPath))

        'Dim fontBytes As Byte() = File.ReadAllBytes(fontPath)

        'Console.WriteLine($"Fontgröße: {fontBytes.Length} Bytes")

        'Dim fontProgram As FontProgram =
        '    FontProgramFactory.CreateFont(fontBytes)

        Console.WriteLine("FontProgram erfolgreich erzeugt")


        Dim fontProgram As FontProgram = FontProgramFactory.CreateFont(fontPath)

        ' 2. PdfFont-Objekt erstellen (mit Identitäts-Encoding und automatischer Einbettung)
        Dim windowsFont As PdfFont = PdfFontFactory.CreateFont(fontPath, PdfEncodings.IDENTITY_H)
        Dim pdfFont As PdfFont =
            PdfFontFactory.CreateFont(
                fontProgram,
                PdfEncodings.IDENTITY_H,
                PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED)

        Debug.WriteLine(
    GetType(PdfFontFactory).Assembly.Location)

        Dim sttest As String = GetType(FontProgramFactory).Assembly.Location
        Debug.WriteLine(
    GetType(FontProgramFactory).Assembly.Location)

        Debug.WriteLine(
    AppContext.BaseDirectory)

        Dim normalFont As PdfFont =
    PdfFontFactory.CreateFont(
        "C:\Windows\Fonts\segoeui.ttf",
        PdfEncodings.IDENTITY_H,
        True)

        Console.WriteLine("PdfFont erfolgreich erzeugt")

    End Sub

    Public Sub TestSonderzeichen()

        Dim pdfPath As String = "D:\test\TestSonderzeichen.pdf"

        Dim fontPath As String = "C:\Windows\Fonts\segoeui.ttf"

        Dim normalFont As PdfFont =
            PdfFontFactory.CreateFont(
                fontPath,
                PdfEncodings.IDENTITY_H,
                PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED)



        Dim symbolFont As PdfFont =
    PdfFontFactory.CreateFont(
        "C:\Windows\Fonts\seguisym.ttf",
        PdfEncodings.IDENTITY_H,
        PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED)

        Using writer As New PdfWriter(pdfPath)
            Using pdf As New PdfDocument(writer)
                Using document As New Document(pdf)

                    document.Add(
                        New Paragraph("Test der genealogischen Sonderzeichen").
                        SetFont(normalFont).
                        SetFontSize(14))

                    document.Add(
                        New Paragraph("∗ ⚭ ♁ ⚮ † ⚰ ⚬").
                        SetFont(normalFont).
                        SetFontSize(20))

                    document.Add(
                        New Paragraph("Einzeltest: ∗").
                        SetFont(normalFont))

                    document.Add(
                        New Paragraph("Einzeltest: ⚭").
                        SetFont(normalFont))

                    document.Add(
                        New Paragraph("Einzeltest: ♁").
                        SetFont(normalFont))

                    document.Add(
                        New Paragraph("Einzeltest: ⚮").
                        SetFont(normalFont))

                    document.Add(
                        New Paragraph("Einzeltest: †").
                        SetFont(normalFont))

                    document.Add(
                        New Paragraph("Einzeltest: ⚰").
                        SetFont(normalFont))

                    document.Add(
                        New Paragraph("Einzeltest: ⚬").
                        SetFont(normalFont))

                    document.Add(
    New Paragraph("∗ ⚭ ♁ ⚮ † ⚰ ⚬").
    SetFont(symbolFont).
    SetFontSize(20))
                End Using
            End Using
        End Using

    End Sub
End Module

