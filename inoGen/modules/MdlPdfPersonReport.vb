Imports System.IO
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports inoGenDLL
Imports inoGenDLL.clsAhnentafelDaten
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

Public Class MdlPdfPersonReport

    Private Shared cGenDB As New ClsGenDB(My.Settings.DBPath)
    Private Shared cAT As New clsAhnentafelDaten(My.Settings.DBPath)

    Public Shared Sub GenerateReport(PID As Integer, dest As String)


        Using writer As New PdfWriter(dest)
            Using pdfDoc As New PdfDocument(writer)
                Using document As New Document(pdfDoc)

                    ' Kopf-/Fußzeilen aktivieren
                    'pdfDoc.AddEventHandler(PdfDocumentEvent.END_PAGE, New HeaderFooterHandler("Allgemeine Kopfzeile"))

                    '' Schriftarten definieren
                    Dim normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)
                    Dim boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD)
                    Dim italicFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_OBLIQUE)

                    Dim linkPattern As String = "\[(.*?)\]\((.*?)\)"

                    ' Outline-Root für Bookmarks
                    Dim rootOutline As PdfOutline = pdfDoc.GetOutlines(False)

                    Dim pd = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = PID}
                    cGenDB.FillPersonData(pd)

                    Dim line As String

                    line = String.Format("# {0} {1}", pd.Vorname, pd.Nachname)


                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                    AusgabePersDetails(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, pd)

                    Dim EDL As New List(Of clsAhnentafelDaten.EventData) 
                    EDL = cGenDB.GetPersonenAdditionalData(PID)
                    If EDL.Count > 0 Then
                        AusgabePersEventDetails(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, EDL)
                    End If

                    Dim FL As New List(Of clsAhnentafelDaten.FamilyData)
                    FL = cGenDB.GetFamilies(PID)
                    Dim fc As Integer = 0
                    If FL.Count > 0 Then
                        For Each f As clsAhnentafelDaten.FamilyData In FL
                            fc += 1
                            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, String.Format("# {0}. Ehe", cAT.ToRoman(fc)))
                            Dim pp As inoGenDLL.clsAhnentafelDaten.PersonData =
                                If(f.VID = PID AndAlso f.MID > 0,
                                   New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = f.MID},
                                   If(f.VID > 0,
                                      New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = f.VID},
                                      Nothing))

                            If pp.ID > 0 Then
                                cGenDB.FillPersonData(pp)
                                pp.EID = f.ID
                                cGenDB.FillFamilieDaten(pp)

                                AusgabeFamilyDetails(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, pp)

                                line = $"# {pp.Vorname} {pp.Nachname}"
                                OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                                AusgabePersDetails(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, pp)


                                'writer.WriteLine("## Kinder")
                                cAT.AddChildren(pp.EID)

                                If cAT.Kinder.Count > 0 Then
                                    line = $"## Kinder"
                                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                                End If
                                For Each k In cAT.Kinder.OrderBy(Function(x) x.Pos)
                                    line = $"### {k.Vorname} {k.Nachname}"
                                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                                    AusgabePersDetails(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, k)
                                Next
                            End If
                        Next

                    End If

                    If pd.M > 0 Or pd.V > 0 Then
                        line = $"# Eltern"
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                        If pd.V > 0 Then
                            Dim pv = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = pd.V}
                            cGenDB.FillPersonData(pv)

                            line = String.Format("# {0} {1}", pv.Vorname, pv.Nachname)

                            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                            AusgabePersDetails(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, pv)
                        End If
                        If pd.M > 0 Then
                            Dim pm = New inoGenDLL.clsAhnentafelDaten.PersonData With {.ID = pd.M}
                            cGenDB.FillPersonData(pm)

                            line = String.Format("# {0} {1}", pm.Vorname, pm.Nachname)

                            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                            AusgabePersDetails(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, pm)
                        End If
                    End If

                End Using
            End Using
        End Using
        Console.WriteLine("PDF erstellt: " & dest)
    End Sub

    Private Shared Sub OutputLine(pdfDoc As PdfDocument, document As Document, normalFont As PdfFont, italicFont As PdfFont, linkPattern As String, rootOutline As PdfOutline, line As String)
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
    End Sub

    Private Shared Sub AddItalicAndNormalParts(para As Paragraph, text As String, normalFont As iText.Kernel.Font.PdfFont, italicFont As iText.Kernel.Font.PdfFont)
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

    Public Shared Sub AusgabePersDetails(pdfDoc As PdfDocument, document As Document, normalFont As PdfFont, italicFont As PdfFont, linkPattern As String, rootOutline As PdfOutline, p As PersonData)
        Dim line As String
        If IsNothing(p.Geburtsdatum) = False Or IsNothing(p.Geburtsort) = False Then
            line = "∗ " & If(IsNothing(p.Geburtsdatum), "    ", p.Geburtsdatum) & " " & If(IsNothing(p.Geburtsort), "", p.Geburtsort)  '★
            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
        End If
        If IsNothing(p.Taufdatum) = False Or IsNothing(p.Taufort) = False Then
            line = "~ " & If(IsNothing(p.Taufdatum), "    ", p.Taufdatum) & " " & If(IsNothing(p.Taufort), "", p.Taufort)
            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
        End If
        If IsNothing(p.Sterbedatum) = False Or IsNothing(p.Sterbeort) = False Then
            line = "† " & If(IsNothing(p.Sterbedatum), "    ", p.Sterbedatum) & " " & If(IsNothing(p.Sterbeort), "", p.Sterbeort)
            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
        End If
        If IsNothing(p.Begräbnisdatum) = False Or IsNothing(p.Begräbnisort) = False Then
            line = "⚰ " & If(IsNothing(p.Begräbnisdatum), "    ", p.Begräbnisdatum) & " " & If(IsNothing(p.Begräbnisort), "", p.Begräbnisort)
            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
        End If
    End Sub

    Public Shared Sub AusgabePersEventDetails(pdfDoc As PdfDocument, document As Document, normalFont As PdfFont, italicFont As PdfFont, linkPattern As String, rootOutline As PdfOutline, EDL As List(Of clsAhnentafelDaten.EventData))
        Dim line As String = ""
        Dim header As String = ""
        For Each ed As EventData In EDL
            If ed.Eventname <> header Then
                header = ed.Eventname
                line = "## " & ed.Eventname
                OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
            End If
            If IsNothing(ed.EventDate) = False Or IsNothing(ed.EventLocation) = False Then
                line = If(IsNothing(ed.EventTopic), "", ed.EventTopic & " ") & If(IsNothing(ed.EventDate), "", ed.EventDate) & " " & If(IsNothing(ed.EventLocation), "", ed.EventLocation)
                OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
            End If
        Next
    End Sub


    Public Shared Sub AusgabeFamilyDetails(pdfDoc As PdfDocument, document As Document, normalFont As PdfFont, italicFont As PdfFont, linkPattern As String, rootOutline As PdfOutline, p As PersonData)
        Dim line As String
        If IsNothing(p.Verlobungsdatum) = False Or IsNothing(p.Verlobungsort) = False Then
            line = "⚬ " & If(IsNothing(p.Verlobungsdatum), "    ", p.Verlobungsdatum) & " " & If(IsNothing(p.Verlobungsort), "", p.Verlobungsort)
            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
        End If
        If IsNothing(p.Heiratdatum) = False Or IsNothing(p.Heiratort) = False Then
            line = "⚭ " & If(IsNothing(p.Heiratdatum), "    ", p.Heiratdatum) & " " & If(IsNothing(p.KHeiratort), "", p.KHeiratort)
            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
        End If
        If IsNothing(p.KHeiratdatum) = False Or IsNothing(p.KHeiratort) = False Then
            line = "♁⚭ " & If(IsNothing(p.KHeiratdatum), "    ", p.KHeiratdatum) & " " & If(IsNothing(p.KHeiratort), "", p.KHeiratort)
            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
        End If
        If IsNothing(p.Scheidungsdatum) = False Or IsNothing(p.Scheidungsort) = False Then
            line = "⚮ " & If(IsNothing(p.Scheidungsdatum), "    ", p.Scheidungsdatum) & " " & If(IsNothing(p.Scheidungsort), "", p.Scheidungsort)
            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
        End If
    End Sub


End Class
