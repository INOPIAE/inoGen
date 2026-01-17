Imports System.Data
Imports System.Text.RegularExpressions
Imports inoGenDLL
Imports iText.IO.Font.Constants
Imports iText.Kernel
Imports iText.Kernel.Font
Imports iText.Kernel.Pdf
Imports iText.Kernel.Pdf.Action
Imports iText.Kernel.Pdf.Annot
Imports iText.Kernel.Pdf.Event
Imports iText.Kernel.Pdf.Navigation
Imports iText.Layout
Imports iText.Layout.Element
Imports iText.Layout.Properties

Public Class MdlPDFVKHReport
    Private Shared cGenDB As New ClsGenDB(My.Settings.DBPath)
    Private Shared cAT As New clsAhnentafelDaten(My.Settings.DBPath)

    Public Shared Sub GenerateReport(dest As String)
        GenerateReport(dest, False, "")
    End Sub

    Public Shared Sub GenerateReport(dest As String, CheckData As Boolean)
        GenerateReport(dest, CheckData, "")
    End Sub

    Public Shared Sub GenerateReport(dest As String, Ort As String)
        GenerateReport(dest, False, Ort)
    End Sub

    Public Shared Sub GenerateReport(dest As String, CheckData As Boolean, Ort As String)
        Using writer As New PdfWriter(dest)
            Using pdfDoc As New PdfDocument(writer)
                Using document As New Document(pdfDoc)

                    ' Kopf-/Fußzeilen aktivieren
                    'pdfDoc.AddEventHandler(PdfDocumentEvent.END_PAGE, New HeaderFooterHandler("Allgemeine Kopfzeile"))
                    '  pdfDoc.AddEventHandler(PdfDocumentEvent.END_PAGE, New ClsPageNumberEventHandler())
                    '' Schriftarten definieren
                    Dim normalFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA)
                    Dim boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD)
                    Dim italicFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_OBLIQUE)

                    Dim linkPattern As String = "\[(.*?)\]\((.*?)\)"

                    ' Outline-Root für Bookmarks
                    Dim rootOutline As PdfOutline = pdfDoc.GetOutlines(CheckData)

                    Dim dt As DataTable = cGenDB.VKH_ReportData()
                    Dim line As String
                    line = String.Format("#{0} {1}", "Kirchenbuch Verkartung aus", dt.Rows(0).Item("BUCH_H"))
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)

                    Dim dtÜ As DataTable = cGenDB.StatisicsVKHeirat

                    With dtÜ.Rows(0)
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Einträge gesamt: " & Format(.Item(0), "#,##0"))
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Bräutigame gesamt: " & Format(.Item(1), "#,##0"))
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Väter des Bräutigams gesamt: " & Format(.Item(2), "#,##0"))
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Mütter des Bräutigams gesamt: " & Format(.Item(3), "#,##0"))
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Bräute gesamt: " & Format(.Item(4), "#,##0"))
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Väter der Braut gesamt: " & Format(.Item(5), "#,##0"))
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Mütter des Braut gesamt: " & Format(.Item(6), "#,##0"))
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Personen gesamt: " & Format(.Item(1) + .Item(2) + .Item(3) + .Item(4) + .Item(5) + .Item(6), "#,##0"))
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Zeugen gesamt: " & Format(.Item(7) + .Item(8) + .Item(9) + .Item(10), "#,##0"))
                    End With


                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Abkürzungen:")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§Bt: Braut")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§Btgm: Bräutigam")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§H: Heimatort der Person")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§M: Mutter der Braut / des Bräutigams")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§V: Vater der Braut / des Bräutigams")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§W: Wohnort der Person")


                    document.Add(New AreaBreak(AreaBreakType.NEXT_PAGE))
                    line = String.Format("#{0}", "Einträge")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)

                    dt = cGenDB.VKH_ReportData()

                    Dim year As Integer = 0
                    Dim counter As Integer = 0

                    For Each dr As DataRow In dt.Rows
                        counter += 1
                        If dr.Item("NR_H").ToString.Substring(0, 4) <> year Then
                            year = dr.Item("NR_H").ToString.Substring(0, 4)
                            line = String.Format("#{0}", year)
                            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)

                        End If
                        line = String.Format("{0} Seite {1}", dr.Item("NR_H"), dr.Item("SEITE_H"))

                        If TestDataValid(dr.Item("OnlineReference")) Then
                            line &= " Onlinequelle: " & dr.Item("OnlineReference")
                        End If
                        OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)

                        line = ""
                        If TestDataValid(dr.Item("HDatum")) Then
                            line = String.Format("Heiratsdatum: {0}", String.Format(dr.Item("HDatum"), "dd.MM.yyyy"))
                        End If
                        If TestDataValid(dr.Item("DimDatum")) Then
                            line &= ", " & dr.Item("DimDatum")
                        End If
                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)


                        line = ""
                        If TestDataValid(dr.Item("FN_BR")) Then
                            line = dr.Item("FN_BR").ToString.ToUpper()
                        End If
                        If TestDataValid(dr.Item("VN_BR")) Then
                            line &= ", " & dr.Item("VN_BR")
                        End If
                        If TestDataValid(dr.Item("GebDatum_BR")) Then
                            line &= ", geb " & dr.Item("GebDatum_BR")
                        End If
                        If TestDataValid(dr.Item("W_BR")) Then
                            line &= ", W: " & dr.Item("W_BR")
                        End If
                        If TestDataValid(dr.Item("H_BR")) Then
                            line &= ", H: " & dr.Item("H_BR")
                        End If
                        If TestDataValid(dr.Item("Z_BR")) Then
                            line &= ", " & dr.Item("Z_BR")
                        End If
                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Btgm: " & line)

                        line = ""
                        If TestDataValid(dr.Item("FN_VBR")) Then
                            line = dr.Item("FN_VBR").ToString.ToUpper()
                        End If
                        If TestDataValid(dr.Item("VN_VBR")) Then
                            line &= ", " & dr.Item("VN_VBR")
                        End If
                        If TestDataValid(dr.Item("Z_VBR")) Then
                            line &= ", " & dr.Item("Z_VBR")
                        End If
                        If TestDataValid(dr.Item("W_EBR")) Then
                            line &= ", " & dr.Item("W_EBR")
                        End If

                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§V: " & line)

                        line = ""
                        If TestDataValid(dr.Item("FN_MBR")) Then
                            line = dr.Item("FN_MBR").ToString.ToUpper()
                        End If

                        If TestDataValid(dr.Item("VN_MBR")) Then
                            line &= ", " & dr.Item("VN_MBR")
                        End If
                        If TestDataValid(dr.Item("Z_MBR")) Then
                            line &= ", " & dr.Item("Z_MBR")
                        End If

                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§M: " & line)

                        AddPageBreakIfNeeded(document, 2)


                        line = ""
                        If TestDataValid(dr.Item("FN_BT")) Then
                            line = dr.Item("FN_BT").ToString.ToUpper()
                        End If

                        If TestDataValid(dr.Item("VN_BT")) Then
                            line &= ", " & dr.Item("VN_BT")
                        End If
                        If TestDataValid(dr.Item("GebDatum_BT")) Then
                            line &= ", geb " & dr.Item("GebDatum_BT")
                        End If
                        If TestDataValid(dr.Item("W_BT")) Then
                            line &= ", W: " & dr.Item("W_BT")
                        End If
                        If TestDataValid(dr.Item("H_BT")) Then
                            line &= ", H: " & dr.Item("H_BT")
                        End If
                        If TestDataValid(dr.Item("Z_BT")) Then
                            line &= ", " & dr.Item("Z_BT")
                        End If

                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "Bt: " & line)

                        line = ""
                        If TestDataValid(dr.Item("FN_VBT")) Then
                            line = dr.Item("FN_VBT").ToString.ToUpper()
                        End If
                        If TestDataValid(dr.Item("VN_VBT")) Then
                            line &= ", " & dr.Item("VN_VBT")
                        End If
                        If TestDataValid(dr.Item("Z_VBT")) Then
                            line &= ", " & dr.Item("Z_VBT")
                        End If
                        If TestDataValid(dr.Item("W_EBT")) Then
                            line &= ", " & dr.Item("W_EBT")
                        End If

                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§V: " & line)

                        line = ""
                        If TestDataValid(dr.Item("FN_MBT")) Then
                            line = dr.Item("FN_MBT").ToString.ToUpper()
                        End If

                        If TestDataValid(dr.Item("VN_MBT")) Then
                            line &= ", " & dr.Item("VN_MBT")
                        End If
                        If TestDataValid(dr.Item("Z_MBT")) Then
                            line &= ", " & dr.Item("Z_MBT")
                        End If

                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§M: " & line)

                        line = ""
                        If TestDataValid(dr.Item("ANM_H")) Then
                            line = "Bemerkung: " & dr.Item("ANM_H")
                        End If


                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)

                        AddPageBreakIfNeeded(document, 2)

                        line = ""
                        If TestDataValid(dr.Item("FN_HZ1")) Then
                            line = dr.Item("FN_HZ1").ToString.ToUpper()
                        End If

                        If TestDataValid(dr.Item("VN_HZ1")) Then
                            line &= ", " & dr.Item("VN_HZ1")
                        End If
                        If TestDataValid(dr.Item("G_HZ1")) Then
                            line &= ", " & dr.Item("G_HZ1")
                        End If
                        If TestDataValid(dr.Item("Z_HZ1")) Then
                            line &= ", " & dr.Item("Z_HZ1")
                        End If
                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§Z: " & line)

                        line = ""
                        If TestDataValid(dr.Item("FN_HZ2")) Then
                            line = dr.Item("FN_HZ2").ToString.ToUpper()
                        End If

                        If TestDataValid(dr.Item("VN_HZ2")) Then
                            line &= ", " & dr.Item("VN_HZ2")
                        End If
                        If TestDataValid(dr.Item("G_HZ2")) Then
                            line &= ", " & dr.Item("G_HZ2")
                        End If
                        If TestDataValid(dr.Item("Z_HZ2")) Then
                            line &= ", " & dr.Item("Z_HZ2")
                        End If
                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§Z: " & line)

                        line = ""
                        If TestDataValid(dr.Item("FN_HZ3")) Then
                            line = dr.Item("FN_HZ3").ToString.ToUpper()
                        End If

                        If TestDataValid(dr.Item("VN_HZ3")) Then
                            line &= ", " & dr.Item("VN_HZ3")
                        End If
                        If TestDataValid(dr.Item("G_HZ3")) Then
                            line &= ", " & dr.Item("G_HZ3")
                        End If
                        If TestDataValid(dr.Item("Z_HZ3")) Then
                            line &= ", " & dr.Item("Z_HZ3")
                        End If
                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§Z: " & line)

                        line = ""
                        If TestDataValid(dr.Item("FN_HZ4")) Then
                            line = dr.Item("FN_HZ4").ToString.ToUpper()
                        End If

                        If TestDataValid(dr.Item("VN_HZ4")) Then
                            line &= ", " & dr.Item("VN_HZ4")
                        End If
                        If TestDataValid(dr.Item("G_HZ4")) Then
                            line &= ", " & dr.Item("G_HZ4")
                        End If
                        If TestDataValid(dr.Item("Z_HZ4")) Then
                            line &= ", " & dr.Item("Z_HZ4")
                        End If
                        If line <> "" Then OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§Z: " & line)

                        AddPageBreakIfNeeded(document, 2)
                        'If counter > 10 Then
                        '    Exit For
                        'End If
                    Next

                    document.Add(New AreaBreak(AreaBreakType.NEXT_PAGE))
                    line = String.Format("#{0}", "Namensverzeichnis")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)

                    dt = cGenDB.VKH_Personen(3)
                    counter = 0
                    Dim OutputObject As String = ""
                    Dim Nr_H As String = ""
                    Dim LineN As String = ""
                    For Each dr As DataRow In dt.Rows
                        counter += 1

                        If dr.Item("Nachname").ToString.Trim <> OutputObject Then
                            If OutputObject <> "" Then
                                OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§" & LineN)
                            End If
                            line = String.Format("{0}", dr.Item("Nachname").ToString)
                            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                            OutputObject = dr.Item("Nachname").ToString.Trim
                            LineN = ""
                            Nr_H = ""
                        End If

                        If dr.Item("NR_H").ToString <> Nr_H Then
                            LineN &= IIf(LineN = "", "", ", ") & dr.Item("NR_H")
                            Nr_H = dr.Item("NR_H").ToString
                        End If

                        'If counter > 100 Then
                        '    Exit For
                        'End If
                    Next
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§" & LineN)

                    document.Add(New AreaBreak(AreaBreakType.NEXT_PAGE))
                    line = String.Format("#{0}", "Ortsverzeichnis")
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)

                    dt = cGenDB.VKH_Orte(Ort)
                    counter = 0
                    OutputObject = ""
                    LineN = ""
                    For Each dr As DataRow In dt.Rows
                        counter += 1
                        If dr.Item("Ort").ToString.Trim <> OutputObject Then
                            If OutputObject <> "" Then
                                OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§" & LineN)
                            End If
                            line = String.Format("{0}", dr.Item("Ort").ToString)
                            OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, line)
                            OutputObject = dr.Item("Ort").ToString.Trim
                            LineN = ""
                            Nr_H = ""
                        End If

                        If dr.Item("NR_H").ToString <> Nr_H Then
                            LineN &= IIf(LineN = "", "", ", ") & dr.Item("NR_H")
                            Nr_H = dr.Item("NR_H").ToString
                        End If

                        'If counter > 100 Then
                        '    Exit For
                        'End If
                    Next
                    OutputLine(pdfDoc, document, normalFont, italicFont, linkPattern, rootOutline, "§" & LineN)
                End Using
            End Using
        End Using
        Console.WriteLine("PDF erstellt: " & dest)
    End Sub

    Private Shared Function TestDataValid(value As Object) As Boolean
        If IsDBNull(value) Then Return False
        If value Is Nothing Then Return False
        If value.ToString().Trim() = "" Then Return False
        Return True
    End Function

    Private Shared Sub OutputLine(pdfDoc As PdfDocument, document As Document, normalFont As PdfFont, italicFont As PdfFont, linkPattern As String, rootOutline As PdfOutline, line As String)
        Dim para As New Paragraph()
        If line.Trim().StartsWith("#") Then

            Dim fSize As Integer = 12
            para.SetFontSize(fSize)

            para.Add(line.Replace("#", ""))
            document.Add(para)


            Dim page = pdfDoc.GetPage(pdfDoc.GetNumberOfPages())
            rootOutline.AddOutline(line.Replace("#", "")).AddDestination(PdfExplicitDestination.CreateFit(page))
        ElseIf line.Trim().StartsWith("§") Then
            para.Add(line.Replace("§", ""))
            para.SetFontSize(7)
            para.SetMarginLeft(10)
            para.SetMarginBottom(1)
            para.SetMarginTop(1)
            document.Add(para)
        Else
            para.Add(line)
            para.SetFontSize(9)
            para.SetMarginBottom(1)
            para.SetMarginTop(3)
            document.Add(para)
        End If
        'document.Add(para)
    End Sub
End Class
