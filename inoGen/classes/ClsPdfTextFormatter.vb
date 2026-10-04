
Imports System.Text
Imports System.Text.RegularExpressions
Imports iText.IO.Font
Imports iText.Kernel.Colors
Imports iText.Kernel.Font
Imports iText.Kernel.Geom
Imports iText.Kernel.Pdf.Action
Imports iText.Kernel.Pdf.Annot
Imports iText.Layout.Element

Public Class ClsPdfTextFormatter

    Private ReadOnly _normalFont As PdfFont
        Private ReadOnly _italicFont As PdfFont
        Private ReadOnly _boldFont As PdfFont
        Private ReadOnly _symbolFont As PdfFont

        Public Sub New()

            _normalFont =
            PdfFontFactory.CreateFont(
                "C:\Windows\Fonts\segoeui.ttf",
                PdfEncodings.IDENTITY_H,
                PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED)

            _italicFont =
            PdfFontFactory.CreateFont(
                "C:\Windows\Fonts\segoeuii.ttf",
                PdfEncodings.IDENTITY_H,
                PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED)

            _boldFont =
            PdfFontFactory.CreateFont(
                "C:\Windows\Fonts\segoeuib.ttf",
                PdfEncodings.IDENTITY_H,
                PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED)

            _symbolFont =
            PdfFontFactory.CreateFont(
                "C:\Windows\Fonts\seguisym.ttf",
                PdfEncodings.IDENTITY_H,
                PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED)

        End Sub


        '==========================================================
        ' Normaler Paragraph
        '==========================================================

        Public Function CreateParagraph(
        text As String,
        Optional fontSize As Single = 10) As Paragraph

            Dim para As New Paragraph()

            Dim parts = SplitItalicText(text)

            For Each part In parts

                AddText(
                para,
                part.Text,
                part.Italic,
                fontSize)

            Next

            Return para

        End Function


    '==========================================================
    ' Paragraph mit Markdown-Links
    '
    ' [Linktext](URL)
    '==========================================================

    Public Function CreateParagraphWithLinks(
        text As String,
        Optional fontSize As Single = 10, Optional marginTop As Single = 0, Optional marginBottom As Single = 2) As Paragraph

        Dim para As New Paragraph()
        para.SetMarginTop(marginTop)
        para.SetMarginBottom(marginBottom)

        Dim pattern As String =
            "\[([^\]]+)\]\(([^)]+)\)"

        Dim matches As MatchCollection =
            Regex.Matches(text, pattern)

        '------------------------------------------------------
        ' Keine Links
        '------------------------------------------------------

        If matches.Count = 0 Then

            AddTextWithItalic(
                para,
                text,
                fontSize)

            Return para

        End If


        '------------------------------------------------------
        ' Text und Links verarbeiten
        '------------------------------------------------------

        Dim lastIndex As Integer = 0

        For Each m As Match In matches

            '==================================================
            ' Text vor dem Link
            '==================================================

            If m.Index > lastIndex Then

                Dim beforeText As String =
                    text.Substring(
                        lastIndex,
                        m.Index - lastIndex)

                AddTextWithItalic(
                    para,
                    beforeText,
                    fontSize)

            End If


            '==================================================
            ' Link
            '==================================================

            Dim linkText As String =
                m.Groups(1).Value

            Dim url As String =
                m.Groups(2).Value

            AddLink(
                para,
                linkText,
                url,
                fontSize)


            lastIndex =
                m.Index + m.Length

        Next


        '======================================================
        ' Rest nach dem letzten Link
        '======================================================

        If lastIndex < text.Length Then

            Dim afterText As String =
                text.Substring(lastIndex)

            AddTextWithItalic(
                para,
                afterText,
                fontSize)

        End If

        Return para

    End Function


    '==========================================================
    ' Text mit *Kursiv-Markup*
    '==========================================================

    Private Sub AddTextWithItalic(
        para As Paragraph,
        text As String,
        fontSize As Single)

            If String.IsNullOrEmpty(text) Then
                Return
            End If

            Dim parts =
            SplitItalicText(text)

            For Each part In parts

                AddText(
                para,
                part.Text,
                part.Italic,
                fontSize)

            Next

        End Sub


        '==========================================================
        ' Text hinzufügen
        '==========================================================

        Private Sub AddText(
        para As Paragraph,
        text As String,
        italic As Boolean,
        fontSize As Single)

            If String.IsNullOrEmpty(text) Then
                Return
            End If

            Dim buffer As New StringBuilder()

            Dim currentIsSymbol As Boolean = False
            Dim firstCharacter As Boolean = True

            For Each ch As Char In text

                Dim isSymbol As Boolean =
                IsSymbolCharacter(ch)

                If Not firstCharacter AndAlso
               isSymbol <> currentIsSymbol Then

                    AddTextPart(
                    para,
                    buffer.ToString(),
                    italic,
                    currentIsSymbol,
                    fontSize)

                    buffer.Clear()

                End If

                If firstCharacter Then

                    currentIsSymbol = isSymbol
                    firstCharacter = False

                End If

                buffer.Append(ch)

            Next


            If buffer.Length > 0 Then

                AddTextPart(
                para,
                buffer.ToString(),
                italic,
                currentIsSymbol,
                fontSize)

            End If

        End Sub


        '==========================================================
        ' Einzelnen Textabschnitt hinzufügen
        '==========================================================

        Private Sub AddTextPart(
        para As Paragraph,
        text As String,
        italic As Boolean,
        isSymbol As Boolean,
        fontSize As Single)

            If String.IsNullOrEmpty(text) Then
                Return
            End If

            Dim t As New Text(text)

            If isSymbol Then

                t.SetFont(_symbolFont)

            ElseIf italic Then

                t.SetFont(_italicFont)

            Else

                t.SetFont(_normalFont)

            End If

            t.SetFontSize(fontSize)

            para.Add(t)

        End Sub


        '==========================================================
        ' LINK
        '==========================================================

        Private Sub AddLink(
        para As Paragraph,
        linkText As String,
        url As String,
        fontSize As Single)

            '------------------------------------------------------
            ' Annotation
            '------------------------------------------------------

            Dim pdfLinkAnnot As New PdfLinkAnnotation(
            New Rectangle(0, 0, 0, 0))

            pdfLinkAnnot.SetAction(
            PdfAction.CreateURI(url))


            '------------------------------------------------------
            ' Kein sichtbarer Rahmen
            '------------------------------------------------------

            Dim borderArr As New iText.Kernel.Pdf.PdfArray()

            borderArr.Add(
            New iText.Kernel.Pdf.PdfNumber(0))

            borderArr.Add(
            New iText.Kernel.Pdf.PdfNumber(0))

            borderArr.Add(
            New iText.Kernel.Pdf.PdfNumber(0))

            pdfLinkAnnot.SetBorder(borderArr)


            '------------------------------------------------------
            ' Kursivbereiche des Linktextes ermitteln
            '------------------------------------------------------

            Dim parts =
            SplitItalicText(linkText)


            For Each part In parts

                AddLinkPart(
                para,
                part.Text,
                part.Italic,
                pdfLinkAnnot,
                fontSize)

            Next

        End Sub


        '==========================================================
        ' Teil eines Links
        '==========================================================

        Private Sub AddLinkPart(
        para As Paragraph,
        text As String,
        italic As Boolean,
        annotation As PdfLinkAnnotation,
        fontSize As Single)

            If String.IsNullOrEmpty(text) Then
                Return
            End If

            Dim buffer As New StringBuilder()

            Dim currentIsSymbol As Boolean = False
            Dim firstCharacter As Boolean = True

            For Each ch As Char In text

                Dim isSymbol As Boolean =
                IsSymbolCharacter(ch)

                If Not firstCharacter AndAlso
               isSymbol <> currentIsSymbol Then

                    AddLinkTextPart(
                    para,
                    buffer.ToString(),
                    italic,
                    currentIsSymbol,
                    annotation,
                    fontSize)

                    buffer.Clear()

                End If

                If firstCharacter Then

                    currentIsSymbol = isSymbol
                    firstCharacter = False

                End If

                buffer.Append(ch)

            Next


            If buffer.Length > 0 Then

                AddLinkTextPart(
                para,
                buffer.ToString(),
                italic,
                currentIsSymbol,
                annotation,
                fontSize)

            End If

        End Sub


        '==========================================================
        ' Einzelner Linktext-Abschnitt
        '==========================================================

        Private Sub AddLinkTextPart(
        para As Paragraph,
        text As String,
        italic As Boolean,
        isSymbol As Boolean,
        annotation As PdfLinkAnnotation,
        fontSize As Single)

            If String.IsNullOrEmpty(text) Then
                Return
            End If

            Dim linkElem As New Link(
            text,
            annotation)

            '------------------------------------------------------
            ' Font
            '------------------------------------------------------

            If isSymbol Then

                linkElem.SetFont(_symbolFont)

            ElseIf italic Then

                linkElem.SetFont(_italicFont)

            Else

                linkElem.SetFont(_normalFont)

            End If


            linkElem.SetFontSize(fontSize)

            linkElem.SetFontColor(
            ColorConstants.BLACK)

            para.Add(linkElem)

        End Sub


        '==========================================================
        ' Sonderzeichen
        '==========================================================

        Private Function IsSymbolCharacter(
        ch As Char) As Boolean

            Select Case ch

                Case "∗"c,   ' U+2217
                 "⚭"c,   ' U+26AD
                 "♁"c,   ' U+2641
                 "⚮"c,   ' U+26AE
                 "⚰"c,   ' U+26B0
                 "⚬"c    ' U+26AC

                    Return True

                Case Else

                    Return False

            End Select

        End Function


        '==========================================================
        ' *Kursiv* aufteilen
        '==========================================================

        Private Function SplitItalicText(
        text As String) As List(Of TextPart)

            Dim result As New List(Of TextPart)

            If String.IsNullOrEmpty(text) Then
                Return result
            End If

            Dim buffer As New StringBuilder()

            Dim italic As Boolean = False

            For Each ch As Char In text

                If ch = "*"c Then

                    If buffer.Length > 0 Then

                        result.Add(
                        New TextPart(
                            buffer.ToString(),
                            italic))

                        buffer.Clear()

                    End If

                    italic = Not italic

                Else

                    buffer.Append(ch)

                End If

            Next


            If buffer.Length > 0 Then

                result.Add(
                New TextPart(
                    buffer.ToString(),
                    italic))

            End If

            Return result

        End Function


        '==========================================================
        ' Interne Textstruktur
        '==========================================================

        Private Class TextPart

            Public Property Text As String
            Public Property Italic As Boolean

            Public Sub New(
            text As String,
            italic As Boolean)

                Me.Text = text
                Me.Italic = italic

            End Sub

        End Class

    End Class

