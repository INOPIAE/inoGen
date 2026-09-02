Imports iText.Kernel.Geom
Imports System.IO
Public Class ClsAhnentafelLayout
    Public Property PageWidth As Single
    Public Property PageHeight As Single
    Public Property MarginLeft As Single = 30
    Public Property MarginRight As Single = 30
    Public Property MarginTop As Single = 30
    Public Property MarginBottom As Single = 30
    Public Property TitleBlockHeight As Single = 90
    Public Property BoxSpacing As Single = 15  ' Abstand zwischen Boxen
    Public Property LineCount As Integer = 15
    Public Property RowCount As Integer = 15

    ' Nutzbare Breite
    Public ReadOnly Property UsableWidth As Single
        Get
            Return PageWidth - MarginLeft - MarginRight
        End Get
    End Property

    ' Nutzbare Höhe (ohne Titelblock)
    Public ReadOnly Property UsableHeight As Single
        Get
            Return PageHeight - MarginTop - MarginBottom - TitleBlockHeight
        End Get
    End Property

    ' Abstand zwischen Linien-Mittelpunkten (horizontal)
    Public ReadOnly Property LineSpacing As Single
        Get
            Return (UsableWidth - BoxWidth) / (LineCount - 1)
        End Get
    End Property

    ' Abstand zwischen Zeilen-Mittelpunkten (vertikal)
    Public ReadOnly Property RowSpacing As Single
        Get
            Return (UsableHeight - BoxHeight) / (RowCount - 1)
        End Get
    End Property

    ' Positionen der vertikalen Referenzlinien (Spalten 1-15)
    ' Linie 8 ist horizontal zentriert
    Public ReadOnly Property LinePositions As List(Of Single)
        Get
            Dim positions As New List(Of Single)
            Dim centerX As Single = PageWidth / 2  ' Seitenmitte

            ' Linie 8 (Index 7) liegt in der Mitte
            For i As Integer = 0 To LineCount - 1
                Dim offset As Single = (i - 7) * LineSpacing  ' 7 = Index von Linie 8
                positions.Add(centerX + offset)
            Next

            Return positions
        End Get
    End Property

    ' Positionen der horizontalen Referenzlinien (Zeilen 1-15)
    ' Zeile 8 ist vertikal zentriert im nutzbaren Bereich
    Public ReadOnly Property RowPositions As List(Of Single)
        Get
            Dim positions As New List(Of Single)
            Dim workAreaBottom As Single = MarginBottom
            Dim workAreaTop As Single = PageHeight - MarginTop - TitleBlockHeight
            Dim centerY As Single = workAreaBottom + (UsableHeight / 2)  ' Vertikale Mitte

            ' Zeile 8 (Index 7) liegt in der Mitte
            For i As Integer = 0 To RowCount - 1
                Dim offset As Single = (7 - i) * RowSpacing  ' 7 = Index von Zeile 8, invertiert da Y von unten nach oben geht
                positions.Add(centerY + offset)
            Next

            Return positions
        End Get
    End Property

    ' Breite eines Kästchens (horizontal)
    Public ReadOnly Property BoxWidth As Single
        Get
            'Return LineSpacing - BoxSpacing
            Return (UsableWidth / (LineCount - 1)) - BoxSpacing * 2
        End Get
    End Property

    ' Höhe eines Kästchens (vertikal)
    Public ReadOnly Property BoxHeight As Single
        Get
            'Return RowSpacing - BoxSpacing
            Return (UsableHeight / (RowCount - 1)) - BoxSpacing * 2
        End Get
    End Property

    ' X-Position einer Box (mittig auf Linie)
    Public Function GetBoxX(columnIndex As Integer) As Single
        Dim lineX As Single = LinePositions(columnIndex)
        Return lineX - (BoxWidth / 2)
    End Function

    ' Y-Position einer Box (mittig auf Zeile)
    Public Function GetBoxY(rowIndex As Integer) As Single
        Dim rowY As Single = RowPositions(rowIndex)
        Return rowY - (BoxHeight / 2)
    End Function

    ' Vollständige Box-Informationen für Position (Zeile, Spalte)
    Public Function GetBox(rowIndex As Integer, columnIndex As Integer) As Rectangle
        Dim x As Single = GetBoxX(columnIndex)
        Dim y As Single = GetBoxY(rowIndex)
        Return New Rectangle(x, y, BoxWidth, BoxHeight)
    End Function

    ' Titelblock-Bereich
    Public ReadOnly Property TitleBlockArea As Rectangle
        Get
            Return New Rectangle(MarginLeft,
                               PageHeight - MarginTop - TitleBlockHeight,
                               UsableWidth,
                               TitleBlockHeight)
        End Get
    End Property

    ' Konstruktor 1: Einfach - nur Seitengröße mit Standardwerten
    Public Sub New(pageSize As (width As Single, height As Single), Landscape As Boolean)
        AdjustLandscape(pageSize, Landscape)
    End Sub

    Private Sub AdjustLandscape(pageSize As (width As Single, height As Single), Landscape As Boolean)
        If Landscape Then
            PageWidth = pageSize.height
            PageHeight = pageSize.width
        Else
            PageWidth = pageSize.width
            PageHeight = pageSize.height
        End If
    End Sub

    ' Konstruktor 2: Erweitert - mit anpassbaren Rändern und Titelblock
    Public Sub New(pageSize As (width As Single, height As Single),
                   margin As Single,
                   titleBlockHeight As Single, Landscape As Boolean)
        AdjustLandscape(pageSize, Landscape)
        MarginLeft = margin
        MarginRight = margin
        MarginTop = margin
        MarginBottom = margin
        Me.TitleBlockHeight = titleBlockHeight
    End Sub

    ' Konstruktor 3: Vollständig - individuelle Ränder
    Public Sub New(pageSize As (width As Single, height As Single),
                   marginLeft As Single,
                   marginRight As Single,
                   marginTop As Single,
                   marginBottom As Single,
                   titleBlockHeight As Single, Landscape As Boolean)
        AdjustLandscape(pageSize, Landscape)
        Me.MarginLeft = marginLeft
        Me.MarginRight = marginRight
        Me.MarginTop = marginTop
        Me.MarginBottom = marginBottom
        Me.TitleBlockHeight = titleBlockHeight
    End Sub

    ' Ausgabe aller Berechnungen
    Public Sub PrintLayout()
        Dim fileName As String = "D:\test\testLayout.log"
        Try
            Kill(fileName)
        Catch ex As Exception

        End Try
        File.AppendAllText(fileName, "=== SEITENLAYOUT ===" & Environment.NewLine)
        File.AppendAllText(fileName, $"Seitenbreite: {PageWidth} pt ({PageWidth / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, $"Seitenhöhe: {PageHeight} pt ({PageHeight / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, $"Seitenränder: Links={MarginLeft} pt, Rechts={MarginRight} pt, Oben={MarginTop} pt, Unten={MarginBottom} pt" & Environment.NewLine)
        File.AppendAllText(fileName, $"Titelblock-Höhe: {TitleBlockHeight} pt ({TitleBlockHeight / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, Environment.NewLine)

        File.AppendAllText(fileName, "=== NUTZBARE BEREICHE ===" & Environment.NewLine)
        File.AppendAllText(fileName, $"Nutzbare Breite: {UsableWidth} pt ({UsableWidth / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, $"Nutzbare Höhe (ohne Titelblock): {UsableHeight} pt ({UsableHeight / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, Environment.NewLine)

        File.AppendAllText(fileName, "=== ABSTÄNDE ===" & Environment.NewLine)
        File.AppendAllText(fileName, $"Horizontaler Linienabstand: {LineSpacing:F2} pt ({LineSpacing / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, $"Vertikaler Zeilenabstand: {RowSpacing:F2} pt ({RowSpacing / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, $"Kästchenbreite: {BoxWidth:F2} pt ({BoxWidth / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, $"Kästchenhöhe: {BoxHeight:F2} pt ({BoxHeight / 28.35F:F2} cm)" & Environment.NewLine)
        File.AppendAllText(fileName, $"Abstand zwischen Kästchen: {BoxSpacing} pt" & Environment.NewLine)
        File.AppendAllText(fileName, Environment.NewLine)

        File.AppendAllText(fileName, "=== LINIENPOSITIONEN (Spalten) ===" & Environment.NewLine)
        File.AppendAllText(fileName, $"Seitenmitte: {PageWidth / 2} pt" & Environment.NewLine)
        For i As Integer = 0 To LinePositions.Count - 1
            Dim marker As String = If(i = 7, " <-- ZENTRIERT", "" & Environment.NewLine)
            File.AppendAllText(fileName, $"  Spalte {i + 1,2}: {LinePositions(i),7:F2} pt ({LinePositions(i) / 28.35F,5:F2} cm){marker}" & Environment.NewLine)
        Next
        File.AppendAllText(fileName, Environment.NewLine)

        File.AppendAllText(fileName, "=== ZEILENPOSITIONEN (Zeilen) ===" & Environment.NewLine)
        Dim workAreaCenter As Single = MarginBottom + (UsableHeight / 2)
        File.AppendAllText(fileName, $"Arbeitsbereich-Mitte: {workAreaCenter:F2} pt" & Environment.NewLine)
        For i As Integer = 0 To RowPositions.Count - 1
            Dim marker As String = If(i = 7, " <-- ZENTRIERT", "" & Environment.NewLine)
            File.AppendAllText(fileName, $"  Zeile {i + 1,2}: {RowPositions(i),7:F2} pt ({RowPositions(i) / 28.35F,5:F2} cm){marker}" & Environment.NewLine)
        Next
    End Sub

End Class
'' Test-Instanz der Klasse
'Dim layout As New ClsAhnentafelLayout(PdfPageSizes.A4)

'' Erste Spalte prüfen
'Dim box1 = layout.GetBox(0, 0)
'Console.WriteLine($"Erste Box X: {box1.GetX()} (erwartet: {layout.MarginLeft})")

'' Letzte Spalte prüfen
'Dim box15 = layout.GetBox(0, 14)
'Dim expectedRight = layout.PageWidth - layout.MarginRight
'Console.WriteLine($"Letzte Box rechts: {box15.GetX() + box15.GetWidth()} (erwartet: {expectedRight})")

'' Erste Zeile prüfen (oben)
'Dim boxTop = layout.GetBox(0, 0)
'Dim expectedTop = layout.PageHeight - layout.MarginTop - layout.TitleBlockHeight
'Console.WriteLine($"Erste Box oben: {boxTop.GetY() + boxTop.GetHeight()} (max: {expectedTop})")

'' Letzte Zeile prüfen (unten)
'Dim boxBottom = layout.GetBox(14, 0)
'Console.WriteLine($"Letzte Box unten: {boxBottom.GetY()} (erwartet: {layout.MarginBottom})")
