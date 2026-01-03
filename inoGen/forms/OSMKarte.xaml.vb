Imports System.IO
Imports System.Windows.Forms
Imports System.Windows.Media.Imaging
Imports inoGenDLL
Imports inoGenDLL.clsAhnentafelDaten
Imports iText.IO.Font.Constants
Imports iText.IO.Image
Imports iText.Kernel.Font
Imports iText.Kernel.Pdf
Imports iText.Layout
Imports iText.Layout.Element
Imports Mapsui
Imports Mapsui.Layers
Imports Mapsui.Nts
Imports Mapsui.Projections
Imports Mapsui.Providers
Imports Mapsui.Styles
Imports Mapsui.Tiling
Imports Mapsui.UI.Wpf



Public Class OSMKarte

    Private Persons As List(Of clsAhnentafelDaten.PersonData)
    Private PLocations As List(Of ClsOSMKarte.marker)
    Public Sub New()
        InitializeComponent()


        Dim map = New Mapsui.Map()
        map.Layers.Add(Mapsui.Tiling.OpenStreetMap.CreateTileLayer())
        mapControl.Map = map


        ' Liste mit Markerkoordinaten (lat, lon, label)
        Dim locations = New List(Of Tuple(Of Double, Double, String)) From {
            Tuple.Create(50.7352621, 7.1024635, "Bonn"),
            Tuple.Create(52.520008, 13.404954, "Berlin"),
            Tuple.Create(48.137154, 11.576124, "München")
        }

        ' Marker hinzufügen
        Dim markerFeatures As New List(Of IFeature)
        For Each location In locations
            markerFeatures.Add(CreateMarker(location.Item1, location.Item2, location.Item3))
            'markerFeatures.Add(AddPin(location.Item1, location.Item2, location.Item3))
        Next

        Dim markerLayer = New MemoryLayer("Marker") With {
            .IsMapInfoLayer = True,
            .Features = markerFeatures
        }
        map.Layers.Add(markerLayer)

        ' Nach Home -> auf alle Marker zoomen
        map.Home = Sub(n As Navigator)
                       Dim env = markerLayer.Extent
                       If env IsNot Nothing Then
                           Dim buffered = env.Grow(env.Width * 0.1, env.Height * 0.1)
                           n.ZoomToBox(buffered)
                       End If
                   End Sub

        mapControl.Refresh()

    End Sub

    Public Sub New(GenLocations As List(Of ClsOSMKarte.marker))
        InitializeComponent()

        Dim map = New Mapsui.Map()
        map.Layers.Add(Mapsui.Tiling.OpenStreetMap.CreateTileLayer())
        mapControl.Map = map
        PLocations = GenLocations

        ' Marker hinzufügen
        Dim markerFeatures As New List(Of IFeature)
        For Each location In GenLocations
            markerFeatures.Add(CreateMarker(location.lat, location.lon, location.title))
        Next

        Dim markerLayer = New MemoryLayer("Marker") With {
            .IsMapInfoLayer = True,
            .Features = markerFeatures
        }
        map.Layers.Add(markerLayer)

        ' Nach Home -> auf alle Marker zoomen
        map.Home = Sub(n As Navigator)
                       Dim env = markerLayer.Extent
                       If env IsNot Nothing Then
                           Dim buffered = env.Grow(env.Width * 0.1, env.Height * 0.1)
                           n.ZoomToBox(buffered)
                       End If
                   End Sub

        mapControl.Refresh()

    End Sub

    Public Sub New(GenLocations As List(Of ClsOSMKarte.marker), PPersons As List(Of clsAhnentafelDaten.PersonData))
        InitializeComponent()

        Dim map = New Mapsui.Map()
        map.Layers.Add(Mapsui.Tiling.OpenStreetMap.CreateTileLayer())
        mapControl.Map = map
        Persons = PPersons
        PLocations = GenLocations

        ' Marker hinzufügen
        Dim markerFeatures As New List(Of IFeature)
        For Each location In GenLocations
            markerFeatures.Add(CreateMarker(location.lat, location.lon, location.title))
            'markerFeatures.Add(AddPin(location.Item1, location.Item2, location.Item3))
        Next

        Dim markerLayer = New MemoryLayer("Marker") With {
            .IsMapInfoLayer = True,
            .Features = markerFeatures
        }
        map.Layers.Add(markerLayer)

        ' Nach Home -> auf alle Marker zoomen
        map.Home = Sub(n As Navigator)
                       Dim env = markerLayer.Extent
                       If env IsNot Nothing Then
                           Dim buffered = env.Grow(env.Width * 0.1, env.Height * 0.1)
                           n.ZoomToBox(buffered)
                       End If
                   End Sub

        mapControl.Refresh()

    End Sub

    Private Sub AddMarker(mapControl As MapControl, lat As Double, lon As Double, label As String)
        Dim sm = SphericalMercator.FromLonLat(lon, lat)
        Dim p = New MPoint(sm.x, sm.y)

        ' Feature erstellen
        Dim f = New PointFeature(p)

        ' Marker-Stil (einfarbiger Kreis)
        f.Styles.Add(New SymbolStyle With {
            .SymbolScale = 1.0,
            .Fill = New Brush(Color.Red),
            .Outline = New Pen(Color.White, 2)
        })

        'label-Stil direkt am Feature
        f.Styles.Add(New LabelStyle With {
            .Text = label,
            .BackColor = New Brush(Color.FromArgb(255, 255, 255, 200)),
            .Halo = New Pen(Color.Black, 2)
        })

        ' Layer mit Feature(n)
        Dim markerLayer = New MemoryLayer("Marker") With {
            .IsMapInfoLayer = True,
            .Features = New List(Of IFeature) From {f}
        }

        ' Layer zur Karte hinzufügen
        mapControl.Map.Layers.Add(markerLayer)
    End Sub

    Private Sub btnBild_Click(sender As Object, e As RoutedEventArgs) Handles btnBild.Click
        Dim sfd As New SaveFileDialog With {
            .Filter = "PNG Dateien (*.png)|*.png",
            .Title = "Karte als Bild speichern",
            .FileName = "Karte.png"
        }
        If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then

            SaveMapAsImage(mapControl, sfd.FileName)
        End If

    End Sub

    Private Function CreateMarker(lat As Double, lon As Double, label As String) As IFeature
        ' WGS84 -> WebMercator
        Dim sm = SphericalMercator.FromLonLat(lon, lat)
        Dim p = New MPoint(sm.x, sm.y)

        ' Punkt-Feature
        Dim f = New PointFeature(p)

        ' Symbol (roter Kreis)
        f.Styles.Add(New SymbolStyle With {
            .SymbolScale = 0.9,
            .Fill = New Brush(Color.Red),
            .Outline = New Pen(Color.White, 1)
        })

        ' Label
        'f.Styles.Add(New LabelStyle With {
        '    .Text = label,
        '    .BackColor = New Brush(Color.FromArgb(255, 255, 255, 200)),
        '    .Halo = New Pen(Color.Black, 2)
        '})

        Return f
    End Function

    Private Function AddPin(lat As Double, lon As Double, label As String) As IFeature
        ' WGS84 -> WebMercator
        Dim merc = Projections.SphericalMercator.FromLonLat(lon, lat)
        Dim p = New MPoint(merc.x, merc.y)

        ' Feature
        Dim f = New PointFeature(p)

        ' Pin-Bild laden (Pfad anpassen, PNG mit Transparent empfohlen)
        'Dim bitmapId = BitmapRegistry.Instance.Register("pack://application:,,,/inoGen;component/assets/PinRot.png")
        Dim path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "assets\PinRot.png")
        If File.Exists(path) = False Then
            MessageBox.Show("Pin-Bild nicht gefunden: " & path)
        Else
            MessageBox.Show("Pin-Bild gefunden: " & path)
        End If
        Try
            Dim bitmapId = BitmapRegistry.Instance.Register(path)


            If bitmapId = 0 Then
                MessageBox.Show("Bild konnte nicht geladen werden!")
            End If

            ' SymbolStyle mit Icon
            Dim style As New SymbolStyle With {
                .BitmapId = bitmapId,
                .SymbolScale = 0.5,     ' Größe anpassen
                .SymbolOffset = New Offset(0, 1) ' verschiebt den Ankerpunkt (Pin-Spitze nach unten)
            }
            f.Styles.Add(style)

            ' Label zusätzlich (optional)
            f.Styles.Add(New LabelStyle With {
                .Text = label,
                .BackColor = New Brush(Color.FromArgb(255, 255, 255, 200)),
                .Halo = New Pen(Color.Black, 2),
                .Offset = New Offset(0, -20) ' Text oberhalb des Pins
            })
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        '' Layer hinzufügen
        'Dim markerLayer = New MemoryLayer("Marker") With {
        '    .IsMapInfoLayer = True,
        '    .Features = New List(Of IFeature) From {f}
        '}
        'mapControl.Map.Layers.Add(markerLayer)
        Return f
    End Function

    Public Sub SaveMapAsImage(mapControl As MapControl, filename As String)
        ' MapControl in Bitmap rendern
        Dim width = CInt(mapControl.ActualWidth)
        Dim height = CInt(mapControl.ActualHeight)

        If width = 0 OrElse height = 0 Then
            MessageBox.Show("MapControl hat keine Größe.")
            Return
        End If

        ' RenderTargetBitmap erzeugen
        Dim renderBitmap = New RenderTargetBitmap(width, height, 96, 96, System.Windows.Media.PixelFormats.Pbgra32)
        renderBitmap.Render(mapControl)

        ' Bitmap als PNG speichern
        Dim encoder = New PngBitmapEncoder()
        encoder.Frames.Add(BitmapFrame.Create(renderBitmap))

        Using stream As New FileStream(filename, FileMode.Create)
            encoder.Save(stream)
        End Using

        If MessageBox.Show("Karte als Bild gespeichert: " & filename & vbCrLf & "Daeti öffnen", "Hinweis", MessageBoxButtons.YesNo) = System.Windows.MessageBoxResult.Yes Then
            OpenPdfFile(filename)
        End If
    End Sub

    Public Sub SaveMapAsPdf(mapControl As MapControl, pdfFile As String)
        ' Karte als Bitmap rendern
        Dim w = CInt(mapControl.ActualWidth), h = CInt(mapControl.ActualHeight)
        If w = 0 OrElse h = 0 Then Return

        Dim rtb = New RenderTargetBitmap(w, h, 96, 96, System.Windows.Media.PixelFormats.Pbgra32)
        rtb.Render(mapControl)

        'Dim ms As New MemoryStream()
        'Dim encoder = New PngBitmapEncoder()
        'encoder.Frames.Add(BitmapFrame.Create(rtb))
        'encoder.Save(ms)
        'ms.Position = 0

        Dim converted As New FormatConvertedBitmap(rtb, PixelFormats.Bgr24, Nothing, 0)

        Dim encoder As New PngBitmapEncoder()
        encoder.Frames.Add(BitmapFrame.Create(converted))
        Dim ms As New MemoryStream()
        encoder.Save(ms)
        ms.Position = 0

        Dim pageSize = iText.Kernel.Geom.PageSize.A4.Rotate()

        ' iText 7: PDF erzeugen und PNG einbinden
        Using writer = New PdfWriter(pdfFile)
            Using pdf = New PdfDocument(writer)
                Using doc = New Document(pdf)
                    Dim header As New Paragraph("Lebensorte der Vorfahren von " & Persons(0).Vorname & " " & Persons(0).Nachname)
                    Dim boldFont As PdfFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD)
                    header.SetFontSize(18)
                    header.SetFont(boldFont)
                    header.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER)
                    doc.Add(header)

                    doc.Add(New Paragraph(" ")) ' Leerzeile

                    ' Bild aus Karte
                    Dim imgData = ImageDataFactory.Create(ms.ToArray())
                    Dim img = New iText.Layout.Element.Image(imgData)
                    img.SetAutoScale(True)
                    img.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER)
                    doc.Add(img)

                    doc.Add(New Paragraph(" ")) ' Leerzeile

                    ' Text unterhalb der Karte
                    Dim footer As New Paragraph(String.Format("Von den {0} Ahnen konnten {1} Lebensorte identifizert werden.", Persons.Count, PLocations.Count))
                    footer.SetFontSize(12)
                    footer.SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT)
                    doc.Add(footer)

                End Using
            End Using
        End Using

        MessageBox.Show($"PDF gespeichert unter {pdfFile}")
        If MessageBox.Show("Karte als PDF gespeichert: " & pdfFile & vbCrLf & "Datei öffnen", "Hinweis", MessageBoxButtons.YesNo) = System.Windows.MessageBoxResult.Yes Then
            OpenPdfFile(pdfFile)
        End If
    End Sub

    Private Sub btnPDF_Click(sender As Object, e As RoutedEventArgs) Handles btnPDF.Click
        Dim sfd As New SaveFileDialog With {
            .Filter = "PDF Dateien (*.pdf)|*.pdf",
            .Title = "Karte als PDF speichern",
            .FileName = "Karte.pdf"
        }
        If sfd.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            SaveMapAsPdf(mapControl, sfd.FileName)
        End If
    End Sub

    Public Sub OpenPdfFile(filePath As String)
        Try
            If IO.File.Exists(filePath) Then
                Process.Start(New ProcessStartInfo(filePath) With {
                    .UseShellExecute = True
                })
            Else
                MessageBox.Show("Datei nicht gefunden: " & filePath)
            End If
        Catch ex As Exception
            MessageBox.Show("Fehler beim Öffnen der PDF: " & ex.Message)
        End Try
    End Sub
End Class
