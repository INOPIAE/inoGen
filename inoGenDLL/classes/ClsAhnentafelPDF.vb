Public Class ClsAhnentafelPDF
    Public Enum DinFormat
        A0
        A1
        A2
        A3
        A4
    End Enum

    Public Class PdfPageSizes
        Public Shared ReadOnly sizes As New Dictionary(Of DinFormat, (width As Single, height As Single)) From {
        {DinFormat.A0, (2384, 3370)},
        {DinFormat.A1, (1684, 2384)},
        {DinFormat.A2, (1191, 1684)},
        {DinFormat.A3, (842, 1191)},
        {DinFormat.A4, (595, 842)}
    }

        Public Shared Function GetSize(format As DinFormat) As (width As Single, height As Single)
            Return sizes(format)
        End Function

        Public Shared Function GetLandscape(format As DinFormat) As (width As Single, height As Single)
            Dim size = sizes(format)
            Return (size.height, size.width)
        End Function

        ' In cm
        Public Shared Function GetSizeInCm(format As DinFormat) As (width As Single, height As Single)
            Dim size = sizes(format)
            Return (size.width / 28.35F, size.height / 28.35F)
        End Function
    End Class

    '' Verwendung:
    'Dim pageSize = PdfPageSizes.GetSize(DinFormat.A4)
    'Dim landscape = PdfPageSizes.GetLandscape(DinFormat.A3)
    'Dim sizeInCm = PdfPageSizes.GetSizeInCm(DinFormat.A4) ' (21cm, 29.7cm)


    Public Class Koordinaten
        Public Property KZeile As Integer
        Public Property KSpalte As Integer

        Public Sub New(z1 As Integer, z2 As Integer)
            KZeile = z1
            KSpalte = z2
        End Sub
    End Class


    Public Function CreateAhnentafelDictionaryGen7() As Dictionary(Of Integer, Koordinaten)


        Dim dict As New Dictionary(Of Integer, Koordinaten)
        ' Definiere alle Positionen manuell (da Muster sehr komplex ist)
        Dim positions As Integer(,) = {
            {1, 1, 64}, {1, 3, 66}, {1, 5, 72}, {1, 7, 74}, {1, 9, 96}, {1, 11, 98}, {1, 13, 104}, {1, 15, 106},
            {2, 1, 32}, {2, 2, 16}, {2, 3, 33}, {2, 5, 36}, {2, 6, 18}, {2, 7, 37}, {2, 9, 48}, {2, 10, 24}, {2, 11, 49}, {2, 13, 52}, {2, 14, 26}, {2, 15, 53},
            {3, 1, 65}, {3, 3, 67}, {3, 5, 73}, {3, 7, 75}, {3, 9, 97}, {3, 11, 99}, {3, 13, 105}, {3, 15, 107},
            {4, 2, 8}, {4, 4, 4}, {4, 6, 9}, {4, 10, 12}, {4, 12, 6}, {4, 14, 13},
            {5, 1, 68}, {5, 3, 70}, {5, 5, 76}, {5, 7, 78}, {5, 9, 100}, {5, 11, 102}, {5, 13, 108}, {5, 15, 110},
            {6, 1, 34}, {6, 2, 17}, {6, 3, 35}, {6, 5, 38}, {6, 6, 19}, {6, 7, 39}, {6, 9, 50}, {6, 10, 25}, {6, 11, 51}, {6, 13, 54}, {6, 14, 27}, {6, 15, 55},
            {7, 1, 69}, {7, 3, 71}, {7, 5, 77}, {7, 7, 79}, {7, 9, 101}, {7, 11, 103}, {7, 13, 109}, {7, 15, 111},
            {8, 4, 2}, {8, 8, 1}, {8, 12, 3},
            {9, 1, 80}, {9, 3, 82}, {9, 5, 88}, {9, 7, 90}, {9, 9, 112}, {9, 11, 114}, {9, 13, 120}, {9, 15, 122},
            {10, 1, 40}, {10, 2, 20}, {10, 3, 41}, {10, 5, 44}, {10, 6, 22}, {10, 7, 45}, {10, 9, 56}, {10, 10, 28}, {10, 11, 57}, {10, 13, 60}, {10, 14, 30}, {10, 15, 61},
            {11, 1, 81}, {11, 3, 83}, {11, 5, 89}, {11, 7, 91}, {11, 9, 113}, {11, 11, 115}, {11, 13, 121}, {11, 15, 123},
            {12, 2, 10}, {12, 4, 5}, {12, 6, 11}, {12, 10, 14}, {12, 12, 7}, {12, 14, 15},
            {13, 1, 84}, {13, 3, 86}, {13, 5, 92}, {13, 7, 94}, {13, 9, 116}, {13, 11, 118}, {13, 13, 124}, {13, 15, 126},
            {14, 1, 42}, {14, 2, 21}, {14, 3, 43}, {14, 5, 46}, {14, 6, 23}, {14, 7, 47}, {14, 9, 58}, {14, 10, 29}, {14, 11, 59}, {14, 13, 62}, {14, 14, 31}, {14, 15, 63},
            {15, 1, 85}, {15, 3, 87}, {15, 5, 93}, {15, 7, 95}, {15, 9, 117}, {15, 11, 119}, {15, 13, 125}, {15, 15, 127}
        }

        ' Dictionary befüllen
        For i As Integer = 0 To positions.GetLength(0) - 1
            Dim row As Integer = positions(i, 0)
            Dim col As Integer = positions(i, 1)
            Dim nummer As Integer = positions(i, 2)
            Try
                dict.Add(nummer, New Koordinaten(row, col))
            Catch ex As Exception
            End Try
        Next

        Return dict
    End Function
End Class
