Imports System.Globalization

Public Class ClsJahrBackgroundConverter
    Implements IValueConverter

    Public Function Convert(
        value As Object,
        targetType As Type,
        parameter As Object,
        culture As CultureInfo
    ) As Object Implements IValueConverter.Convert

        If value Is Nothing OrElse value Is DBNull.Value Then
            Return Brushes.Transparent
        End If

        Dim anzahl As Integer = System.Convert.ToInt32(value)

        Select Case anzahl
            Case 1
                Return Brushes.Yellow

            Case > 1
                Return Brushes.LightGreen

            Case Else
                Return Brushes.Transparent
        End Select

    End Function

    Public Function ConvertBack(
        value As Object,
        targetType As Type,
        parameter As Object,
        culture As CultureInfo
    ) As Object Implements IValueConverter.ConvertBack

        Return Binding.DoNothing

    End Function
End Class
