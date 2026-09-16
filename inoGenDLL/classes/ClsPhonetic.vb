Imports ColognePhoneticsSharp
Imports iText.Kernel.XMP.Impl
Imports NinjaNye.SearchExtensions
Imports XSoundex
Public Class ClsPhonetic
    Public Function GetNameCPhonetik(sName As String) As String
        If String.IsNullOrWhiteSpace(sName) Then
            Return ""
        End If

        Return ColognePhonetics.GetEncoding(sName).TrimStart("0"c)
    End Function

    Public Function GetNameCPhonetikRaw(sName As String) As String
        If String.IsNullOrWhiteSpace(sName) Then
            Return ""
        End If

        Return ColognePhonetics.GetEncoding(sName)
    End Function

    Public Function GetNameSoundex(sName As String) As String
        If String.IsNullOrWhiteSpace(sName) Then
            Return ""
        End If
        Return sName.ToSoundex
    End Function
End Class
