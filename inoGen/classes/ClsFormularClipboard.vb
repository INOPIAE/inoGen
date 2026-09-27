Public Class ClsFormularClipboard
    Private Shared _daten As ClsFormularDatenCopy = Nothing

    Private Shared _aktivesFormular As IFormularClipboard = Nothing

    '==================================================
    ' EVENT
    '==================================================

    Public Shared Event CopyCompleted(info As String)

    '--------------------------------------------------
    ' Aktives Formular setzen
    '--------------------------------------------------

    Public Shared Sub SetActiveForm(
        formular As IFormularClipboard)

        _aktivesFormular = formular

    End Sub


    '--------------------------------------------------
    ' Aktives Formular
    '--------------------------------------------------

    Public Shared ReadOnly Property AktivesFormular As IFormularClipboard

        Get
            Return _aktivesFormular
        End Get

    End Property


    '--------------------------------------------------
    ' Daten kopieren
    '--------------------------------------------------

    Public Shared Function Copy() As Boolean

        If _aktivesFormular Is Nothing Then
            Return False
        End If

        _daten = _aktivesFormular.CopyFormData()

        If _daten Is Nothing Then
            Return False
        End If

        RaiseEvent CopyCompleted(_daten.Datum)

        Return True

    End Function


    '--------------------------------------------------
    ' Daten vorhanden?
    '--------------------------------------------------

    Public Shared ReadOnly Property HatDaten As Boolean

        Get
            Return _daten IsNot Nothing
        End Get

    End Property


    '--------------------------------------------------
    ' Daten einfügen
    '--------------------------------------------------

    Public Shared Sub Paste()

        If _aktivesFormular Is Nothing Then Return

        If _daten Is Nothing Then Return

        _aktivesFormular.PasteFormData(_daten)

    End Sub
End Class
