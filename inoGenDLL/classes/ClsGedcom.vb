Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Text

Public Class ClsGedcom

    Implements IDisposable

    Private _gedcomFilePath As String
    Private swGedcom As StreamWriter
    Private disposedValue As Boolean
    Private cDB As ClsGenDB

    Public Sub New()
        ' Standardkonstrukt
        ' or
    End Sub
    Public Sub New(GedcomFilePath As String, DBFilePath As String)
        swGedcom = New StreamWriter(GedcomFilePath, False, New UTF8Encoding(False))
        swGedcom.AutoFlush = True

        cDB = New ClsGenDB(DBFilePath)
    End Sub

    Public Sub CreateGedcomFile(GedcomFilePath As String, DBFilePath As String, user As String)
        swGedcom = New StreamWriter(GedcomFilePath, False, New UTF8Encoding(False))
        swGedcom.AutoFlush = True
        cDB = New ClsGenDB(DBFilePath)
        AddHeader(user)
        GetPersons()
        GetFamily()
        AddTrailer()
    End Sub

    Public Sub AddHeader(user As String)
        swGedcom.WriteLine("0 HEAD")
        swGedcom.WriteLine("1 SOUR inoGEN")
        swGedcom.WriteLine("2 VERS 1.0")
        swGedcom.WriteLine("2 NAME inoGEN")
        swGedcom.WriteLine("1 SUBM " + user)
        swGedcom.WriteLine("1 GEDC")
        swGedcom.WriteLine("2 VERS 5.5.1")
        swGedcom.WriteLine("2 FORM LINEAGE-LINKED")
        swGedcom.WriteLine("1 CHAR UTF-8")
    End Sub

    Public Sub AddTrailer()
        swGedcom.WriteLine("0 TRLR")
        swGedcom.Close()
    End Sub

    Public Sub AddIndividual(individualData As Dictionary(Of String, Object))
        ' IndividualId ist Pflichtfeld
        If Not individualData.ContainsKey("IndividualId") Then
            Throw New ArgumentException("IndividualId ist erforderlich")
        End If

        swGedcom.WriteLine(String.Format("0 @I{0}@ INDI", individualData("IndividualId")))

        ' Name
        If individualData.ContainsKey("Name") AndAlso Not String.IsNullOrEmpty(individualData("Name").ToString()) Then
            swGedcom.WriteLine(String.Format("1 NAME {0}", individualData("Name")))

            If individualData.ContainsKey("GivenName") AndAlso Not String.IsNullOrEmpty(individualData("GivenName").ToString()) Then
                swGedcom.WriteLine(String.Format("2 GIVN {0}", individualData("GivenName")))
            End If

            If individualData.ContainsKey("Surname") AndAlso Not String.IsNullOrEmpty(individualData("Surname").ToString()) Then
                swGedcom.WriteLine(String.Format("2 SURN {0}", individualData("Surname")))
            End If
        End If

        ' Geschlecht
        If individualData.ContainsKey("Sex") AndAlso Not String.IsNullOrEmpty(individualData("Sex").ToString()) Then
            swGedcom.WriteLine(String.Format("1 SEX {0}", individualData("Sex")))
        End If

        ' Geburt
        If individualData.ContainsKey("BirthDate") OrElse individualData.ContainsKey("BirthPlace") Then
            swGedcom.WriteLine("1 BIRT")

            If individualData.ContainsKey("BirthDate") AndAlso Not String.IsNullOrEmpty(individualData("BirthDate").ToString()) Then
                swGedcom.WriteLine(String.Format("2 DATE {0}", ConvertDateToGedcom(individualData("BirthDate"))))
            End If

            If individualData.ContainsKey("BirthPlace") AndAlso Not String.IsNullOrEmpty(individualData("BirthPlace").ToString()) Then
                swGedcom.WriteLine(String.Format("2 PLAC {0}", individualData("BirthPlace")))
            End If
        End If

        ' Taufe
        If individualData.ContainsKey("ChristDate") OrElse individualData.ContainsKey("ChristPlace") Then
            swGedcom.WriteLine("1 CHR")
            If individualData.ContainsKey("ChristDate") AndAlso Not String.IsNullOrEmpty(individualData("ChristDate").ToString()) Then
                swGedcom.WriteLine(String.Format("2 DATE {0}", ConvertDateToGedcom(individualData("ChristDate"))))
            End If
            If individualData.ContainsKey("ChristPlace") AndAlso Not String.IsNullOrEmpty(individualData("ChristPlace").ToString()) Then
                swGedcom.WriteLine(String.Format("2 PLAC {0}", individualData("ChristPlace")))
            End If
        End If

        ' Tod
        If individualData.ContainsKey("DeathDate") OrElse individualData.ContainsKey("DeathPlace") Then
            swGedcom.WriteLine("1 DEAT")

            If individualData.ContainsKey("DeathDate") AndAlso Not String.IsNullOrEmpty(individualData("DeathDate").ToString()) Then
                swGedcom.WriteLine(String.Format("2 DATE {0}", ConvertDateToGedcom(individualData("DeathDate"))))
            End If

            If individualData.ContainsKey("DeathPlace") AndAlso Not String.IsNullOrEmpty(individualData("DeathPlace").ToString()) Then
                swGedcom.WriteLine(String.Format("2 PLAC {0}", individualData("DeathPlace")))
            End If
        End If

        ' Begräbnis
        If individualData.ContainsKey("BurialDate") OrElse individualData.ContainsKey("BurialPlace") Then
            swGedcom.WriteLine("1 BURI")
            If individualData.ContainsKey("BurialDate") AndAlso Not String.IsNullOrEmpty(individualData("BurialDate").ToString()) Then
                swGedcom.WriteLine(String.Format("2 DATE {0}", ConvertDateToGedcom(individualData("BurialDate"))))
            End If
            If individualData.ContainsKey("BurialPlace") AndAlso Not String.IsNullOrEmpty(individualData("BurialPlace").ToString()) Then
                swGedcom.WriteLine(String.Format("2 PLAC {0}", individualData("BurialPlace")))
            End If
        End If

        ' Familien-Verweise
        If individualData.ContainsKey("FamilySIds") Then
            Dim familySIds = TryCast(individualData("FamilySIds"), IEnumerable(Of String))
            If familySIds IsNot Nothing Then
                For Each familyId As String In familySIds
                    swGedcom.WriteLine(String.Format("1 FAMS @F{0}@", familyId))
                Next
            End If
        End If

        If individualData.ContainsKey("FamilyCId") Then
            swGedcom.WriteLine(String.Format("1 FAMC @F{0}@", individualData("FamilyCId")))
        End If
    End Sub

    Sub Basis(filepath As String)
        ' UTF-8 ohne BOM ist der Standard für moderne GEDCOM-Dateien
        Using writer As New StreamWriter(filepath, False, New UTF8Encoding(False))

            ' --- KOPFZEILE (HEADER) ---
            writer.WriteLine("0 HEAD")
            writer.WriteLine("1 SOUR VB_NET_CREATOR")
            writer.WriteLine("2 VERS 1.0")
            writer.WriteLine("2 NAME Mein Stammbaum Generator")
            writer.WriteLine("1 SUBM Donald Duck")
            writer.WriteLine("1 GEDC")
            writer.WriteLine("2 VERS 5.5.1")
            writer.WriteLine("2 FORM LINEAGE-LINKED")
            writer.WriteLine("1 CHAR UTF-8")

            ' --- PERSON 1: VATER ---
            writer.WriteLine("0 @I1@ INDI")
            writer.WriteLine("1 NAME Max /Mustermann/")
            writer.WriteLine("2 GIVN Max")
            writer.WriteLine("2 SURN Mustermann")
            writer.WriteLine("1 SEX M")
            writer.WriteLine("1 BIRT")
            writer.WriteLine("2 DATE 12 MAY 1950")
            writer.WriteLine("2 PLAC Berlin, Deutschland")
            writer.WriteLine("1 FAMS @F1@") ' Verweist auf die Familie F1 als Ehemann

            ' --- PERSON 2: SOHN ---
            writer.WriteLine("0 @I2@ INDI")
            writer.WriteLine("1 NAME Erika /Mustermann/")
            writer.WriteLine("2 GIVN Erika")
            writer.WriteLine("2 SURN Mustermann")
            writer.WriteLine("1 SEX F")
            writer.WriteLine("1 BIRT")
            writer.WriteLine("2 DATE 24 AUG 1985")
            writer.WriteLine("2 PLAC Hamburg, Deutschland")
            writer.WriteLine("1 FAMC @F1@") ' Verweist auf die Familie F1 als Kind

            ' --- FAMILIEN-VERKNÜPFUNG (FAMILY RECORD) ---
            writer.WriteLine("0 @F1@ FAM")
            writer.WriteLine("1 HUSB @I1@") ' Vater zuweisen
            writer.WriteLine("1 CHIL @I2@") ' Kind zuweisen

            ' --- DATEIENDE (TRAILER) ---
            writer.WriteLine("0 TRLR")

        End Using

        Console.WriteLine("GEDCOM-Datei erfolgreich erstellt!")
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' Managed Resources freigeben
                If swGedcom IsNot Nothing Then
                    swGedcom.Flush()
                    swGedcom.Close()
                    swGedcom.Dispose()
                    swGedcom = Nothing
                End If
            End If
            disposedValue = True
        End If
    End Sub

    Public Function ConvertDateToGedcom(germanDate As String) As String
        If String.IsNullOrEmpty(germanDate) Then
            Return String.Empty
        End If

        If germanDate.StartsWith("um") Then
            germanDate = germanDate.Replace("um", "ABT ").Trim()
        End If

        If germanDate.StartsWith("vor") Then
            germanDate = germanDate.Replace("vor", "BEF ").Trim()
        End If

        If germanDate.StartsWith("<") Then
            germanDate =  germanDate.Replace("<", "BEF ").Trim()
        End If

        If germanDate.StartsWith("nach") Then
            germanDate = germanDate.Replace("nach", "AFT ").Trim()
        End If

        If germanDate.StartsWith(">") Then
            germanDate = germanDate.Replace(">", "AFT ").Trim()
        End If

        If germanDate.StartsWith("err") Then
            germanDate = germanDate.Replace("err", "CAL ").Trim()
        End If

        germanDate = System.Text.RegularExpressions.Regex.Replace(germanDate, "\s+", " ")


        Dim monthNames As String() = {"JAN", "FEB", "MAR", "APR", "MAY", "JUN",
                                   "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"}

        Dim DateParts As String() = germanDate.Split(" ")

        Dim DatePart As String = String.Empty
        Dim calculatedDate As String = String.Empty
        If DateParts.Length = 1 Then
            DatePart = DateParts(0)
        Else
            DatePart = DateParts(1)
            calculatedDate = DateParts(0) & " "
        End If


        Try
            Dim splitDate As String() = DatePart.Split(".")
            Select Case splitDate.Length
                Case 1 ' Nur Jahr
                    Return calculatedDate & splitDate(0)
                Case 2 ' Monat und Jahr
                    Dim month As Integer = Integer.Parse(splitDate(0))
                    Dim year As Integer = Integer.Parse(splitDate(1))
                    Return calculatedDate & String.Format("{0} {1}", monthNames(month - 1), year)
                Case 3 ' Tag, Monat und Jahr
                    Dim day As Integer = Integer.Parse(splitDate(0))
                    Dim month As Integer = Integer.Parse(splitDate(1))
                    Dim year As Integer = Integer.Parse(splitDate(2))
                    Return calculatedDate & String.Format("{0} {1} {2}", day, monthNames(month - 1), year)
                Case Else
                    Return germanDate
            End Select
        Catch ex As Exception
            Return germanDate
        End Try
    End Function

    Public Sub GetPersons()
        Dim dtPersons As DataTable = cDB.GetGedcomPersonenEreignisse()
        Dim PID As Long
        Dim individualData As New Dictionary(Of String, Object)
        Dim dtFamily As DataTable
        Dim familyIds As List(Of String)
        Dim z As Long
        For Each row As DataRow In dtPersons.Rows
            z += 1
            If PID <> row("tblPersonID") Then
                If PID > 0 Then
                    dtFamily = cDB.GetGedcomPersonFamilie(PID)

                    ' Liste für FamilyIds erstellen
                    familyIds = New List(Of String)

                    For Each frow As DataRow In dtFamily.Rows
                        familyIds.Add(frow("tblFamilieID").ToString()) ' Passen Sie "FamilyID" an Ihren Spaltennamen an
                    Next

                    ' Zum Dictionary hinzufügen (nur wenn Familien vorhanden)
                    If familyIds.Count > 0 Then
                        individualData.Add("FamilySIds", familyIds)
                    End If
                    'File.AppendAllText("D:\test\test.log", "PID: " & PID & Environment.NewLine)
                    ' Add the previous individual to the GEDCOM file
                    AddIndividual(individualData)
                End If

                PID = row("tblPersonID")
                individualData = New Dictionary(Of String, Object)

                Dim Surname As String = If(row("Nachname") IsNot DBNull.Value, row("Nachname").ToString(), "N.")

                individualData.Add("IndividualId", row("tblPersonID"))
                individualData.Add("Name", row("Vorname") & " /" & Surname & "/")
                individualData.Add("GivenName", row("Vorname"))
                individualData.Add("Surname", Surname)
                If row("tblFamilieID") > 0 Then
                    individualData.Add("FamilyCId", row("tblFamilieID"))
                End If
                individualData.Add("Sex", UCase(row("Sex")))


            End If
            Try
                If row("tblEreignisArtID") IsNot DBNull.Value Then
                    Select Case row("tblEreignisArtID")
                        Case 1 ' Geburt
                            individualData.Add("BirthDate", row("DatumText"))
                            individualData.Add("BirthPlace", row("Ort"))
                        Case 2 ' Taufe
                            individualData.Add("ChristDate", row("DatumText"))
                            individualData.Add("ChristPlace", row("Ort"))
                        Case 6 ' Tod
                            individualData.Add("DeathDate", row("DatumText"))
                            individualData.Add("DeathPlace", row("Ort"))
                        Case 7 'Begräbnis
                            individualData.Add("BurialDate", row("DatumText"))
                            individualData.Add("BurialPlace", row("Ort"))
                        Case Else

                    End Select
                End If
            Catch ex As Exception

            End Try

        Next

        dtFamily = cDB.GetGedcomPersonFamilie(PID)

        ' Liste für FamilyIds erstellen
        familyIds = New List(Of String)

        For Each frow As DataRow In dtFamily.Rows
            familyIds.Add(frow("tblFamilieID").ToString()) ' Passen Sie "FamilyID" an Ihren Spaltennamen an
        Next

        ' Zum Dictionary hinzufügen (nur wenn Familien vorhanden)
        If familyIds.Count > 0 Then
            individualData.Add("FamilySIds", familyIds)
        End If

        ' Add the previous individual to the GEDCOM file
        AddIndividual(individualData)
    End Sub

    Public Sub AddFamiliy(individualData As Dictionary(Of String, Object))
        ' IndividualId ist Pflichtfeld
        If Not individualData.ContainsKey("FId") Then
            Throw New ArgumentException("FId ist erforderlich")
        End If

        swGedcom.WriteLine(String.Format("0 @F{0}@ FAM", individualData("FId")))

        ' Heirat
        If individualData.ContainsKey("MarriageDate") OrElse individualData.ContainsKey("MarriagePlace") Then
            swGedcom.WriteLine("1 MARR")
            swGedcom.WriteLine("2 TYPE Common Law")
            If individualData.ContainsKey("MarriageDate") AndAlso Not String.IsNullOrEmpty(individualData("MarriageDate").ToString()) Then
                swGedcom.WriteLine(String.Format("2 DATE {0}", ConvertDateToGedcom(individualData("MarriageDate"))))
            End If
            If individualData.ContainsKey("MarriagePlace") AndAlso Not String.IsNullOrEmpty(individualData("MarriagePlace").ToString()) Then
                swGedcom.WriteLine(String.Format("2 PLAC {0}", individualData("MarriagePlace")))
            End If
        End If

        'Heirat kirchlich
        If individualData.ContainsKey("MarriageChrDate") OrElse individualData.ContainsKey("MarriageChrPlace") Then
            swGedcom.WriteLine("1 MARR")
            If individualData.ContainsKey("MarriageChrDate") AndAlso Not String.IsNullOrEmpty(individualData("MarriageChrDate").ToString()) Then
                swGedcom.WriteLine(String.Format("2 DATE {0}", ConvertDateToGedcom(individualData("MarriageChrDate"))))
            End If
            If individualData.ContainsKey("MarriageChrPlace") AndAlso Not String.IsNullOrEmpty(individualData("MarriageChrPlace").ToString()) Then
                swGedcom.WriteLine(String.Format("2 PLAC {0}", individualData("MarriageChrPlace")))
            End If
        End If

        ' Scheidung
        If individualData.ContainsKey("DivorceDate") OrElse individualData.ContainsKey("DivorcePlace") Then
            swGedcom.WriteLine("1 DIV")
            If individualData.ContainsKey("DivorceDate") AndAlso Not String.IsNullOrEmpty(individualData("DivorceDate").ToString()) Then
                swGedcom.WriteLine(String.Format("2 DATE {0}", ConvertDateToGedcom(individualData("DivorceDate"))))
            End If
            If individualData.ContainsKey("DivorcePlace") AndAlso Not String.IsNullOrEmpty(individualData("DivorcePlace").ToString()) Then
                swGedcom.WriteLine(String.Format("2 PLAC {0}", individualData("DivorcePlace")))
            End If
        End If

        ' Verlobung
        If individualData.ContainsKey("EngagementDate") OrElse individualData.ContainsKey("EngagementPlace") Then
            swGedcom.WriteLine("1 ENGA")
            If individualData.ContainsKey("EngagementDate") AndAlso Not String.IsNullOrEmpty(individualData("EngagementDate").ToString()) Then
                swGedcom.WriteLine(String.Format("2 DATE {0}", ConvertDateToGedcom(individualData("EngagementDate"))))
            End If
            If individualData.ContainsKey("EngagementPlace") AndAlso Not String.IsNullOrEmpty(individualData("EngagementPlace").ToString()) Then
                swGedcom.WriteLine(String.Format("2 PLAC {0}", individualData("EngagementPlace")))
            End If
        End If


        ' Familien-Verweise


        If individualData.ContainsKey("VId") Then
            swGedcom.WriteLine(String.Format("1 HUSB @I{0}@", individualData("VId")))
        End If

        If individualData.ContainsKey("MId") Then
            swGedcom.WriteLine(String.Format("1 WIFE @I{0}@", individualData("MId")))
        End If

        If individualData.ContainsKey("ChildIds") Then
            Dim ChildIds = TryCast(individualData("ChildIds"), IEnumerable(Of String))
            If ChildIds IsNot Nothing Then
                For Each PId As String In ChildIds
                    swGedcom.WriteLine(String.Format("1 CHIL @I{0}@", PId))
                Next
            End If
        End If

    End Sub


    Public Sub GetFamily()
        Dim dtFamily As DataTable = cDB.GetGedcomFamilie()
        Dim FID As Long
        Dim individualData As New Dictionary(Of String, Object)
        Dim ChildIds As New List(Of String)

        For Each row As DataRow In dtFamily.Rows
            If FID <> row("tblFamilieID") Then
                If FID > 0 Then
                    If ChildIds.Count > 0 Then
                        individualData.Add("ChildIds", ChildIds)
                    End If

                    ' Add the previous individual to the GEDCOM file
                    AddFamiliy(individualData)
                End If

                FID = row("tblFamilieID")
                individualData = New Dictionary(Of String, Object)
                ChildIds = New List(Of String)

                individualData.Add("FId", FID)
                If row("tblPersonIDV") IsNot DBNull.Value Then
                    individualData.Add("VId", row("tblPersonIDV"))
                End If
                If row("tblPersonIDM") IsNot DBNull.Value Then
                    individualData.Add("MId", row("tblPersonIDM"))
                End If

                Dim dtFamilyEvents As DataTable = cDB.GetGedcomFamilieEreignisse(FID)
                For Each eventRow As DataRow In dtFamilyEvents.Rows
                    If eventRow("tblEreignisArtID") IsNot DBNull.Value Then
                        Select Case eventRow("tblEreignisArtID")
                            Case 3 ' Hochzeit
                                individualData.Add("MarriageDate", eventRow("DatumText"))
                                individualData.Add("MarriagePlace", eventRow("Ort"))
                            Case 4 ' Heirat kirchlich
                                individualData.Add("MarriageChrDate", eventRow("DatumText"))
                                individualData.Add("MarriageChrPlace", eventRow("Ort"))
                            Case 5 ' Scheidung
                                individualData.Add("DivorceDate", eventRow("DatumText"))
                                individualData.Add("DivorcePlace", eventRow("Ort"))
                            Case 8 ' Verlobung
                                individualData.Add("EngagementDate", eventRow("DatumText"))
                                individualData.Add("EngagementPlace", eventRow("Ort"))
                            Case Else
                                ' Andere Ereignisse können hier behandelt werden
                        End Select
                    End If
                Next

            End If

            If row("tblPersonID") IsNot DBNull.Value Then
                ChildIds.Add(row("tblPersonID").ToString())
            End If

        Next

        If ChildIds.Count > 0 Then
            individualData.Add("ChildIds", ChildIds)
        End If

        ' Add the previous individual to the GEDCOM file
        AddFamiliy(individualData)
    End Sub
End Class
