Imports System.Data.OleDb
Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Reflection.Emit
Imports System.Security.Cryptography
Imports inoGenDLL.ClsOSMKarte

Public Class clsAhnentafelDaten

    Public DBPath As String =""
    Public connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", "")

    Public cGenDB As inoGenDLL.ClsGenDB
    Public cOSM As New inoGenDLL.ClsOSMKarte


    Public Structure PersonData
        Public ID As Integer
        Public PS As String
        Public Nachname As String
        Public Vorname As String
        Public Geschlecht As String
        Public Geburtsdatum As String
        Public Geburtsort As String
        Public Taufdatum As String
        Public Taufort As String
        Public Sterbedatum As String
        Public Sterbeort As String
        Public Begräbnisdatum As String
        Public Begräbnisort As String
        Public Konfession As String
        Public Heiratdatum As String
        Public Heiratort As String
        Public KHeiratdatum As String
        Public KHeiratort As String
        Public Scheidungsdatum As String
        Public Scheidungsort As String
        Public Verlobungsdatum As String
        Public Verlobungsort As String
        Public Sonstige As Boolean
        Public FID As Integer 'Eltern-Familie ID
        Public Pos As Long
        Public Gen As Long
        Public V As Long
        Public M As Long
        Public EP As Long
        Public EID As Integer 'Eigene Familie ID
        Public FSID As String
    End Structure

    Public Structure EventData
        Public ID As Integer
        Public Person As Boolean
        Public EventID As Integer
        Public Eventname As String
        Public EventDate As String
        Public EventLocation As String
        Public EventTopic As String
    End Structure

    Public Structure FamilyData
        Public ID As Integer
        Public VID As Integer?
        Public MID As Integer?
    End Structure

    Public Persons As New List(Of PersonData)
    Public Kinder As New List(Of PersonData)
    Public Ehe As New List(Of Integer)
    Public LocationList As New List(Of ClsOSMKarte.marker)
    Public EventList As New List(Of EventData)
    Public FamilyList As New List(Of FamilyData)

    Public Property RootPersonID As Integer = 0

    Public Sub New(dbFileString As String)
        DBPath = dbFileString
        connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", DBPath)
        cGenDB = New inoGenDLL.ClsGenDB(DBPath)
    End Sub
    Public Function GetPersonByID(id As Integer) As PersonData?
        For Each person In Persons
            If person.ID = id Then
                Return person
            End If
        Next
        Return Nothing ' Person nicht gefunden
    End Function

    Public Sub AddPerson(person As PersonData)
        Persons.Add(person)
    End Sub

    Public Sub NewList()
        Persons.Clear()
        Dim rootP As New PersonData
        rootP.ID = RootPersonID
        rootP.Pos = 1
        rootP.Gen = 1
        cGenDB.FillPersonData(rootP)
        Persons.Add(rootP)
        Dim index As Integer = 0

        While index < Persons.Count
            Dim p = Persons(index)

            If p.V > 0 Then
                Dim newP As New PersonData
                newP.ID = p.V
                newP.Pos = p.Pos * 2
                newP.Gen = p.Gen + 1
                newP.EID = p.FID
                cGenDB.FillPersonData(newP)
                Persons.Add(newP)
            End If

            If p.M > 0 Then
                Dim newP As New PersonData
                newP.ID = p.M
                newP.Pos = p.Pos * 2 + 1
                newP.Gen = p.Gen + 1
                newP.EID = p.FID
                cGenDB.FillPersonData(newP)
                Persons.Add(newP)
            End If

            index += 1
            If index > 1000 Then
                Exit While ' Sicherheitsabfrage, um endlose Schleifen zu vermeiden
            End If
        End While
        ErstelleLocationList()
    End Sub

    Public Function NewPerson(ID As Long, Pos As Long) As PersonData
        Dim np As New PersonData()
        np.ID = ID
        np.Pos = Pos
        cGenDB.FillPersonData(np)

        Return np
    End Function

    Public Sub PrintTree()
        For Each p In Persons.OrderBy(Function(x) x.Pos)
            Dim gen As Integer = CInt(Math.Log(p.Pos, 2)) ' Generation = Tiefe im Binärbaum
            Console.WriteLine(New String(" "c, gen * 4) &
                              $"[{p.Pos}] {p.Vorname} {p.Nachname} (V:{p.V}, M:{p.M})")
            Debug.Print(New String(" "c, p.Gen * 4) &
                              $"[{p.Pos}] {p.Vorname} {p.Nachname} (V:{p.V}, M:{p.M})")
        Next
    End Sub

    Public Sub WriteTreeToFile(filePath As String)
        Ehe.Clear()
        Dim Generation As Integer = 0
        Using writer As New StreamWriter(filePath, False, System.Text.Encoding.UTF8)
            For Each p In Persons.OrderBy(Function(x) x.Pos)
                If Generation <> p.Gen Then
                    Generation = p.Gen
                    writer.WriteLine("# " & ToRoman(Generation) & ". Generation")
                End If
                writer.WriteLine("# " & p.Pos & ". " & OutputVorname(p.Vorname) & " " & p.Nachname.ToUpper & " " & FamilySearchLinkPerson(p.FSID))

                AusgabePersDetails(writer, p)

                If p.EID > 0 And Ehe.Contains(p.EID) = False Then
                    If IsNothing(p.Verlobungsdatum) = False Or IsNothing(p.Verlobungsort) = False Then
                        writer.WriteLine("⚬ " & If(IsNothing(p.Verlobungsdatum), "    ", p.Verlobungsdatum) & " " & If(IsNothing(p.Verlobungsort), "", p.Verlobungsort) & vbCrLf)
                    End If
                    If IsNothing(p.Heiratdatum) = False Or IsNothing(p.Heiratort) = False Then
                        writer.WriteLine("⚭ " & If(IsNothing(p.Heiratdatum), "    ", p.Heiratdatum) & " " & If(IsNothing(p.KHeiratort), "", p.KHeiratort) & vbCrLf)
                    End If
                    If IsNothing(p.KHeiratdatum) = False Or IsNothing(p.KHeiratort) = False Then
                        writer.WriteLine("♁⚭ " & If(IsNothing(p.KHeiratdatum), "    ", p.KHeiratdatum) & " " & If(IsNothing(p.KHeiratort), "", p.KHeiratort) & vbCrLf)
                    End If
                    If IsNothing(p.Scheidungsdatum) = False Or IsNothing(p.Scheidungsort) = False Then
                        writer.WriteLine("⚮ " & If(IsNothing(p.Scheidungsdatum), "    ", p.Scheidungsdatum) & " " & If(IsNothing(p.Scheidungsort), "", p.Scheidungsort) & vbCrLf)
                    End If
                    writer.WriteLine("## Kinder")
                    AddChildren(p.EID)
                    For Each k In Kinder.OrderBy(Function(x) x.Pos)
                        writer.WriteLine("## " & k.Pos & " " & OutputVorname(k.Vorname) & " " & k.Nachname.ToUpper)
                        If GetPersonByID(k.ID) Is Nothing Then
                            AusgabePersDetails(writer, k)
                        Else
                            writer.WriteLine("siehe oben" & vbCrLf)
                        End If
                        Ehe.Add(p.EID)
                    Next
                End If
            Next
        End Using
    End Sub

    Public Sub WriteCompTreeToFile(filePath As String)
        Ehe.Clear()
        Dim Generation As Integer = 0
        Using writer As New StreamWriter(filePath, False, System.Text.Encoding.UTF8)
            For Each p In Persons.OrderBy(Function(x) x.Pos)
                If Generation <> p.Gen Then
                    Generation = p.Gen
                    writer.WriteLine("# " & ToRoman(Generation) & ". Generation")
                End If
                writer.WriteLine("# " & p.Pos & ". " & OutputVorname(p.Vorname) & " " & p.Nachname.ToUpper)
                AusgabePersCompDetails(writer, p)

                If p.EID > 0 And Ehe.Contains(p.EID) = False Then
                    If IsNothing(p.Heiratdatum) = False Or IsNothing(p.Heiratort) = False Then
                        writer.WriteLine("⚭ " & If(IsNothing(p.Heiratdatum), "    ", p.Heiratdatum) & " " & If(IsNothing(p.KHeiratort), "", p.KHeiratort) & vbCrLf)
                    End If
                    If IsNothing(p.KHeiratdatum) = False Or IsNothing(p.KHeiratort) = False Then
                        writer.WriteLine("♁⚭ " & If(IsNothing(p.KHeiratdatum), "    ", p.KHeiratdatum) & " " & If(IsNothing(p.KHeiratort), "", p.KHeiratort) & vbCrLf)
                    End If
                End If
            Next
        End Using
    End Sub

    Public Sub AusgabePersDetails(writer As StreamWriter, p As PersonData)
        If IsNothing(p.Geburtsdatum) = False Or IsNothing(p.Geburtsort) = False Then
            writer.WriteLine("∗ " & If(IsNothing(p.Geburtsdatum), "    ", p.Geburtsdatum) & " " & If(IsNothing(p.Geburtsort), "", p.Geburtsort) & vbCrLf) '★
        End If
        If IsNothing(p.Taufdatum) = False Or IsNothing(p.Taufort) = False Then
            writer.WriteLine("~ " & If(IsNothing(p.Taufdatum), "    ", p.Taufdatum) & " " & If(IsNothing(p.Taufort), "", p.Taufort) & vbCrLf)
        End If
        If IsNothing(p.Sterbedatum) = False Or IsNothing(p.Sterbeort) = False Then
            writer.WriteLine("† " & If(IsNothing(p.Sterbedatum), "    ", p.Sterbedatum) & " " & If(IsNothing(p.Sterbeort), "", p.Sterbeort) & vbCrLf)
        End If
        If IsNothing(p.Begräbnisdatum) = False Or IsNothing(p.Begräbnisort) = False Then
            writer.WriteLine("⚰ " & If(IsNothing(p.Begräbnisdatum), "    ", p.Begräbnisdatum) & " " & If(IsNothing(p.Begräbnisort), "", p.Begräbnisort) & vbCrLf)
        End If
    End Sub

    Public Sub AusgabePersCompDetails(writer As StreamWriter, p As PersonData)
        Dim birth As String = ""
        Dim death As String = ""
        If IsNothing(p.Geburtsdatum) = False Or IsNothing(p.Geburtsort) = False Then
            birth = "∗ " & If(IsNothing(p.Geburtsdatum), "    ", p.Geburtsdatum) & " " & If(IsNothing(p.Geburtsort), "", p.Geburtsort)  '★
        End If
        If IsNothing(p.Taufdatum) = False Or IsNothing(p.Taufort) = False And birth = "" Then
            birth = "~ " & If(IsNothing(p.Taufdatum), "    ", p.Taufdatum) & " " & If(IsNothing(p.Taufort), "", p.Taufort)
        End If
        If IsNothing(p.Sterbedatum) = False Or IsNothing(p.Sterbeort) = False Then
            death = "† " & If(IsNothing(p.Sterbedatum), "    ", p.Sterbedatum) & " " & If(IsNothing(p.Sterbeort), "", p.Sterbeort)
        End If
        If IsNothing(p.Begräbnisdatum) = False Or IsNothing(p.Begräbnisort) = False And death = "" Then
            death = "⚰ " & If(IsNothing(p.Begräbnisdatum), "    ", p.Begräbnisdatum) & " " & If(IsNothing(p.Begräbnisort), "", p.Begräbnisort)
        End If
        writer.WriteLine(birth & " " & death & vbCrLf)
    End Sub
    Public Sub AddChildren(FID As Integer)
        Dim strSQL As String = "SELECT
                tblPersonID,
                PS,
                Sex,
                Vorname
            FROM
                tblPerson
            WHERE tblFamilieID = ?

            ORDER BY Right(PS, 4);"

        Kinder.Clear()
        Dim Pos As Long = 1
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(strSQL, conn)
                cmd.Parameters.AddWithValue("?", FID) 'Parameter einsetzen

                Using rdr As OleDbDataReader = cmd.ExecuteReader()
                    While rdr.Read()
                        Dim newP As New PersonData
                        newP.ID = rdr("tblPersonID")
                        newP.Pos = Pos
                        cGenDB.FillPersonData(newP)
                        Kinder.Add(newP)
                        Pos += 1
                    End While
                End Using
            End Using
        End Using

    End Sub

    Public Function OutputVorname(vorname As String) As String
        Dim words As List(Of String) = vorname.Split(" "c).ToList()
        Dim output As New List(Of String)()

        Dim i As Integer = 0
        While i < words.Count
            If words(i).StartsWith("*") Then
                If i <= words.Count Then
                    Dim nextWord = words(i)
                    output.Add(nextWord & "*")
                    i += 1
                    Continue While
                End If
            End If
            output.Add(words(i))
            i += 1
        End While

        Dim result As String = String.Join(" ", output)
        Return result
    End Function

    Public Function ToRoman(ByVal number As Integer) As String
        If number <= 0 Or number > 3999 Then
            Throw New ArgumentOutOfRangeException("number", "Nur die Werte 1 - 3999 sind erlaubt.")
        End If

        Dim romanNumerals As (Value As Integer, Symbol As String)() = {
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I")
    }

        Dim result As New System.Text.StringBuilder()

        For Each rn In romanNumerals
            While number >= rn.Value
                result.Append(rn.Symbol)
                number -= rn.Value
            End While
        Next

        Return result.ToString()
    End Function


    Public Sub WriteToCSV(filePath As String)
        Ehe.Clear()
        Dim Generation As Integer = 0
        Using writer As New StreamWriter(filePath, False, System.Text.Encoding.UTF8)
            Dim strTitel As String = "Nr;Vorname;Nachname;Gechlecht;Geburtsdatum;Geburtsort;GebKom;Taufe;Taufort;TaufKom;Konfession;Sterbedatum;Sterbeort;SterbeKom;Heirat;Heiratsort;HeiratKom;Heirat k;Heiratsort k;HeiratKom k;Beruf;Familysearch;VNr;MNr;EPNr"
            writer.WriteLine(strTitel)
            For Each p In Persons.OrderBy(Function(x) x.Pos)
                Dim strLine As String = p.ID & ";" & p.Vorname & ";" & p.Nachname.ToUpper & ";" & p.Geschlecht & ";"
                strLine &= AusgabePersDetailsCSV(p)

                strLine &= If(IsNothing(p.Heiratdatum), "", p.Heiratdatum) & ";" & If(IsNothing(p.KHeiratort), "", p.KHeiratort) & ";;"

                strLine &= If(IsNothing(p.KHeiratdatum), "", p.KHeiratdatum) & ";" & If(IsNothing(p.KHeiratort), "", p.KHeiratort) & ";;"

                strLine &= ";;" & p.V & ";" & p.M & ";" & p.EP
                writer.WriteLine(strLine)
            Next
        End Using
    End Sub

    Public Function AusgabePersDetailsCSV(p As PersonData) As String
        Dim result As String = ""

        result = If(IsNothing(p.Geburtsdatum), "", p.Geburtsdatum) & ";" & If(IsNothing(p.Geburtsort), "", p.Geburtsort) & ";;"

        result &= If(IsNothing(p.Taufdatum), "", p.Taufdatum) & ";" & If(IsNothing(p.Taufort), "", p.Taufort) & ";;;"

        result &= If(IsNothing(p.Sterbedatum), "", p.Sterbedatum) & ";" & If(IsNothing(p.Sterbeort), "", p.Sterbeort) & ";;"

        Return result
    End Function

    Public Function FamilySearchLinkPerson(FSID As String) As String
        If IsNothing(FSID) Or FSID = "" Then
            Return ""
        Else
            Return String.Format("[{0}](https://www.familysearch.org/tree/person/details/{0})", FSID)
        End If
    End Function

    Public Sub ErstelleLocationList()
        LocationList = New List(Of ClsOSMKarte.marker)
        Dim strSql As String = "SELECT
            tblEreignis.tblPersonID,
            tblEreignis.tblFamilieID,
            tblEreignis.tblOrtID,
            tblOrt.Ort,
            tblOrt.Breite,
            tblOrt.Laenge
        FROM
            tblEreignis
            INNER JOIN tblOrt ON tblEreignis.tblOrtID = tblOrt.tblOrtID"
        For Each p In Persons
            Using conn As New OleDbConnection(connectionString)
                conn.Open()

                Using cmd As New OleDbCommand(strSql & " WHERE tblPersonID = ?", conn)
                    cmd.Parameters.AddWithValue("?", p.ID) 'Parameter einsetzen

                    Using rdr As OleDbDataReader = cmd.ExecuteReader()
                        While rdr.Read()
                            If rdr("Breite") Is DBNull.Value Or rdr("Laenge") Is DBNull.Value Then
                                Continue While
                            End If
                            Dim marker As New ClsOSMKarte.marker
                            marker.lat = Convert.ToDouble(rdr("Breite"))
                            marker.lon = Convert.ToDouble(rdr("Laenge"))
                            marker.title = rdr("Ort")
                            marker.Count = 1
                            AddMarkerIfNotExists(marker)
                        End While
                    End Using
                End Using
            End Using

        Next
    End Sub

    Public Sub AddMarkerIfNotExists(newMarker As marker)
        ' Prüfen ob Titel schon vorhanden
        Dim exists = LocationList.Any(Function(m) m.title = newMarker.title)

        If Not exists Then
            LocationList.Add(newMarker)
        Else
            ' Optional: den vorhandenen Count erhöhen
            Dim existing = LocationList.First(Function(m) m.title = newMarker.title)
            Dim updated = existing
            updated.Count += 1
            ' alten ersetzen
            Dim idx = LocationList.IndexOf(existing)
            LocationList(idx) = updated
        End If
    End Sub

    Public Function CalculateChild(person As PersonData) As Long
        Dim childPos As Long = 1
        Select Case person.Gen
            Case 2, 5, 8, 11, 14
                childPos = person.Pos \ 2
            Case 3, 6, 9, 12, 15
                childPos = person.Pos \ 4
            Case 4, 7, 10, 13, 16
                childPos = person.Pos \ 8
        End Select
        Return childPos
    End Function

    Public Function CalculateChildPosChart(person As PersonData) As Long
        Dim childPos As Long = 1
        Dim Versatz As Long = 0
        Dim Start As Long = 0
        Dim VersatzKorrektur As Long = 0
        Select Case person.Gen
            Case 4
                Return person.Pos - 1
            Case 5, 8, 11, 14
                Start = 2 ^ (person.Gen - 1) - 1
                Versatz = 2
                VersatzKorrektur = 0
            Case 6, 9, 12, 15
                Start = 2 ^ (person.Gen - 1) - 3
                Versatz = 4
                VersatzKorrektur = 2
            Case 7, 10, 13, 16
                Start = 2 ^ (person.Gen - 1) - 7
                Versatz = 8
                VersatzKorrektur = 6
            Case Else
                Return person.Pos - 1
        End Select
        childPos = person.Pos - Start
        Do Until childPos <= Versatz + VersatzKorrektur
            childPos = childPos - Versatz

        Loop
        Return childPos
    End Function



End Class
