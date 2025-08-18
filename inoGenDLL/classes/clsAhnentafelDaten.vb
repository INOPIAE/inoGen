Imports System.Data.OleDb
Imports System.Diagnostics.Eventing.Reader
Imports System.IO
Imports System.Reflection.Emit

Public Class clsAhnentafelDaten

    Public DBPath As String =""
    Public connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", "")

    Public cGenDB As inoGenDLL.ClsGenDB

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
    End Structure

    Public Persons As New List(Of PersonData)
    Public Kinder As New List(Of PersonData)
    Public Ehe As New List(Of Integer)
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
        End While
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
                writer.WriteLine("# " & p.Pos & ". " & OutputVorname(p.Vorname) & " " & p.Nachname.ToUpper)
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

End Class
