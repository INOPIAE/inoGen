Imports System.Data.OleDb
Imports System.Text.RegularExpressions

Public Class clsDB

    Public connectionString As String = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", "")


    Public Sub New(dbFileString As String)
        connectionString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";Persist Security Info=True", dbFileString)
        ' Constructor logic if needed
    End Sub

    Private Function VornamenID(Vorname As String) As Int16
        Dim id As Integer = -1
        Dim sqlSelect As String = "SELECT tblVornameID FROM tblVorname WHERE Vorname = ?"
        Dim sqlInsert As String = "INSERT INTO tblVorname (Vorname) VALUES (?)"
        If Trim(Vorname) = "" Then
            Return 0
        End If
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(sqlSelect, conn)
                cmd.Parameters.AddWithValue("@Vorname", Vorname)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    id = Convert.ToInt32(result)
                    Return id
                End If
            End Using

            If MessageBox.Show(String.Format("Soll der Vorname '{0}' angelegt werden?", Vorname), "Vorname anlegen", MessageBoxButton.YesNo) = MessageBoxResult.No Then
                Return -1
            End If
            Using cmdInsert As New OleDbCommand(sqlInsert, conn)
                cmdInsert.Parameters.AddWithValue("@Vorname", Vorname)
                cmdInsert.ExecuteNonQuery()
            End Using

            Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                id = Convert.ToInt32(cmdId.ExecuteScalar())
            End Using
        End Using
        Return id
    End Function

    Public Function VornameAnlegen(Vornamen As String, PID As Int16?, ByRef Optional VT As String = "") As Int16

        If Trim(Vornamen) = "" Then
            Return 0
        End If

        Dim forbidden As Boolean = Regex.IsMatch(Vornamen, "(?<!^)\s*-\s*(?!$)")

        If forbidden Then
            MessageBox.Show("Der Bindestrich darf nur ohne Leerzeichen vorkommen")
            Return -1
        End If

        Dim Vorname() As String = Vornamen.Trim.Split(" "c)
        Dim Reihenfolge As Int16 = 1
        For Each VN In Vorname
            Dim Zeichen As String = ""
            Dim VName As String = VN
            If Reihenfolge = 1 Then
                If VName.Length < 4 Then
                    VT = VName.PadRight(4, "_"c)
                Else
                    VT = VName.Substring(0, 4)
                End If
            End If
            If VName.StartsWith("*") Then
                Zeichen = "*"
                VName = VName.Substring(1)
                If VName.Length < 4 Then
                    VT = VName.PadRight(4, "_"c)
                Else
                    VT = VName.Substring(0, 4)
                End If
            End If
            Dim VID As Int16 = VornamenID(VName)
            If VID = -1 Then Return VID
            If PID > 0 Then
                VornameSpeichern(VID, PID, Reihenfolge, Zeichen)
            End If
            Reihenfolge += 1
        Next
        Return 1
    End Function

    Private Function VornameSpeichern(VornameID As Int16, PersonID As Int16, Reihenfolge As Int16, Zeichen As String) As Int16
        Dim sqlFind As String = "SELECT tblPVornameID FROM  tblPVorname  WHERE tblVornameID = ? AND  tblPersonID = ?"
        Dim sqlInsert As String = "INSERT INTO tblPVorname (tblVornameID, tblPersonID, Reihenfolge, Zeichen) VALUES (?, ?, ?, ?)"
        Dim sqlUpdate As String = "UPDATE tblPVorname SET tblVornameID = ?, tblPersonID = ?, Reihenfolge = ?, Zeichen = ? WHERE tblPVornameID = ?"
        Dim PVID As Int16 = 0
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(sqlFind, conn)
                cmd.Parameters.AddWithValue("@tblVornameID", VornameID)
                cmd.Parameters.AddWithValue("@tblPersonID", PersonID)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    PVID = Convert.ToInt32(result)
                End If
            End Using
            If PVID = 0 Then

                Using cmdInsert As New OleDbCommand(sqlInsert, conn)
                    cmdInsert.Parameters.AddWithValue("@tblVornameID", VornameID)
                    cmdInsert.Parameters.AddWithValue("@tblPersonID", PersonID)
                    cmdInsert.Parameters.AddWithValue("@Reihenfolge", Reihenfolge)
                    cmdInsert.Parameters.AddWithValue("@Zeichen", Zeichen)
                    cmdInsert.ExecuteNonQuery()
                End Using

                Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                    PVID = Convert.ToInt32(cmdId.ExecuteScalar())
                End Using
            Else
                Using cmdUpdate As New OleDbCommand(sqlUpdate, conn)
                    cmdUpdate.Parameters.AddWithValue("@tblVornameID", VornameID)
                    cmdUpdate.Parameters.AddWithValue("@tblPersonID", PersonID)
                    cmdUpdate.Parameters.AddWithValue("@Reihenfolge", Reihenfolge)
                    cmdUpdate.Parameters.AddWithValue("@Zeichen", Zeichen)
                    cmdUpdate.Parameters.AddWithValue("@ID", PVID)
                    cmdUpdate.ExecuteNonQuery()
                End Using

            End If

        End Using

        Return PVID
    End Function

    Public Function NachnamenID(Nachname As String) As Int16
        Dim id As Integer = -1
        Dim sqlSelect As String = "SELECT tblNachnameID FROM tblNachname WHERE Nachname = ?"
        Dim sqlInsert As String = "INSERT INTO tblNachname (Nachname) VALUES (?)"
        If Trim(Nachname) = "" Then
            Return 0
        End If
        Using conn As New OleDbConnection(connectionString)
            conn.Open()

            Using cmd As New OleDbCommand(sqlSelect, conn)
                cmd.Parameters.AddWithValue("@Nachname", Nachname)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    id = Convert.ToInt32(result)
                    Return id
                End If
            End Using

            If MessageBox.Show(String.Format("Soll der Nachname '{0}' angelegt werden?", Nachname), "Nachname anlegen", MessageBoxButton.YesNo) = MessageBoxResult.No Then
                Return -1
            End If
            Using cmdInsert As New OleDbCommand(sqlInsert, conn)
                cmdInsert.Parameters.AddWithValue("@Nachname", Nachname)
                cmdInsert.ExecuteNonQuery()
            End Using

            Using cmdId As New OleDbCommand("SELECT @@IDENTITY", conn)
                id = Convert.ToInt32(cmdId.ExecuteScalar())
            End Using
        End Using
        Return id
    End Function
End Class
