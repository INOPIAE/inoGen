Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Net.Mime.MediaTypeNames
Imports ADOX

Public Class ClsDatabase
    Private connString As String
    Private dbFile As String

    Private sqlPath As String = AppDomain.CurrentDomain.BaseDirectory.Replace("\inoGen\bin\Debug\net8.0-windows\", "") & "\inoGenDLL\SQL\"

    Public Sub New(dbFileString As String)
        connString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";Persist Security Info=True", dbFileString)
        dbFile = dbFileString
        If AppDomain.CurrentDomain.BaseDirectory.Contains("TestInoCook") Then
            sqlPath = AppDomain.CurrentDomain.BaseDirectory.Replace("\TestInoCook\bin\Debug\net8.0\", "") & "\inoGenDLL\SQL\"
        End If
    End Sub

    Public Function CreateDB() As String
        Dim cat As Catalog = New Catalog()
        cat.Create(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";", dbFile))
        ReleaseComObject(cat.ActiveConnection)
        cat.ActiveConnection = Nothing
        cat = Nothing
        Return "Database Created Successfully"
    End Function

    Public Function FillDatabase(strSQLFile As String)
        Dim conn As New OleDbConnection(connString)
        conn.Open()
        Dim cmd As New OleDb.OleDbCommand("", conn)

        Dim strSQL As String = ""
        Using r As StreamReader = New StreamReader(strSQLFile)

            Dim line As String
            line = r.ReadLine

            Do While (Not line Is Nothing)
                If line.Trim <> "" And line.StartsWith("DROP") = False Then
                    strSQL &= line
                    If line.EndsWith(";") Then
                        cmd.CommandText = strSQL
                        cmd.ExecuteNonQuery()
                        strSQL = ""
                    End If

                End If
                line = r.ReadLine
            Loop
        End Using
        conn.Close()
        Return "SQL processed"
    End Function

    Public Function CheckDBVersion() As Long
        Dim strSQLFile As String

        If File.Exists(dbFile) = False Then
            CreateDB()
            strSQLFile = sqlPath & "db.sql"
            FillDatabase(strSQLFile)
        Else
            Dim dbVersion As Long = ReadDBVersion()
            'If dbVersion < 2 Then
            '    strSQLFile = sqlPath & "from1.sql"
            '    FillDatabase(strSQLFile)
            'End If
        End If

            Return ReadDBVersion()
    End Function

    Public Function ReadDBVersion() As Long
        Dim conn As New OleDbConnection(connString)
        conn.Open()

        Dim strSQL As String = "SELECT Version FROM tblVersion"
        Dim Version As Long

        Using comm As OleDbCommand = New OleDbCommand(strSQL, conn)

            Using reader As OleDbDataReader = comm.ExecuteReader()

                While reader.Read()
                    Version = reader.GetValue(0).ToString()
                End While
            End Using
        End Using
        conn.Close()
        Return Version
    End Function

    Private Function ReleaseComObject(ByVal objCom As Object) As Boolean
        Dim Result As Integer = 0
        For i As Integer = 0 To 9
            Result = Runtime.InteropServices.Marshal.FinalReleaseComObject(objCom)
            If Result = 0 Then
                Return True
            End If
            System.Threading.Thread.Sleep(0)
        Next
        Return False
    End Function
End Class
