Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Net.Mime.MediaTypeNames
Imports ADOX

Public Class ClsDatabase
    Private connString As String
    Private dbFile As String

    Private sqlPath As String = IIf(AppDomain.CurrentDomain.BaseDirectory.Contains("Release"), AppDomain.CurrentDomain.BaseDirectory.Replace("\inoGen\bin\Release\net9.0-windows7.0\", ""), AppDomain.CurrentDomain.BaseDirectory.Replace("\inoGen\bin\Debug\net9.0-windows7.0\", "")) & "\inoGenDLL\SQL\"

    Private currentVersion As Long = 12

    Public Sub New(dbFileString As String)
        connString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=""{0}"";Persist Security Info=True", dbFileString)
        dbFile = dbFileString
        If AppDomain.CurrentDomain.BaseDirectory.Contains("TestInoGen") Then
            If AppDomain.CurrentDomain.BaseDirectory.Contains("Release") Then
                sqlPath = AppDomain.CurrentDomain.BaseDirectory.Replace("\TestInoGen\bin\Release\net9.0-windows", "") & "\inoGenDLL\SQL\"
            Else
                sqlPath = AppDomain.CurrentDomain.BaseDirectory.Replace("\TestInoGen\bin\Debug\net9.0-windows", "") & "\inoGenDLL\SQL\"
            End If

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

    Public Function FillDatabase(strSQLFile As String) As String

        Using conn As New OleDbConnection(connString)

            conn.Open()

            Using cmd As New OleDbCommand("", conn)

                Using r As New StreamReader(strSQLFile)

                    Dim strSQL As String = ""
                    Dim line As String = r.ReadLine()

                    Do While line IsNot Nothing

                        If line.Trim <> "" AndAlso
                       Not line.StartsWith("DROP") Then

                            strSQL &= line

                            If line.EndsWith(";") Then

                                cmd.CommandText = strSQL
                                cmd.ExecuteNonQuery()

                                strSQL = ""

                            End If

                        End If

                        line = r.ReadLine()

                    Loop

                End Using

            End Using

        End Using

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
            Dim cGDB As New inoGenDLL.ClsGenDB(dbFile)
            For updateVersion = 2 To currentVersion
                If dbVersion < updateVersion Then
                    strSQLFile = sqlPath & "from_" & (updateVersion - 1).ToString() & ".sql"
                    FillDatabase(strSQLFile)
                    If updateVersion = 12 Then
                        cGDB.FillVornamenPhonetic()
                        cGDB.FillNachnamenPhonetic()
                    End If
                End If
            Next
        End If

        Return ReadDBVersion()
    End Function

    Public Function ReadDBVersion() As Long

        Dim strSQL As String = "SELECT Version FROM tblVersion"

        Using conn As New OleDbConnection(connString)

            conn.Open()

            Using comm As New OleDbCommand(strSQL, conn)

                Using reader As OleDbDataReader = comm.ExecuteReader()

                    If reader.Read() Then
                        Return Convert.ToInt64(reader.GetValue(0))
                    End If

                End Using
            End Using

        End Using

        Return 0

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
