Imports MySql.Data.MySqlClient
Imports MySql.Data

Public Class MySQL_Connection
    Implements IDisposable
    Private isDisposed As Boolean = False

    Private srv As String = ""
    Private user As String = ""
    Private pw As String = ""
    Private db As String = ""
    Private dboptions As String = ""
    Private quer As String = ""
    Private prms As New Dictionary(Of String, String)
    Private conn As MySqlConnection
    Private comm As MySqlCommand
    Private data As MySqlDataAdapter

    Property database() As String
        Get
            Return db
        End Get

        Set(value As String)
            db = value
        End Set
    End Property
    Property server() As String
        Get
            Return srv
        End Get

        Set(value As String)
            srv = value
        End Set
    End Property
    Property username() As String
        Get
            Return user
        End Get

        Set(value As String)
            user = value
        End Set
    End Property
    Property password() As String
        Get
            Return pw
        End Get

        Set(value As String)
            pw = value
        End Set
    End Property
    Property options() As String
        Get
            Return dboptions
        End Get

        Set(value As String)
            dboptions = value
        End Set
    End Property
    Property query() As String
        Get
            Return quer
        End Get
        Set(value As String)
            quer = value
        End Set
    End Property


    Sub New(ByVal serverAddress As String, ByVal userName As String, ByVal password As String, ByVal database As String, Optional ByVal options As String = "")
        srv = serverAddress
        user = userName
        pw = password
        db = database
        If (options <> "") Then
            dboptions = options
        End If

        If (isDisposed = True) Then Throw New ObjectDisposedException("Cannot open - connection has been disposed")

        If Not ((srv = "") OrElse (user = "") OrElse (pw = "") OrElse (db = "")) Then
            Try
                Dim str As String = "Data Source=" & srv & "; user id=" & user & "; password=" & pw & "; database=" & db
                If Not (dboptions = "") Then
                    str = str & "; " & dboptions
                End If
                conn = New MySqlConnection(str)
                conn.Open()
            Catch ex As Exception
                Throw ex
            End Try
        Else
            Throw New Exception("Invalid connection parameters for database")
        End If
    End Sub

    Sub New(connectionString As String)
        If (isDisposed = True) Then Throw New ObjectDisposedException("Cannot open - connection has been disposed")

        If Not (connectionString = "") Then
            Try
                conn = New MySqlConnection(connectionString)
                conn.Open()
            Catch ex As Exception
                Throw ex
            End Try
        Else
            Throw New Exception("Invalid connection string for database")
        End If
    End Sub


    Public Sub addParam(ByVal key As String, ByVal value As String)
        prms.Add(key, value)
    End Sub

    Public Sub removeParam(ByVal key As String)
        prms.Remove(key)
    End Sub

    Public Sub clearParams()
        prms.Clear()
    End Sub

    '' a select query where you expect only one string returned
    Public Function selectQueryForSingleValue(Optional ByVal query As String = "") As String
        If (isDisposed = True) Then Throw New ObjectDisposedException("Cannot execute query - connection has been disposed")
        If (query = "") Then query = quer
        If (query = "") Then Throw New Exception("Cannot execute empty query")

        comm = New MySqlCommand(query, conn)
        '' if parameters were set, loop through and set them in the command
        If (prms.Count > 0) Then
            For Each pair In prms
                comm.Parameters.AddWithValue(pair.Key, pair.Value)
            Next
        End If

        Return Convert.ToString(comm.ExecuteScalar())
    End Function

    '' a select query where you expect only one record returned - returns a dictionary(of string, string)
    Public Function selectQueryForSingleRecord(Optional ByVal query As String = "") As Dictionary(Of String, String)
        If (isDisposed = True) Then Throw New ObjectDisposedException("Cannot execute query - connection has been disposed")
        If (query = "") Then query = quer
        If (query = "") Then Throw New Exception("Cannot execute empty query")

        comm = New MySqlCommand(query, conn)
        '' if parameters were set, loop through and set them in the command
        If (prms.Count > 0) Then
            For Each pair In prms
                comm.Parameters.AddWithValue(pair.Key, pair.Value)
            Next
        End If

        Dim record As New Dictionary(Of String, String)
        Dim reader As MySqlDataReader = comm.ExecuteReader()
        '' if there are rows
        If reader.HasRows Then
            reader.Read()
            '' get the keys/values for the first record and add them to the dictionary
            For i As Integer = 0 To (reader.FieldCount - 1)
                record.Add(reader.GetName(i), reader(i).ToString)
            Next
        End If
        reader.Close()
        comm.Dispose()

        Return record
    End Function
    
    '' a select query where you expect only one column returned for any number of records - returns a List(of String)
    Public Function selectQueryForSingleColumn(Optional ByVal query As String = "") As List(Of String)
        If (isDisposed = True) Then Throw New ObjectDisposedException("Cannot execute query - connection has been disposed")
        If (query = "") Then query = quer
        If (query = "") Then Throw New Exception("Cannot execute empty query")

        comm = New MySqlCommand(query, conn)
        '' if parameters were set, loop through and set them in the command
        If (prms.Count > 0) Then
            For Each pair In prms
                comm.Parameters.AddWithValue(pair.Key, pair.Value)
            Next
        End If

        Dim allRecords As New List(Of String)
        Dim reader As MySqlDataReader = comm.ExecuteReader()

        If reader.HasRows Then
            While reader.Read()
                allRecords.Add(reader(0).ToString)
            End While
        End If
        reader.Close()
        comm.Dispose()
        Return allRecords
    End Function

    '' a select query where you expect (the possibility of) multiple columns and/or multiple records
    '' returns a List(Of Dictionary(Of String, String)) of all the records
    '' where each List.item is a record and each Dictionary is (key, value) representing each column in the record
    Public Function selectQuery(Optional ByVal query As String = "") As List(Of Dictionary(Of String, String))
        If (isDisposed = True) Then Throw New ObjectDisposedException("Cannot execute query - connection has been disposed")
        If (query = "") Then query = quer
        If (query = "") Then Throw New Exception("Cannot execute empty query")

        comm = New MySqlCommand(query, conn)
        '' if parameters were set, loop through and set them in the command
        If (prms.Count > 0) Then
            For Each pair In prms
                comm.Parameters.AddWithValue(pair.Key, pair.Value)
            Next
        End If

        Dim allRecords As New List(Of Dictionary(Of String, String))
        Dim record As Dictionary(Of String, String)
        Dim reader As MySqlDataReader = comm.ExecuteReader()

        If reader.HasRows Then
            While reader.Read()
                '' create a new dictionary for the record, adding each column and value
                record = New Dictionary(Of String, String)
                For i As Integer = 0 To (reader.FieldCount - 1)
                    record.Add(reader.GetName(i), reader(i).ToString)
                Next
                '' add the dictionary to the List
                allRecords.Add(record)
            End While
        End If

        reader.Close()
        comm.Dispose()
        Return allRecords
    End Function

    '' an insert query, returns the number of rows affected
    Public Function insertQuery(Optional ByVal query As String = "") As Integer
        If (isDisposed = True) Then Throw New ObjectDisposedException("Cannot execute query - connection has been disposed")
        If (query = "") Then query = quer
        If (query = "") Then Throw New Exception("Cannot execute empty query")

        comm = New MySqlCommand(query, conn)
        '' if parameters were set, loop through and set them in the command
        If (prms.Count > 0) Then
            For Each pair In prms
                comm.Parameters.AddWithValue(pair.Key, pair.Value)
            Next
        End If
        Return comm.ExecuteNonQuery()
    End Function

    '' updating is essentially the same as inserting - returns the number of rows affected
    Public Function updateQuery(Optional ByVal query As String = "") As Integer
        Return insertQuery(query)
    End Function

    '' deleting is essentially the same as inserting - returns the number of rows affected
    Public Function deleteQuery(Optional ByVal query As String = "") As Integer
        Return insertQuery(query)
    End Function


    '' implementing IDisposable
    'Public Sub Dispose() Implements System.IDisposable.Dispose
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overridable Overloads Sub Dispose(disposing As Boolean)
        '' if already disposed, don't do anything
        If Not (isDisposed) Then
            If (disposing) Then
                conn.Close()
                conn = Nothing

                comm.Dispose()
                comm = Nothing

                srv = Nothing
                user = Nothing
                pw = Nothing
                db = Nothing
                dboptions = Nothing
                data = Nothing

                isDisposed = True
            End If
        End If

    End Sub

    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub

    Public Sub closeConnection()
        If (isDisposed = True) Then Throw New ObjectDisposedException("Cannot close - connection has been disposed")
        Dispose()
    End Sub

End Class
