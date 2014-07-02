MySQL_Connection_VBnet
======================

The purpose of this class is to make it simpler to use the MySQL connector in VB.net projects

Created by Jeremy Forsythe 2014
<jdforsythe@gmail.com>

Requirements:
-------------
This class is for VB.net projects and requires you to add a reference to the MySQL Connector (available at www.mysql.com)

Usage:
------
In order to use the MySQL_Connection, you must create an instance. The class implements iDisposable, so please use the "Using" syntax to ensure proper disposal of your connections. There are two overloads for the constructor - one takes the entire MySQL connection string, the other takes the server, username, password, databse, and options string separately. It is recommended to store this information in constants (possibly in a separate data module) to avoid retyping this information for every connection.

Example:

    Module DBConnect
      Private Const SQL_SERVER As String = "servername"
      Private Const SQL_USER As String = "username"
      Private Const SQL_PW As String = "password"
      Private Const SQL_DB As String = "database"
      Private Const SQL_OPTIONS As String = "Connection Timeout=60; DefaultcommandTimeout=5400; AllowZeroDateTime=true"
      Public Const SQL_CONNECTION_STRING As String = "Data Source=" & SQL_SERVER & "; user id=" & SQL_USER & _
                                                      "; password=" & SQL_PW & "; database=" & SQL_DB & "; " & SQL_OPTIONS
    End Module
    Using sql As New MySQL_Connection(SQL_CONNECTION_STRING)
    ...
    End Using
    
-or-

    Using sql As New MySQL_Connection("servername", "username", "password", "database", "options")
    ...
    End Using

When you get results back, they are in string, dictionary, or list format, so there is no need to keep the connection open after the results are returned.

There are two ways to set the query string for your queries. You can either set it as a property on the MySQL_Connection instance (i.e. sql.query = "") or pass it as a parameter to whichever query function you are running.

You are encouraged to use parameterized queries to prevent injection attacks and other problems. You can use the methods:
.addParam("@ParamName", "ParamValue")
.removeParam("@ParamName")
.clearParams()
to add, remove, and clear out all parameters for the instance. If you set no parameters, the query functions assume you have constructed your queries properly without them.


Examples:
1) Simple SELECT for a single string value (one value from one record)

    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      dim str as string = sql.selectQueryForSingleValue("SELECT Last_Name FROM Movie_Stars LIMIT 1")
      Messagebox.Show(str)
    End Using


    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "SELECT First_Name FROM Movie_Stars LIMIT 1"
      MessageBox.Show(sql.selectQueryForSingleValue())
    End Using
     
     
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "SELECT First_Name FROM Movie_Stars WHERE Last_Name=@LastName LIMIT 1"
      sql.addParam("@LastName", "Kidman")
      MessageBox.Show(sql.selectQueryForSingleValue())

2) Simple SELECT for a single RECORD (any number of columns from a single record) - returns a Dictionary(Of String, String)

    Dim record As New Dictionary(Of String, String)
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "SELECT * From Movie_Stars LIMIT 1"
      record = sql.selectQueryForSingleRecord()
    End Using
     
    For each pair In record
      MessageBox.Show(pair.Key & ": " & pair.Value)
    Next
     
     
    Dim record As New Dictionary(Of String, String)
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "SELECT * From Movie_Stars WHERE Last_Name=@LastName LIMIT 1"
      sql.addParam("@LastName", "Kidman")
      record = sql.selectQueryForSingleRecord()
    End Using
     
    For each pair In record
      MessageBox.Show(pair.Key & ": " & pair.Value)
    Next

3)  SELECT for a single COLUMN  (any number of records with a single column) - returns a List(Of String)

    Dim results As New List(Of String)
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "SELECT DISTINCT Last_Name From Movie_Stars"
      results = sql.selectQueryForSingleColumn()
    End Using
     
    For Each lastName As String In results
      MessageBox.Show("Last Name: " & lastName)
    Next
     
4) SELECT for multiple records - returns a List of records, each of which is a Dictionary(Of String, String)

    Dim results As New List(Of Dictionary(Of String, String)
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "SELECT * FROM Movie_Stars"
      results = sql.selectQuery()
    End Using
     
    For Each record as Dictionary(Of String, String) in results
      For Each pair in record
        MessageBox.Show(pair.Key & ": " & pair.Value)
      Next
    Next
     
     
    Dim results As New List(Of Dictionary(Of String, String)
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "SELECT * FROM Movie_Stars WHERE Last_Name=@LastName"
      sql.addParam("@LastName", "Smith")
      results = sql.selectQuery()
    End Using
     
    For Each record as Dictionary(Of String, String) in results
      For Each pair in record
        MessageBox.Show(pair.Key & ": " & pair.Value)
      Next
    Next
    
5) Insert, Update, Delete (these return the number of records affected)

    Dim deletions as Integer = 0
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "INSERT INTO Movie_Stars (First_Name, Last_Name) Values ('Nicole', 'Kidman')"
      deletions = sql.insertQuery()
    End Using
    MessageBox.Show("Inserted " & deletions.toString & " row(s)")


    Dim deletions as Integer = 0
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "INSERT INTO Movie_Stars (First_Name, Last_Name) Values (@FirstName, @LastName)"
      sql.addParam("@FirstName", "Nicole")
      sql.addParam("@LastName", "Kidman")
      deletions = sql.insertQuery()
    End Using
    MessageBox.Show("Inserted " & deletions.toString & " row(s)")
     
     
    Dim updates As Integer = 0
    Dim inserts As Integer = 0
    Using sql As New MySQL_Connection(MYSQL_CONNECTION_STRING)
      sql.query = "UPDATE Movie_Stars SET Last_Name=@NewLastName WHERE First_Name=@FirstName AND Last_Name=@OldLastName"
      sql.addParam("@OldLastName", "Kidman")
      sql.addParam("@FirstName", "Nicole")
      sql.addParam("@NewLastName", "Urban")
      updates = sql.updateQuery()
      
      sql.query = "INSERT INTO Movie_Stars (First_Name, Last_Name) Values (@FirstName, @LastName)"
      sql.removeParam("@OldLastName") '' yes, it's redundant - it's just an example!
      sql.clearParams()
      sql.addParam("@FirstName", "Tom")
      sql.addParam("@LastName", "Cruise")
      inserts = sql.insertQuery()
    End Using
    MessageBox.Show("Updated " & updates.toString & " row(s)")
    MessageBox.Show("Inserted " & inserts.toString & " row(s)")
    
    
If you have any questions, comments, etc. you know what to do. Feel free to submit pull requests and issues. I hope this helps speed up your SQL connection coding!


    


























