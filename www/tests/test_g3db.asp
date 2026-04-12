<%
Dim db
Set db = Server.CreateObject("G3DB")

' Verify the object was created and properties are accessible
If IsObject(db) Then
    Response.Write "G3DB object created OK" & vbCrLf
    Response.Write "IsOpen: " & db.IsOpen & vbCrLf
    Response.Write "Driver: '" & db.Driver & "'" & vbCrLf
    Response.Write "LastError: '" & db.LastError & "'" & vbCrLf

    ' Try to open a SQLite in-memory database
    Dim opened
    opened = db.Open("sqlite", ":memory:")
    Response.Write "Open(':memory:'): " & opened & vbCrLf
    Response.Write "IsOpen after open: " & db.IsOpen & vbCrLf

    If opened Then
        ' Exec — create table
        Dim res
        Set res = db.Exec("CREATE TABLE t(id INTEGER PRIMARY KEY, val TEXT)")
        Response.Write "CREATE TABLE result: " & (Not IsEmpty(res)) & vbCrLf

        ' Exec — insert rows
        db.Exec "INSERT INTO t(val) VALUES (?)", "hello"
        db.Exec "INSERT INTO t(val) VALUES (?)", "world"

        ' Query — SELECT
        Dim rs
        Set rs = db.Query("SELECT id, val FROM t")
        Response.Write "ResultSet EOF after open: " & rs.EOF & vbCrLf

        Dim rowCount
        rowCount = 0
        Do While Not rs.EOF
            rowCount = rowCount + 1
            Response.Write "  Row " & rowCount & ": id=" & rs("id") & " val=" & rs("val") & vbCrLf
            rs.MoveNext
        Loop
        rs.Close

        Response.Write "Total rows: " & rowCount & vbCrLf

        ' GetRows test
        Set rs = db.Query("SELECT id, val FROM t")
        Dim arr
        arr = rs.GetRows()
        Response.Write "GetRows cols: " & UBound(arr) + 1 & vbCrLf
        rs.Close

        ' Fields.Count test
        Set rs = db.Query("SELECT id, val FROM t")
        Response.Write "Fields.Count: " & rs.Fields.Count & vbCrLf
        rs.Close

        ' Stats
        Dim stats
        Set stats = db.Stats()
        If IsObject(stats) Then
            Response.Write "Stats.OpenConnections: " & stats("OpenConnections") & vbCrLf
        End If

        db.Close
        Response.Write "Closed: IsOpen=" & db.IsOpen & vbCrLf
    Else
        Response.Write "Open failed: " & db.LastError & vbCrLf
    End If
Else
    Response.Write "FAIL: G3DB object not created" & vbCrLf
End If
%>
