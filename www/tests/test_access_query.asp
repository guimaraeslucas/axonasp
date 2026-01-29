<%
' Simple Access Database Test
debug_asp_code = "TRUE"

response.write "<h1>Simple Access Database Query Test</h1>"

On Error Resume Next

' Create connection
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")

' Use sample.mdb from www/db folder - use absolute path for testing
Dim dbPath
' First try MapPath
dbPath = Server.MapPath("../db/sample.mdb")
response.write "<p>MapPath result: " & dbPath & "</p>"

' If path doesn't contain drive letter, prepend working directory
If InStr(dbPath, ":") = 0 Then
    dbPath = "E:\lucas\Desktop\Sites\LGGM-TCP\modules\image\ASP\axonasp\www\db\sample.mdb"
    response.write "<p>Using absolute path: " & dbPath & "</p>"
End If

response.write "<p>Final database path: " & dbPath & "</p>"

Dim connStr
' Try ACE OLEDB 12.0 first (more modern and likely to be installed)
connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
response.write "<p>Connection string: " & Server.HTMLEncode(connStr) & "</p>"

conn.ConnectionString = connStr
conn.Open

If Err.Number <> 0 Then
    response.write "<p><strong style='color:red'>Connection Error:</strong> " & Err.Description & "</p>"
    Response.End
End If

response.write "<p><strong style='color:green'>✓ Connected successfully!</strong> (State: " & conn.State & ")</p>"

' Execute a simple SELECT query on country table
response.write "<h2>Query Test: SELECT * FROM country</h2>"

Dim rs
Set rs = conn.Execute("SELECT * FROM country")

response.write "<p>After Execute - Checking rs...</p>"

If Err.Number <> 0 Then
    response.write "<p><strong style='color:red'>Query Error:</strong> " & Err.Description & " (Error #" & Err.Number & ")</p>"
    Err.Clear
End If

' Check if recordset is valid
response.write "<p>Is rs Nothing? " & (rs Is Nothing) & "</p>"

If rs Is Nothing Then
    response.write "<p><strong style='color:red'>Recordset is Nothing - No data returned</strong></p>"
    response.write "<p>Debugging info:</p>"
    response.write "<ul>"
    response.write "<li>Connection State: " & conn.State & "</li>"
    response.write "<li>Connection String: " & Server.HTMLEncode(conn.ConnectionString) & "</li>"
    response.write "</ul>"
Else
    response.write "<p><strong style='color:green'>✓ Recordset object received!</strong></p>"
    response.write "<p>TypeName(rs): " & TypeName(rs) & "</p>"
    
    ' Try to read properties
    On Error Resume Next
    Dim eofVal
    eofVal = rs.EOF
    If Err.Number <> 0 Then
        response.write "<p><strong style='color:red'>Error reading EOF:</strong> " & Err.Description & "</p>"
        Err.Clear
    Else
        response.write "<p>EOF: " & eofVal & "</p>"
    End If
    
    Dim bofVal
    bofVal = rs.BOF
    If Err.Number <> 0 Then
        response.write "<p><strong style='color:red'>Error reading BOF:</strong> " & Err.Description & "</p>"
        Err.Clear
    Else
        response.write "<p>BOF: " & bofVal & "</p>"
    End If
    
    On Error Goto 0
    
    ' Try to display records
    If Not rs.EOF Then
        response.write "<h3>Country Records:</h3>"
        response.write "<table border='1' cellpadding='5' cellspacing='0'>"
        response.write "<tr><th>iId</th><th>sText</th></tr>"
        
        ' Display data rows
        Dim rowCount
        rowCount = 0
        On Error Resume Next
        Do While Not rs.EOF And rowCount < 100
            Dim iId, sText
            
            ' Try to access fields
            iId = rs.Fields("iId")
            If Err.Number <> 0 Then
                response.write "<tr><td colspan='2'>Error accessing iId: " & Err.Description & "</td></tr>"
                Err.Clear
                Exit Do
            End If
            
            sText = rs.Fields("sText")
            If Err.Number <> 0 Then
                response.write "<tr><td>" & iId & "</td><td>Error accessing sText: " & Err.Description & "</td></tr>"
                Err.Clear
            Else
                response.write "<tr><td>" & iId & "</td><td>" & Server.HTMLEncode(sText) & "</td></tr>"
            End If
            
            rs.MoveNext
            If Err.Number <> 0 Then
                response.write "<tr><td colspan='2'>Error in MoveNext: " & Err.Description & "</td></tr>"
                Err.Clear
                Exit Do
            End If
            
            rowCount = rowCount + 1
        Loop
        On Error Goto 0
        
        response.write "</table>"
        response.write "<p><strong>Displayed " & rowCount & " record(s)</strong></p>"
    Else
        response.write "<p>No records found (EOF = True)</p>"
    End If
    
    On Error Resume Next
    rs.Close
    If Err.Number <> 0 Then
        response.write "<p>Error closing recordset: " & Err.Description & "</p>"
        Err.Clear
    End If
    On Error Goto 0
End If

conn.Close
Set rs = Nothing
Set conn = Nothing

response.write "<h3>Test Complete</h3>"
response.write "<p><a href='default.asp'>Back to Tests</a></p>"

On Error Goto 0
%>
