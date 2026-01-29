<%
' Test Access Database Connectivity
' This test demonstrates Access database support

debug_asp_code = "TRUE"

response.write "<h1>Access Database Connection Test</h1>"
response.write "<p>This test attempts to connect to Access databases on Windows systems.</p>"

' Test 1: Jet OLEDB 4.0 connection (older Access format)
response.write "<h2>Test 1: Microsoft Jet OLEDB 4.0</h2>"

Dim conn1
Set conn1 = Server.CreateObject("ADODB.Connection")

' Using absolute path resolved via Server.MapPath
Dim dbPath1
dbPath1 = Server.MapPath("../db/data.mdb")
response.write "<p>Database path: " & dbPath1 & "</p>"

Dim connStr1
connStr1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath1
response.write "<p>Connection string: " & Server.HTMLEncode(connStr1) & "</p>"

On Error Resume Next

conn1.ConnectionString = connStr1
conn1.Open

Dim errMsg
errMsg = ""
If Err.Number <> 0 Then
    errMsg = Err.Description
    response.write "<p><strong>Error during Open:</strong> " & errMsg & "</p>"
End If

If conn1.State = 1 Then
    response.write "<p><strong style='color:green'>✓ Success!</strong> Connection opened successfully.</p>"
    
    ' Try to execute a simple query
    Dim rs1
    Set rs1 = Server.CreateObject("ADODB.Recordset")
    
    If Err.Number = 0 Then
        Set rs1.ActiveConnection = conn1
        rs1.Open "SELECT COUNT(*) AS RecordCount FROM Table1", conn1, 1, 1
        
        If Err.Number = 0 Then
            response.write "<p>Query executed successfully.</p>"
            If Not rs1.EOF Then
                response.write "<p>Record count in Table1: " & rs1("RecordCount") & "</p>"
            End If
            rs1.Close
        Else
            response.write "<p>Query execution info: " & Err.Description & "</p>"
        End If
    End If
    
    conn1.Close
Else
    response.write "<p><strong style='color:red'>✗ Connection failed</strong> (State: " & conn1.State & ")</p>"
    response.write "<p>This is expected if:</p>"
    response.write "<ul>"
    response.write "<li>The database file doesn't exist at the specified path</li>"
    response.write "<li>Running on a non-Windows platform</li>"
    response.write "<li>OLEDB drivers are not installed</li>"
    response.write "<li>The file doesn't have proper read permissions</li>"
    response.write "</ul>"
End If

Err.Clear

Set rs1 = Nothing
Set conn1 = Nothing

' Test 2: ACE OLEDB 12.0 connection (newer Access format)
response.write "<h2>Test 2: Microsoft ACE OLEDB 12.0</h2>"

Dim conn2
Set conn2 = Server.CreateObject("ADODB.Connection")

Dim dbPath2
dbPath2 = Server.MapPath("../db/data.mdb")
response.write "<p>Database path: " & dbPath2 & "</p>"

Dim connStr2
connStr2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath2
response.write "<p>Connection string: " & Server.HTMLEncode(connStr2) & "</p>"

conn2.ConnectionString = connStr2
conn2.Open

If Err.Number <> 0 Then
    response.write "<p><strong>Error during Open:</strong> " & Err.Description & "</p>"
End If

If conn2.State = 1 Then
    response.write "<p><strong style='color:green'>✓ Success!</strong> Connection opened successfully.</p>"
    conn2.Close
Else
    response.write "<p><strong style='color:red'>✗ Connection failed</strong> (State: " & conn2.State & ")</p>"
    response.write "<p>This is expected if the database file doesn't exist or ACE OLEDB driver is not installed.</p>"
End If

Err.Clear

Set conn2 = Nothing

response.write "<h2>Test Complete</h2>"
response.write "<p><a href='default.asp'>Back to Tests</a></p>"

On Error Goto 0
%>
