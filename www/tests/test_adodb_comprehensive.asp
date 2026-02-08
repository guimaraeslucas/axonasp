<%
' Comprehensive ADODB.Connection and ADODB.Recordset Test
' Testing: SQLite in-memory database with full CRUD operations

Response.Write("<h1>ADODB Comprehensive Test</h1>")
Response.Write("<hr/>")

' --- Test 1: Create connection to in-memory SQLite database ---
Response.Write("<h2>Test 1: Connection to SQLite In-Memory Database</h2>")
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = "sqlite::memory:"
conn.Open()

If conn.State = 1 Then
    Response.Write("<p style='color:green;'>Connection opened successfully. State = " & conn.State & "</p>")
Else
    Response.Write("<p style='color:red;'>Failed to open connection</p>")
    Response.End()
End If

' --- Test 2: Create table and insert data ---
Response.Write("<h2>Test 2: Create Table and Insert Data</h2>")

' Create table
sqlCreateTable = "CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT, email TEXT, age INTEGER)"
result = conn.Execute(sqlCreateTable)
Response.Write("<p>Table created successfully</p>")

' Insert data
sqlInsert1 = "INSERT INTO users (name, email, age) VALUES ('Alice Johnson', 'alice@example.com', 30)"
sqlInsert2 = "INSERT INTO users (name, email, age) VALUES ('Bob Smith', 'bob@example.com', 25)"
sqlInsert3 = "INSERT INTO users (name, email, age) VALUES ('Carol White', 'carol@example.com', 35)"

conn.Execute(sqlInsert1)
conn.Execute(sqlInsert2)
conn.Execute(sqlInsert3)

Response.Write("<p>3 records inserted successfully</p>")

' --- Test 3: Open and read recordset ---
Response.Write("<h2>Test 3: Open Recordset and Read All Records</h2>")

Set rs = Server.CreateObject("ADODB.Recordset")
sqlSelect = "SELECT id, name, email, age FROM users ORDER BY id"
rs.Open sqlSelect, conn

Response.Write("<table border='1' cellpadding='5'>")
Response.Write("<tr><th>ID</th><th>Name</th><th>Email</th><th>Age</th></tr>")

' Loop through all records
While Not rs.EOF
    id = rs("id")
    name = rs("name")
    email = rs("email")
    age = rs("age")
    
    Response.Write("<tr>")
    Response.Write("<td>" & id & "</td>")
    Response.Write("<td>" & name & "</td>")
    Response.Write("<td>" & email & "</td>")
    Response.Write("<td>" & age & "</td>")
    Response.Write("</tr>")
    
    rs.MoveNext()
Wend

Response.Write("</table>")
Response.Write("<p>Total records: " & rs.RecordCount & "</p>")

' --- Test 3b: Connection.Execute returns a recordset ---
Response.Write("<h2>Test 3b: Connection.Execute Recordset</h2>")
Dim rsExec
Set rsExec = conn.Execute("SELECT name FROM users ORDER BY id")

If IsObject(rsExec) Then
    Response.Write("<p style='color:green;'>Execute returned a recordset.</p>")
    If Not rsExec.EOF Then
        Response.Write("<p>First name from Execute: " & rsExec("name") & "</p>")
    End If
    rsExec.Close
Else
    Response.Write("<p style='color:red;'>Execute did not return a recordset.</p>")
End If

' --- Test 4: Test navigation methods ---
Response.Write("<h2>Test 4: Recordset Navigation Methods</h2>")

rs.MoveFirst()
Response.Write("<p>After MoveFirst(): BOF=" & rs.BOF & ", EOF=" & rs.EOF & ", Current ID=" & rs("id") & "</p>")

rs.MoveLast()
Response.Write("<p>After MoveLast(): Current ID=" & rs("id") & ", EOF=" & rs.EOF & "</p>")

rs.MovePrevious()
Response.Write("<p>After MovePrevious(): Current ID=" & rs("id") & "</p>")

rs.MoveFirst()

' --- Test 5: Test with SQL WHERE clause ---
Response.Write("<h2>Test 5: Query with WHERE Clause</h2>")

rs.Close()

Set rs2 = Server.CreateObject("ADODB.Recordset")
sqlWhere = "SELECT name, age FROM users WHERE age >= 30"
rs2.Open sqlWhere, conn

Response.Write("<p>Records with age >= 30:</p>")
Response.Write("<ul>")
While Not rs2.EOF
    Response.Write("<li>" & rs2("name") & " (Age: " & rs2("age") & ")</li>")
    rs2.MoveNext()
Wend
Response.Write("</ul>")

rs2.Close()

' --- Test 6: Test Fields collection ---
Response.Write("<h2>Test 6: Fields Collection</h2>")

Set rs3 = Server.CreateObject("ADODB.Recordset")
rs3.Open "SELECT id, name, email, age FROM users LIMIT 1", conn

Response.Write("<p>First record fields:</p>")
Response.Write("<ul>")
Response.Write("<li>id = " & rs3.Fields.Item("id") & "</li>")
Response.Write("<li>name = " & rs3.Fields.Item("name") & "</li>")
Response.Write("<li>email = " & rs3.Fields.Item("email") & "</li>")
Response.Write("<li>age = " & rs3.Fields.Item("age") & "</li>")
Response.Write("</ul>")

rs3.Close()

' --- Test 6b: Update and Delete ---
Response.Write("<h2>Test 6b: Update and Delete</h2>")

Dim rsUpdate
Set rsUpdate = Server.CreateObject("ADODB.Recordset")
rsUpdate.Open "SELECT id, name, email, age FROM users WHERE id = 1", conn
If Not rsUpdate.EOF Then
    rsUpdate("age") = 31
    rsUpdate.Update
    Response.Write("<p>Updated user id=1 age to 31.</p>")
End If
rsUpdate.Close

Dim rsDelete
Set rsDelete = Server.CreateObject("ADODB.Recordset")
rsDelete.Open "SELECT id FROM users WHERE id = 2", conn
If Not rsDelete.EOF Then
    rsDelete.Delete
    Response.Write("<p>Deleted user id=2.</p>")
End If
rsDelete.Close

Dim rsCount
Set rsCount = conn.Execute("SELECT COUNT(*) AS total FROM users")
If IsObject(rsCount) Then
    Response.Write("<p>Remaining records after delete: " & rsCount("total") & "</p>")
    rsCount.Close
End If

' --- Test 7: Test connection properties ---
Response.Write("<h2>Test 7: Connection Properties</h2>")

Response.Write("<p>Connection String: " & conn.ConnectionString & "</p>")
Response.Write("<p>Connection State: " & conn.State & " (1=Open)</p>")
Response.Write("<p>Connection Mode: " & conn.Mode & "</p>")

' --- Test 8: Close connection ---
Response.Write("<h2>Test 8: Close Connection</h2>")

rs.Close()
conn.Close()

If conn.State = 0 Then
    Response.Write("<p style='color:green;'>Connection closed successfully. State = " & conn.State & "</p>")
Else
    Response.Write("<p style='color:red;'>Connection still open</p>")
End If

Response.Write("<hr/>")
Response.Write("<p>All ADODB tests completed!</p>")
%>
