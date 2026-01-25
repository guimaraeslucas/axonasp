<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <title>G3pix AxonASP - ADODB Connection & Recordset Test</title>
    <meta charset="utf-8">
    <style>
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .section { border-left: 4px solid #667eea; padding: 15px; margin: 15px 0; background: #f9f9f9; border-radius: 4px; }
        .success { background: #d4edda; border-color: #c3e6cb; }
        .error { background: #f8d7da; border-color: #f5c6cb; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background: #667eea; color: white; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - ADODB Connection & Recordset Test</h1>
        <div class="intro">
            <p>Tests ADODB.Connection and ADODB.Recordset objects with SQL table creation, data insertion, navigation and query execution.</p>
        </div>

    <%
    ' Test ADODB.Connection
    Dim conn, rs, affected
    Set conn = Server.CreateObject("ADODB.Connection")
    
    ' Create in-memory SQLite database for testing
    conn.ConnectionString = "sqlite::memory:"
    conn.Open()
    %>
    
    <div class="section success">
        <h3>1. Connection Setup</h3>
        <%
        Response.Write "Connection State: " & conn.State & " (1=Open)<br>"
        Response.Write "Connection Mode: " & conn.Mode & "<br>"
        Response.Write "Connection String: " & conn.ConnectionString & "<br>"
        %>
    </div>

    <div class="section success">
        <h3>2. Create Test Table and Insert Data</h3>
        <%
        ' Create table
        affected = conn.Execute("CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT, email TEXT, age INTEGER)")
        Response.Write "Table created<br>"
        sqlInsert1 =  "INSERT INTO users (id, name, email, age) VALUES (1, 'John Doe', 'john@example.com', 30)"
        ' Insert rows
        conn.Execute(sqlInsert1)
        conn.Execute("INSERT INTO users (id, name, email, age) VALUES (2, 'Jane Smith', 'jane@example.com', 25)")
        conn.Execute("INSERT INTO users (id, name, email, age) VALUES (3, 'Bob Johnson', 'bob@example.com', 35)")
        Response.Write "Data inserted (3 rows)<br>"
        %>
    </div>

    <div class="section success">
        <h3>3. Recordset Navigation</h3>
        <%
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open "SELECT * FROM users", conn
        
        Response.Write "Recordset State: " & rs.State & " (1=Open)<br>"
        Response.Write "Record Count: " & rs.RecordCount & "<br>"
        Response.Write "EOF on open: " & rs.EOF & "<br>"
        Response.Write "BOF on open: " & rs.BOF & "<br>"
        %>
    </div>

    <div class="section success">
        <h3>4. Display Records</h3>
        <table>
            <tr>
                <th>ID</th>
                <th>Name</th>
                <th>Email</th>
                <th>Age</th>
            </tr>
            <%
            rs.MoveFirst()
            Do While Not rs.EOF
                Response.Write "<tr>"
                Response.Write "<td>" & rs("id") & "</td>"
                Response.Write "<td>" & rs("name") & "</td>"
                Response.Write "<td>" & rs("email") & "</td>"
                Response.Write "<td>" & rs("age") & "</td>"
                Response.Write "</tr>"
                rs.MoveNext()
            Loop
            %>
        </table>
    </div>

    <div class="section success">
        <h3>5. Navigate Record by Record</h3>
        <%
        rs.MoveFirst()
        Response.Write "First record - Name: " & rs("name") & "<br>"
        
        rs.MoveNext()
        Response.Write "Second record - Name: " & rs("name") & "<br>"
        
        rs.MoveLast()
        Response.Write "Last record - Name: " & rs("name") & "<br>"
        
        rs.MovePrevious()
        Response.Write "Previous from last - Name: " & rs("name") & "<br>"
        %>
    </div>

    <div class="section success">
        <h3>6. Cleanup</h3>
        <%
        rs.Close
        conn.Close
        Response.Write "Recordset and Connection closed<br>"
        Response.Write "Connection State after close: " & conn.State & " (0=Closed)<br>"
        %>
    </div>

    <hr>
    <p><small>ADODB Recordset implementation supports SQLite (in-memory and file), MySQL, and PostgreSQL via connection string parsing.</small></p>
    
    <h3>Connection String Examples:</h3>
    <pre>
' SQLite (in-memory)
conn.ConnectionString = "sqlite::memory:"

' SQLite (file)
conn.ConnectionString = "sqlite:./mydata.db"

' MySQL
conn.ConnectionString = "Driver={MySQL ODBC Driver};Server=localhost;Database=mydb;UID=root;PWD=password"

' PostgreSQL
conn.ConnectionString = "Driver={PostgreSQL ODBC Driver};Server=localhost;Database=mydb;UID=postgres;PWD=password;Port=5432"

' MS SQL Server
conn.ConnectionString = "Driver={ODBC Driver 17 for SQL Server};Server=localhost;Database=mydb;UID=sa;PWD=password;Port=1433"
    </pre>
</body>
</html>

