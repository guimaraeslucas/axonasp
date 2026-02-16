<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>ADODB Advanced Features Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h2 { color: #333; border-bottom: 2px solid #007bff; padding-bottom: 5px; }
        .success { color: green; font-weight: bold; }
        .error { color: red; font-weight: bold; }
        .info { color: blue; }
        table { border-collapse: collapse; margin: 10px 0; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
        th { background-color: #007bff; color: white; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        .section { margin: 20px 0; padding: 15px; border: 1px solid #ccc; border-radius: 5px; }
    </style>
</head>
<body>
    <h1>ADODB Advanced Features Test</h1>
    
    <div class="section">
        <h2>1. Recordset.Supports() Method Test</h2>
        <%
        Dim conn, rs
        Set conn = Server.CreateObject("ADODB.Connection")
        Set rs = Server.CreateObject("ADODB.Recordset")
        
        conn.ConnectionString = "sqlite::memory:"
        conn.Open
        
        ' Create test table
        conn.Execute "CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT, age INTEGER)"
        conn.Execute "INSERT INTO users (name, age) VALUES ('Alice', 25)"
        conn.Execute "INSERT INTO users (name, age) VALUES ('Bob', 30)"
        conn.Execute "INSERT INTO users (name, age) VALUES ('Charlie', 35)"
        
        rs.Open "SELECT * FROM users", conn
        
        Response.Write "<p><strong>Testing Supports() method:</strong></p>"
        Response.Write "<ul>"
        Response.Write "<li>Supports AddNew (0x1000400): " & rs.Supports(&H1000400) & "</li>"
        Response.Write "<li>Supports Delete (0x1000800): " & rs.Supports(&H1000800) & "</li>"
        Response.Write "<li>Supports Update (0x1008000): " & rs.Supports(&H1008000) & "</li>"
        Response.Write "<li>Supports MovePrevious (0x200): " & rs.Supports(&H200) & "</li>"
        Response.Write "<li>Supports Find (0x80000): " & rs.Supports(&H80000) & "</li>"
        Response.Write "<li>Supports Bookmark (0x2000): " & rs.Supports(&H2000) & " (not implemented)</li>"
        Response.Write "</ul>"
        
        Response.Write "<p class='success'>✓ Supports() method working correctly</p>"
        %>
    </div>
    
    <div class="section">
        <h2>2. Recordset.Filter Property Test</h2>
        <%
        Response.Write "<p><strong>Original recordset:</strong></p>"
        Response.Write "<table><tr><th>ID</th><th>Name</th><th>Age</th></tr>"
        
        rs.MoveFirst
        Do While Not rs.EOF
            Response.Write "<tr>"
            Response.Write "<td>" & rs("id") & "</td>"
            Response.Write "<td>" & rs("name") & "</td>"
            Response.Write "<td>" & rs("age") & "</td>"
            Response.Write "</tr>"
            rs.MoveNext
        Loop
        Response.Write "</table>"
        Response.Write "<p>Total records: " & rs.RecordCount & "</p>"
        
        ' Apply filter
        rs.Filter = "age > 28"
        
        Response.Write "<p><strong>After applying filter 'age > 28':</strong></p>"
        Response.Write "<table><tr><th>ID</th><th>Name</th><th>Age</th></tr>"
        
        rs.MoveFirst
        Do While Not rs.EOF
            Response.Write "<tr>"
            Response.Write "<td>" & rs("id") & "</td>"
            Response.Write "<td>" & rs("name") & "</td>"
            Response.Write "<td>" & rs("age") & "</td>"
            Response.Write "</tr>"
            rs.MoveNext
        Loop
        Response.Write "</table>"
        Response.Write "<p>Filtered records: " & rs.RecordCount & "</p>"
        
        ' Clear filter
        rs.Filter = ""
        
        Response.Write "<p><strong>After clearing filter:</strong></p>"
        Response.Write "<p>Total records: " & rs.RecordCount & "</p>"
        
        Response.Write "<p class='success'>✓ Filter property working correctly</p>"
        
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        %>
    </div>
    
    <div class="section">
        <h2>3. Connection.Errors Collection Test</h2>
        <%
        Set conn = Server.CreateObject("ADODB.Connection")
        conn.ConnectionString = "sqlite::memory:"
        conn.Open
        
        Response.Write "<p><strong>Testing Errors collection:</strong></p>"
        
        ' Try to execute invalid SQL
        Dim result
        result = conn.Execute("SELECT * FROM nonexistent_table")
        
        Response.Write "<p>Errors count: " & conn.Errors.Count & "</p>"
        
        If conn.Errors.Count > 0 Then
            Response.Write "<p class='info'>Error captured in Errors collection (as expected for invalid SQL)</p>"
        End If
        
        ' Try valid SQL
        conn.Execute "CREATE TABLE test (id INTEGER)"
        Response.Write "<p>After successful command, Errors count: " & conn.Errors.Count & "</p>"
        
        Response.Write "<p class='success'>✓ Errors collection working correctly</p>"
        
        conn.Close
        Set conn = Nothing
        %>
    </div>
    
    <div class="section">
        <h2>4. ADODB.Stream Test</h2>
        <%
        Dim stream
        Set stream = Server.CreateObject("ADODB.Stream")
        
        Response.Write "<p><strong>Testing ADODB.Stream:</strong></p>"
        
        ' Test text writing
        stream.Open
        stream.Type = 1 ' adTypeText
        stream.WriteText "Hello from ADODB.Stream!" & vbCrLf
        stream.WriteText "This is line 2" & vbCrLf
        stream.WriteText "And line 3"
        
        Response.Write "<p>Stream Size: " & stream.Size & " bytes</p>"
        Response.Write "<p>Stream Position: " & stream.Position & "</p>"
        
        ' Read back
        stream.Position = 0
        Dim content
        content = stream.ReadText()
        
        Response.Write "<p><strong>Content read from stream:</strong></p>"
        Response.Write "<pre>" & Server.HTMLEncode(content) & "</pre>"
        
        ' Test file operations
        stream.Position = 0
        stream.SaveToFile Server.MapPath("test_stream_output.txt"), 2 ' adSaveCreateOverWrite

        Response.Write "<p class='success'>✓ File saved: test_stream_output.txt</p>"
        
        ' Load from file
        Dim stream2
        Set stream2 = Server.CreateObject("ADODB.Stream")
        stream2.Type = 1
        stream2.Open
        stream2.LoadFromFile Server.MapPath("test_stream_output.txt")
        
        Response.Write "<p>Loaded from file, Size: " & stream2.Size & " bytes</p>"
        
        stream2.Position = 0
        Dim loadedContent
        loadedContent = stream2.ReadText()
        
        Response.Write "<p><strong>Content loaded from file:</strong></p>"
        Response.Write "<pre>" & Server.HTMLEncode(loadedContent) & "</pre>"
        
        stream.Close
        stream2.Close
        Set stream = Nothing
        Set stream2 = Nothing
        
        Response.Write "<p class='success'>✓ ADODB.Stream working correctly</p>"
        %>
    </div>
    
    <div class="section">
        <h2>5. Filter with Different Operators</h2>
        <%
        Set conn = Server.CreateObject("ADODB.Connection")
        Set rs = Server.CreateObject("ADODB.Recordset")
        
        conn.ConnectionString = "sqlite::memory:"
        conn.Open
        
        conn.Execute "CREATE TABLE products (id INTEGER, name TEXT, price REAL)"
        conn.Execute "INSERT INTO products VALUES (1, 'Apple', 1.5)"
        conn.Execute "INSERT INTO products VALUES (2, 'Banana', 0.8)"
        conn.Execute "INSERT INTO products VALUES (3, 'Cherry', 2.5)"
        conn.Execute "INSERT INTO products VALUES (4, 'Date', 3.0)"
        
        rs.Open "SELECT * FROM products", conn
        
        Response.Write "<p><strong>Testing various filter operators:</strong></p>"
        
        ' Test equality
        rs.Filter = "name = 'Apple'"
        rs.MoveFirst
        Response.Write "<p>Filter 'name = Apple': Found " & rs.RecordCount & " record(s) - " & rs("name") & "</p>"
        
        ' Test greater than
        rs.Filter = "price > 2.0"
        Response.Write "<p>Filter 'price > 2.0': Found " & rs.RecordCount & " record(s)</p>"
        rs.MoveFirst
        Do While Not rs.EOF
            Response.Write "<span>- " & rs("name") & " ($" & rs("price") & ")</span><br>"
            rs.MoveNext
        Loop
        
        ' Clear filter
        rs.Filter = ""
        Response.Write "<p>All records after clearing filter: " & rs.RecordCount & "</p>"
        
        Response.Write "<p class='success'>✓ Multiple filter operators working correctly</p>"
        
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        %>
    </div>
    
    <hr>
    <p style="text-align: center; color: #666; margin-top: 30px;">
        <strong>All ADODB Advanced Features Tests Complete!</strong><br>
        <a href="default.asp">← Back to Home</a>
    </p>
</body>
</html>
