<%
' Test database with actual data
Response.Write "<h1>Testing Database Recordset</h1>"

' Create connection
dim conn : set conn = Server.CreateObject("ADODB.Connection")
Response.Write "Connection created<br>"

conn.Open "Driver={SQLite3};DBQ=" & Server.MapPath("../db/sample.mdb")
Response.Write "Connection opened<br>"

' Create and open recordset
dim rs : set rs = conn.Execute("SELECT * FROM testdata LIMIT 5")
Response.Write "Query executed<br>"

if rs is nothing then
    Response.Write "rs is NOTHING after execute!<br>"
else
    Response.Write "rs.EOF: " & rs.EOF & "<br>"
    Response.Write "rs.RecordCount: " & rs.RecordCount & "<br>"
    Response.Write "Fields.Count: " & rs.Fields.Count & "<br>"
    
    if not rs.EOF then
        Response.Write "<h2>Field Names:</h2><ul>"
        dim field
        for each field in rs.Fields
            Response.Write "<li>" & field.Name & "</li>"
        next
        Response.Write "</ul>"
    else
        Response.Write "No records found<br>"
    end if
end if

conn.Close
%>
