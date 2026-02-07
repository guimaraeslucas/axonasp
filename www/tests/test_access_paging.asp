<%
' Test Access database with paging similar to DataTables
Option Explicit

Response.Write "<h1>Access Database Paging Test</h1>"

Dim conn, rs, sql
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("../db/sample.mdb")

Response.Write "<p>Connection opened successfully</p>"

' Create recordset with cursor and lock types like asplite
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorType = 1  ' adOpenKeyset
rs.LockType = 3    ' adLockOptimistic
Set rs.ActiveConnection = conn

sql = "SELECT contact.iId, contact.sText, contact.iNumber, contact.dDate, country.sText as countryName " & _
      "FROM contact LEFT JOIN country ON contact.iCountryID = country.iId " & _
      "ORDER BY contact.iId"

Response.Write "<p>Opening recordset with SQL:<br>" & Server.HTMLEncode(sql) & "</p>"

' Open the recordset
rs.Open sql

Response.Write "<p>Recordset opened successfully</p>"

' Test RecordCount (this might trigger materialization)
Response.Write "<p>Testing RecordCount...</p>"
Dim rCount
rCount = rs.RecordCount
Response.Write "<p>RecordCount: " & rCount & "</p>"

' Test AbsolutePosition
Response.Write "<p>Testing AbsolutePosition...</p>"
If rCount > 0 Then
    rs.AbsolutePosition = 1
    Response.Write "<p>AbsolutePosition set to 1</p>"
End If

' Test PageSize
Response.Write "<p>Testing PageSize...</p>"
rs.PageSize = 10
Response.Write "<p>PageSize set to 10</p>"

' Display first few records
Response.Write "<h2>First 3 Records:</h2><table border='1'>"
Response.Write "<tr><th>iId</th><th>sText</th><th>iNumber</th><th>dDate</th><th>countryName</th></tr>"

Dim i
For i = 1 To 3
    If Not rs.EOF Then
        Response.Write "<tr>"
        Response.Write "<td>" & rs.Fields("iId").Value & "</td>"
        Response.Write "<td>" & rs.Fields("sText").Value & "</td>"
        Response.Write "<td>" & rs.Fields("iNumber").Value & "</td>"
        Response.Write "<td>" & rs.Fields("dDate").Value & "</td>"
        Response.Write "<td>" & rs.Fields("countryName").Value & "</td>"
        Response.Write "</tr>"
        rs.MoveNext
    End If
Next

Response.Write "</table>"

rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing

Response.Write "<p><strong>Test completed successfully!</strong></p>"
%>
