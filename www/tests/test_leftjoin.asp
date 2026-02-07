<%
' Test Access LEFT JOIN - the specific issue
Option Explicit

Response.Write "<h1>Access LEFT JOIN Test</h1>"

Dim conn, rs, sql
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("../db/sample.mdb")

Response.Write "<p>Connection opened</p>"

' Create and open recordset with LEFT JOIN (like sampleform21_data)
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorType = 1
rs.LockType = 3
Set rs.ActiveConnection = conn

sql = "SELECT contact.iId, contact.sText, contact.iNumber, contact.dDate, country.sText as countryName " & _
      "FROM contact LEFT JOIN country ON contact.iCountryID = country.iId " & _
      "ORDER BY contact.sText"

Response.Write "<p>About to call rs.Open with LEFT JOIN query...</p>"
Response.Flush

rs.Open sql

Response.Write "<p>rs.Open succeeded!</p>"
Response.Flush

Response.Write "<p>About to get RecordCount...</p>"
Response.Flush

Dim rCount
rCount = rs.RecordCount
Response.Write "<p>RecordCount = " & rCount & "</p>"

rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing

Response.Write "<p><strong>LEFT JOIN test completed!</strong></p>"
%>
