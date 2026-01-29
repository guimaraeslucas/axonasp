<%
' Simplest possible Access test
debug_asp_code = "TRUE"

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\lucas\Desktop\Sites\LGGM-TCP\modules\image\ASP\axonasp\www\db\sample.mdb"
conn.Open

Dim rs
Set rs = conn.Execute("SELECT * FROM country")

response.write "TypeName(rs): " & TypeName(rs) & "<br>"
response.write "EOF: " & rs.EOF & "<br><br>"

' Test 1: Access first record
response.write "<h3>First Record:</h3>"
Dim val1, val2
val1 = rs.Fields("iId")
val2 = rs.Fields("sText")
response.write "iId: " & val1 & " (TypeName: " & TypeName(val1) & ")<br>"
response.write "sText: " & val2 & " (TypeName: " & TypeName(val2) & ")<br><br>"

' Test 2: Move to next and try again
rs.MoveNext
response.write "<h3>After MoveNext:</h3>"
response.write "EOF: " & rs.EOF & "<br>"
Dim val3, val4
val3 = rs.Fields("iId")
val4 = rs.Fields("sText")
response.write "iId: " & val3 & " (TypeName: " & TypeName(val3) & ")<br>"
response.write "sText: " & val4 & " (TypeName: " & TypeName(val4) & ")<br>"

rs.Close
conn.Close
%>
