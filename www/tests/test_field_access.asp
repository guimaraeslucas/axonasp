<%
' Debug Access Field Access
debug_asp_code = "TRUE"

response.write "<h1>Debug Field Access</h1>"

On Error Resume Next

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")

Dim dbPath
dbPath = "E:\lucas\Desktop\Sites\LGGM-TCP\modules\image\ASP\axonasp\www\db\sample.mdb"

conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
conn.Open

response.write "<p>Connected: " & (conn.State = 1) & "</p>"

Dim rs
Set rs = conn.Execute("SELECT * FROM country")

response.write "<p>Recordset received: " & Not (rs Is Nothing) & "</p>"
response.write "<p>EOF: " & rs.EOF & "</p>"

' Test Fields property
response.write "<h2>Testing Fields Property</h2>"

Dim flds
Set flds = rs.Fields

response.write "<p>Fields object: " & Not (flds Is Nothing) & "</p>"
response.write "<p>Fields TypeName: " & TypeName(flds) & "</p>"

' Test getting field by CallMethod
response.write "<h2>Testing Field Access Methods</h2>"

Dim fieldValue1
fieldValue1 = flds.Item("iId")
response.write "<p>flds.Item('iId'): " & fieldValue1 & " (TypeName: " & TypeName(fieldValue1) & ")</p>"

' Test direct subscript
Dim fieldValue2
fieldValue2 = flds("iId")
response.write "<p>flds('iId'): " & fieldValue2 & " (TypeName: " & TypeName(fieldValue2) & ")</p>"

' Test via rs.Fields
Dim fieldValue3
fieldValue3 = rs.Fields("iId")
response.write "<p>rs.Fields('iId'): " & fieldValue3 & " (TypeName: " & TypeName(fieldValue3) & ")</p>"

' Test iId field value
response.write "<h2>Trying different access patterns</h2>"

' Pattern 1: Direct Fields.Item call
Dim val1
val1 = rs.Fields.Item("iId")
response.write "<p>rs.Fields.Item('iId'): " & val1 & "</p>"

' Pattern 2: Two-step access
Dim fld
Set fld = rs.Fields
val1 = fld("iId")
response.write "<p>Two-step fld('iId'): " & val1 & "</p>"

rs.Close
conn.Close

response.write "<h3>Test Complete</h3>"
%>
