<%
' Simple test to verify Fields enumeration
Response.Write "<h1>Testing Fields Collection</h1>"

' Include aspLite
dim aspL : set aspL = Server.CreateObject("Scripting.Dictionary")
Response.Write "aspL created<br>"

' Create a simple recordset
dim rs : set rs = Server.CreateObject("ADODB.Recordset")
Response.Write "rs created: " & TypeName(rs) & "<br>"

if rs is nothing then
    Response.Write "rs is NOTHING!<br>"
else
    Response.Write "rs is not nothing<br>"
    Response.Write "rs.EOF: " & rs.EOF & "<br>"
    Response.Write "rs.Fields type: " & TypeName(rs.Fields) & "<br>"
    
    ' Check if Fields has any items
    Response.Write "Fields.Count: " & rs.Fields.Count & "<br>"
end if
%>
