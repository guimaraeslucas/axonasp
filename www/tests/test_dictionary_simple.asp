<%
@ Language = "VBScript" CodePage = "65001"
%>
<%
Dim dict2
Set dict2 = Server.CreateObject("Scripting.Dictionary")
dict2.CompareMode = 1   ' vbTextCompare – Case - insensitive
dict2("Alpha") = 100
dict2("Beta") = 200
dict2("Gamma") = 300
dict2("Delta") = 400

Response.ContentType = "text/html"
Response.Write "<table border='1'>" & vbCrLf
Response.Write "<tr><th>Key</th><th>Value</th><th>Exists?</th></tr>" & vbCrLf

Dim dk
For Each dk In dict2
    Response.Write "<tr><td>" & dk & "</td><td>" & dict2(dk) & "</td><td>" & dict2.Exists(dk) & "</td></tr>" & vbCrLf
Next

Response.Write "<tr><td>Epsilon</td><td>—</td><td>" & dict2.Exists("Epsilon") & "</td></tr>" & vbCrLf
Response.Write "</table>" & vbCrLf

Response.Write "<p><strong>Count:</strong> " & dict2.Count & "</p>" & vbCrLf
Response.Write "<p><strong>Keys:</strong> " & Join(dict2.Keys(), ", ") & "</p>" & vbCrLf
%>
