<%
@Language = "VBSCRIPT" CodePage = "65001"
%>
<%
Option Explicit

Dim json
Dim payload
Dim nameParam
Dim Output

Response.ContentType = "application/json"
Response.Charset = "utf-8"

nameParam = Trim(Request.QueryString("name"))

If nameParam = "" Then
    nameParam = "guest"
End If

Set json = Server.CreateObject("G3JSON")
Set payload = json.NewObject()

payload("success") = True
payload("service") = "rest-basic"
payload("endpoint") = "echo"
payload("name") = nameParam
payload("message") = "Echo endpoint executed successfully"

Output = json.Stringify(payload)
Response.Write Output

Set payload = Nothing
Set json = Nothing
%>
