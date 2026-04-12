<%
@Language = "VBSCRIPT" CodePage = "65001"
%>
<%
Option Explicit

Dim json
Dim payload
Dim Output

Response.ContentType = "application/json"
Response.Charset = "utf-8"

Set json = Server.CreateObject("G3JSON")
Set payload = json.NewObject()

payload("success") = True
payload("service") = "rest-basic"
payload("endpoint") = "status"
payload("message") = "REST status endpoint is running"
payload("timestamp") = CStr(Now())

Output = json.Stringify(payload)
Response.Write Output

Set payload = Nothing
Set json = Nothing
%>
