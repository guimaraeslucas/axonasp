<%@Language="VBSCRIPT"%>
<%
Option Explicit

Response.Write "<html><body>"
Response.Write "<h1>Testing G3STRINGBUILDER Library</h1>"

Dim sb
Set sb = Server.CreateObject("G3STRINGBUILDER")

If IsObject(sb) Then
    Response.Write "<p>PASS: G3STRINGBUILDER object created successfully.</p>"
Else
    Response.Write "<p>FAIL: Could not create G3STRINGBUILDER object.</p>"
End If

Dim i
For i = 1 To 5
    sb.Append "Line"
    sb.Append CStr(i)
    sb.Append ":OK;"
Next

Dim result
result = sb.ToString()

If result = "Line1:OK;Line2:OK;Line3:OK;Line4:OK;Line5:OK;" Then
    Response.Write "<p>PASS: Append and ToString returned expected output.</p>"
Else
    Response.Write "<p>FAIL: Unexpected output: " & Server.HTMLEncode(result) & "</p>"
End If

Set sb = Nothing
Response.Write "</body></html>"
%>