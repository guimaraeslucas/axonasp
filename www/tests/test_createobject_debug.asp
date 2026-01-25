<%@ Page Language="VBScript" %>
<html>
<head><title>CreateObject Debug</title></head>
<body>
<h1>CreateObject Debug Test</h1>

<%
Response.Write "<h2>Test 1: G3JSON</h2>"
On Error Resume Next
Dim obj1
Set obj1 = Server.CreateObject("G3JSON")
Response.Write "Error Number: " & Err.Number & "<br>"
Response.Write "Error Description: " & Err.Description & "<br>"
Response.Write "Object Type: " & TypeName(obj1) & "<br>"
Response.Write "IsObject: " & IsObject(obj1) & "<br>"
Response.Write "IsEmpty: " & IsEmpty(obj1) & "<br>"
If IsObject(obj1) Then
    Response.Write "<p style='color:green'>SUCCESS: Object created</p>"
Else
    Response.Write "<p style='color:red'>FAILED: Not an object</p>"
End If
Err.Clear

Response.Write "<hr><h2>Test 2: ADODB.Connection</h2>"
Dim obj2
Set obj2 = Server.CreateObject("ADODB.Connection")
Response.Write "Error Number: " & Err.Number & "<br>"
Response.Write "Error Description: " & Err.Description & "<br>"
Response.Write "Object Type: " & TypeName(obj2) & "<br>"
Response.Write "IsObject: " & IsObject(obj2) & "<br>"
If IsObject(obj2) Then
    Response.Write "<p style='color:green'>SUCCESS: Object created</p>"
Else
    Response.Write "<p style='color:red'>FAILED: Not an object</p>"
End If
Err.Clear

Response.Write "<hr><h2>Test 3: Scripting.Dictionary</h2>"
Dim obj3
Set obj3 = Server.CreateObject("Scripting.Dictionary")
Response.Write "Error Number: " & Err.Number & "<br>"
Response.Write "Error Description: " & Err.Description & "<br>"
Response.Write "Object Type: " & TypeName(obj3) & "<br>"
Response.Write "IsObject: " & IsObject(obj3) & "<br>"
If IsObject(obj3) Then
    Response.Write "<p style='color:green'>SUCCESS: Object created</p>"
    ' Test using it
    obj3.Add "test", "value"
    Response.Write "Dictionary Count: " & obj3.Count & "<br>"
Else
    Response.Write "<p style='color:red'>FAILED: Not an object</p>"
End If
Err.Clear

Response.Write "<hr><h2>Test 4: Direct Call (no Set)</h2>"
Dim result
result = Server.HTMLEncode("<test>")
Response.Write "HTMLEncode result: " & result & "<br>"
Response.Write "Type: " & TypeName(result) & "<br>"

Response.Write "<hr><h2>Test 5: MapPath</h2>"
Dim path
path = Server.MapPath("/")
Response.Write "MapPath result: " & path & "<br>"
%>

</body>
</html>
