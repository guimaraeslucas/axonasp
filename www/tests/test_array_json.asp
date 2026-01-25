<%@ language="VBScript" %>
<% Option Explicit %>
<!DOCTYPE html>
<html>
<head>
    <title>Array to JSON Test</title>
</head>
<body>
    <h1>Array to JSON Test</h1>
    <pre>
<%
Response.Write "=== Test 1: Simple Array ===<br>"
Dim arr1 : arr1 = Array("item1", "item2", "item3")
Response.Write "Array created with " & (UBound(arr1) + 1) & " items<br>"

Dim i
For i = 0 To UBound(arr1)
    Response.Write "arr1(" & i & ") = " & arr1(i) & "<br>"
Next
Response.Write "<br>"

Response.Write "=== Test 2: ReDim Array ===<br>"
Dim arr2 : arr2 = Array()
ReDim arr2(2)
Response.Write "Array after ReDim(2): " & (UBound(arr2) + 1) & " items<br>"

Set arr2(0) = Server.CreateObject("Scripting.Dictionary")
arr2(0).Add "type", "text"
arr2(0).Add "name", "field1"

Set arr2(1) = Server.CreateObject("Scripting.Dictionary")
arr2(1).Add "type", "submit"
arr2(1).Add "name", "button1"

Set arr2(2) = Server.CreateObject("Scripting.Dictionary")
arr2(2).Add "type", "hidden"
arr2(2).Add "name", "token"

Response.Write "Filled array with dictionaries<br>"
For i = 0 To UBound(arr2)
    Response.Write "arr2(" & i & ") = " & TypeName(arr2(i)) & "<br>"
    If TypeName(arr2(i)) = "Dictionary" Then
        Response.Write "  type: " & arr2(i)("type") & "<br>"
        Response.Write "  name: " & arr2(i)("name") & "<br>"
    End If
Next
Response.Write "<br>"

Response.Write "=== Test 3: For Each over array ===<br>"
Dim item
For Each item In arr2
    Response.Write "Item TypeName: " & TypeName(item) & "<br>"
    If IsObject(item) Then
        If TypeName(item) = "Dictionary" Then
            Response.Write "  type: " & item("type") & ", name: " & item("name") & "<br>"
        End If
    End If
Next
Response.Write "<br>"

Response.Write "=== Test 4: Check if aspLite is available ===<br>"
On Error Resume Next
Dim testAspl
Set testAspl = aspL
If Err.Number = 0 Then
    Response.Write "aspL object is available<br>"
    If Not testAspl Is Nothing Then
        Response.Write "aspL.json available: " & (Not testAspl.json Is Nothing) & "<br>"
        
        If Not testAspl.json Is Nothing Then
            Response.Write "<br>=== Test 5: Call aspl.json.toJson ===<br>"
            Dim jsonResult
            jsonResult = testAspl.json.toJson("testArray", arr2, False)
            Response.Write "JSON Result:<br>"
            Response.Write jsonResult & "<br>"
        End If
    End If
Else
    Response.Write "aspL object NOT available (Error: " & Err.Description & ")<br>"
End If
On Error GoTo 0
%>
    </pre>
</body>
</html>
