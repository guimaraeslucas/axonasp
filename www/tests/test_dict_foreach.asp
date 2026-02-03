<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Write "=== Dictionary For Each Test ===" & vbCrLf & vbCrLf

' Create dictionary
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "key1", "value1"
dict.Add "key2", "value2"
dict.Add "key3", "value3"

Response.Write "Dictionary Count: " & dict.Count & vbCrLf & vbCrLf

' Test 1: For Each over Keys() return value
Response.Write "Test 1: For Each over dict.Keys" & vbCrLf
keysArr = dict.Keys
Response.Write "Type of keysArr: " & TypeName(keysArr) & vbCrLf
Response.Write "Iterating..." & vbCrLf
For Each k In keysArr
    Response.Write "  Key: " & k & vbCrLf
Next
Response.Write vbCrLf

' Test 2: For Each over Items() return value
Response.Write "Test 2: For Each over dict.Items" & vbCrLf
itemsArr = dict.Items
Response.Write "Type of itemsArr: " & TypeName(itemsArr) & vbCrLf
Response.Write "Iterating..." & vbCrLf
For Each it In itemsArr
    Response.Write "  Item: " & it & vbCrLf
Next
Response.Write vbCrLf

' Test 3: Direct For Each over dictionary
Response.Write "Test 3: For Each over dict (direct)" & vbCrLf
For Each k In dict
    Response.Write "  Key: " & k & vbCrLf
Next
Response.Write vbCrLf

' Test 4: For Each directly over dict.Keys
Response.Write "Test 4: For Each In dict.Keys (direct call)" & vbCrLf
For Each k In dict.Keys
    Response.Write "  Key: " & k & vbCrLf
Next
Response.Write vbCrLf

Response.Write "=== Test Complete ===" & vbCrLf
%>
