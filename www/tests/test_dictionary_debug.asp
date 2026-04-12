<%
@ Language = "VBScript" CodePage = "65001"
%>
<%
Option Explicit

Dim dict2
Set dict2 = Server.CreateObject("Scripting.Dictionary")

' Test 1: Set with Dict.Add method (traditional ASP style)
dict2.Add "key1", "value1_add"

' Test 2: Set with subscript (modern style)
dict2("key2") = "value2_subscript"

' Test 3: Set with Add again
dict2.Add "key3", 333

' Test 4: Read back values
Response.ContentType = "text/plain"
Response.Write("Test 1 (Add): key1 = " & dict2("key1") & vbCrLf)
Response.Write("Test 2 (Subscript set): key2 = " & dict2("key2") & vbCrLf)
Response.Write("Test 3 (Add numeric): key3 = " & dict2("key3") & vbCrLf)
Response.Write("Dict Count: " & dict2.Count & vbCrLf)

' Test 5: Check Keys
Dim k
Response.Write("Keys: ")
For Each k In dict2.Keys
    Response.Write(k & " ")
Next
Response.Write(vbCrLf)

' Test 6: Check Items
Response.Write("Items: ")
Dim Item
For Each Item In dict2.Items
    Response.Write("[" & Item & "] ")
Next
Response.Write(vbCrLf)
%>
