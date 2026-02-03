<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Write "<h2>For Each Dictionary.Items Test</h2>"
Response.Flush

' Create a dictionary with some test objects
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")

' Add some simple values
dict.Add "key1", "value1"
dict.Add "key2", "value2"
dict.Add "key3", "value3"

Response.Write "<p>Dictionary created with " & dict.Count & " items</p>"
Response.Flush

' Test For Each with Keys
Response.Write "<p>Testing For Each key In dict.Keys:</p><ul>"
Dim key
For Each key In dict.Keys
    Response.Write "<li>Key: " & key & "</li>"
Next
Response.Write "</ul>"
Response.Flush

' Test For Each with Items  
Response.Write "<p>Testing For Each item In dict.Items:</p><ul>"
Dim item
For Each item In dict.Items
    Response.Write "<li>Item: " & item & "</li>"
Next
Response.Write "</ul>"
Response.Flush

' Now test with objects (like UploadedFile)
Response.Write "<hr><p><b>Testing with objects...</b></p>"
Response.Flush

Class TestFile
    Public FileName
    Public Size
End Class

Dim objDict
Set objDict = Server.CreateObject("Scripting.Dictionary")

Dim obj1, obj2
Set obj1 = New TestFile
obj1.FileName = "file1.txt"
obj1.Size = 100

Set obj2 = New TestFile
obj2.FileName = "file2.txt"
obj2.Size = 200

objDict.Add "f1", obj1
objDict.Add "f2", obj2

Response.Write "<p>Object dictionary created with " & objDict.Count & " items</p>"
Response.Flush

Response.Write "<p>Testing For Each item In objDict.Items:</p><ul>"
Dim fileItem
For Each fileItem In objDict.Items
    Response.Write "<li>Type: " & TypeName(fileItem) & ", FileName: " & fileItem.FileName & ", Size: " & fileItem.Size & "</li>"
Next
Response.Write "</ul>"
Response.Flush

Response.Write "<p>Test complete!</p>"
%>
