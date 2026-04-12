<%
' Test Dictionary indexing fix
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "name", "John"
dict.Add "age", 30

' Test direct indexing
Response.Write "Name: " & dict("name") & "<br>"
Response.Write "Age: " & dict("age") & "<br>"

' Test with array
Dim arr : arr = Array()
ReDim arr(1)
Set arr(0) = dict

Dim dict2
Set dict2 = Server.CreateObject("Scripting.Dictionary")
dict2.Add "name", "Jane"
dict2.Add "age", 25
Set arr(1) = dict2

' Test For Each over array with dictionary indexing
Response.Write "<br>People:<br>"
Dim item
For Each item In arr
    Response.Write "- " & item("name") & " is " & item("age") & " years old<br>"
Next

Response.Write "<br>SUCCESS: Dictionary indexing works!"
%>
