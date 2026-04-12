<%
on error resume next

Dim x
Response.Write "x TypeName: " & TypeName(x) & vbCrLf
Response.Write "x Is Nothing: " & (x Is Nothing) & " | Err: " & err.number & " | Desc: " & err.description & vbCrLf
err.clear

Set x = Nothing
Response.Write "x TypeName: " & TypeName(x) & vbCrLf
Response.Write "x Is Nothing: " & (x Is Nothing) & " | Err: " & err.number & " | Desc: " & err.description & vbCrLf
err.clear

Dim d
Set d = CreateObject("Scripting.Dictionary")
Response.Write "d TypeName: " & TypeName(d) & vbCrLf
Response.Write "d Is Nothing: " & (d Is Nothing) & " | Err: " & err.number & " | Desc: " & err.description & vbCrLf
err.clear
%>
