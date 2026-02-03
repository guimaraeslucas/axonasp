<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
Response.Write "=== Right() and Path Test ===" & vbCrLf & vbCrLf

path1 = Server.MapPath("/uploads/")
Response.Write "Server.MapPath(""/uploads/"") = [" & path1 & "]" & vbCrLf

lastChar = Right(path1, 1)
Response.Write "Right(path1, 1) = [" & lastChar & "]" & vbCrLf

Response.Write "lastChar = ""\"" ? " & (lastChar = "\") & vbCrLf
Response.Write "lastChar <> ""\"" ? " & (lastChar <> "\") & vbCrLf

' Test the full condition
If Right(path1, 1) <> "\" Then 
    path1 = path1 & "\"
    Response.Write "Path was modified to: [" & path1 & "]" & vbCrLf
Else
    Response.Write "Path already ends with backslash" & vbCrLf
End If

Response.Write vbCrLf & "Final path: [" & path1 & "]" & vbCrLf

Response.Write vbCrLf & "=== Test Complete ===" & vbCrLf
%>
