<%
Option Explicit
debug_asp_code = "TRUE"
Response.ContentType = "text/plain"

Response.Write "Testing MD5 Components" & vbCrLf & vbCrLf

' Test CLng with large numbers
Response.Write "CLng tests:" & vbCrLf
Response.Write "CLng(2147483647) = " & CLng(2147483647) & vbCrLf
Response.Write "CLng(1073741824) = " & CLng(1073741824) & vbCrLf & vbCrLf

' Test array
Response.Write "Array test:" & vbCrLf
Dim arr(5)
arr(0) = CLng(1)
arr(1) = CLng(255)
Response.Write "arr(0) = " & arr(0) & vbCrLf
Response.Write "arr(1) = " & arr(1) & vbCrLf & vbCrLf

' Load the MD5 class
Set md5Obj = aspL.plugin("md5")

' Call MD5 with a simple string
Response.Write "Calling MD5..." & vbCrLf
On Error Resume Next
Dim result
result = md5Obj.md5("a", 32)
If Err.Number <> 0 Then
    Response.Write "ERROR: " & Err.Description & " (#" & Err.Number & ")" & vbCrLf
   Response.Write "Source: " & Err.Source & vbCrLf
    Err.Clear
Else
    Response.Write "Result: " & result & vbCrLf
    Response.Write "Length: " & Len(result) & vbCrLf
End If
On Error Goto 0
%>
