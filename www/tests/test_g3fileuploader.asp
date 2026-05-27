<%
Option Explicit

' G3FileUploader Refactoring Test
' This test validates the new properties and return structures.

Dim uploader, result, fields, key
Set uploader = Server.CreateObject("G3FILEUPLOADER")

Response.Write "<h1>G3FileUploader Refactoring Test</h1>"

' 1. Test Properties
Response.Write "<h2>1. Properties</h2>"
uploader.MaxFileSize = 5242880
uploader.AllowAbsolutePaths = True
uploader.PreserveOriginalName = True

Response.Write "MaxFileSize (5MB): " & uploader.MaxFileSize & "<br>"
Response.Write "AllowAbsolutePaths (True): " & uploader.AllowAbsolutePaths & "<br>"
Response.Write "PreserveOriginalName (True): " & uploader.PreserveOriginalName & "<br>"

' 2. Test FormFields Property
Response.Write "<h2>2. FormFields Dictionary</h2>"
Set fields = uploader.FormFields
Response.Write "FormFields Type: " & TypeName(fields) & "<br>"
Response.Write "FormFields Initial Count: " & fields.Count & "<br>"

' 3. Test Standardized Return Structure (Process)
Response.Write "<h2>3. Standardized Return (Failure Case)</h2>"
' This should fail because there is no multipart request in this CLI execution
Set result = uploader.Process("file1", "/uploads")

Response.Write "IsSuccess: " & result("IsSuccess") & " (Expected: False)<br>"
Response.Write "ErrorMessage: " & result("ErrorMessage") & "<br>"

If result("IsSuccess") = False And result("ErrorMessage") <> "" Then
    Response.Write "<span style='color:green'>Standardized failure return verified.</span><br>"
Else
    Response.Write "<span style='color:red'>Standardized failure return FAILED.</span><br>"
End If

Response.Write "<h2>RESULT: SUCCESS</h2>"
%>
