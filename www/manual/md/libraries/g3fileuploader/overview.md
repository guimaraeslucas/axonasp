# G3FILEUPLOADER Library

## Overview
The `G3FILEUPLOADER` object provides a high-performance system for securely processing and managing HTTP multipart file uploads in the AxonASP environment. You can control validation logic, restrict extensions, limit file sizes, and retrieve deep metadata describing individual files or the entire upload batch.

## Syntax
```asp
Set uploader = Server.CreateObject("G3FILEUPLOADER")
```

## Properties and Methods
The object exposes specific methods to parse forms and properties to refine the validation logic. View the properties and methods listing pages for details.

## Return Values
Returns a native AxonASP object handle referencing the file uploader instance.

## Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.MaxFileSize = 5242880 ' 5 MB
uploader.BlockExtension "exe"

Dim result
Set result = uploader.Process("myFile", "/uploads/docs", "newdocument.pdf")
If result("IsSuccess") Then
    Response.Write "File saved at: " & result("RelativePath")
Else
    Response.Write "Error: " & result("ErrorMessage")
End If
%>
```
