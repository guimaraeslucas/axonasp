# ProcessAll Method

## Overview
Processes and saves all files included in the multipart request. Also supports the `SaveAll` alias.

## Syntax
```asp
results = uploader.ProcessAll(targetDir)
```

## Parameters and Arguments
- `targetDir` (String, Optional): The destination virtual directory. Defaults to "./".

## Return Values
Returns an **Array of Dictionary** objects. Each Dictionary represents the result for one processed file, including `IsSuccess`, `ErrorMessage`, and file metadata.

## Remarks
- If any file fails validation (e.g., restricted extension), its entry in the result array will have `IsSuccess` set to **False**.
- Files are saved in the order they are received in the HTTP request.

## Code Example
```asp
<%
Dim uploader, results, i, res
Set uploader = Server.CreateObject("G3FILEUPLOADER")
results = uploader.ProcessAll("/uploads/batch")

For i = 0 To UBound(results)
    Set res = results(i)
    If res("IsSuccess") Then
        Response.Write "Saved: " & res("OriginalFileName") & "<br>"
    Else
        Response.Write "Failed: " & res("OriginalFileName") & " (" & res("ErrorMessage") & ")<br>"
    End If
Next
%>
```
