# IsValidExtension Method

## Overview
Checks if a given file extension is allowed based on the current configuration of the uploader.

## Syntax
```asp
isValid = uploader.IsValidExtension(extension)
```

## Parameters and Arguments
- `extension` (String, Required): The extension to validate (e.g., "png").

## Return Values
Returns a **Boolean** indicating if the extension is allowed (**True**) or blocked/not allowed (**False**).

## Remarks
- The validation logic considers both `BlockedExtensions` and `AllowedExtensions` (if `SetUseAllowedOnly` is enabled).
- Leading dots are optional and handled automatically.

## Code Example
```asp
<%
Dim uploader
Set uploader = Server.CreateObject("G3FILEUPLOADER")
uploader.BlockExtension "exe"

If uploader.IsValidExtension("jpg") Then
    Response.Write "JPG is allowed."
End If

If Not uploader.IsValidExtension("exe") Then
    Response.Write "EXE is blocked."
End If
%>
```
