# FormValue Method

## Overview
Alias for the `Form` method. Retrieves a non-file form field value.

## Syntax
```asp
value = uploader.FormValue(fieldName)
```

## Parameters and Arguments
- `fieldName` (String, Required): The name of the form input field.

## Return Values
Returns a **String** value or **Empty**.

## Remarks
- See the `Form` method documentation for detailed behavior.

## Code Example
```asp
<%
Dim uploader, userId
Set uploader = Server.CreateObject("G3FILEUPLOADER")
userId = uploader.FormValue("userId")
%>
```
