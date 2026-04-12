# Send Method

## Overview

Sends data using the current transport configuration.

## Syntax

```asp
result = obj.Send(...)
```

## Parameters and Arguments

- smtpHost (String, Optional): SMTP host if not set by properties.
- smtpPort (Integer, Optional): SMTP port.
- userName (String, Optional): SMTP auth user.
- password (String, Optional): SMTP auth password.
- useTLS (Boolean, Optional): Enables TLS/STARTTLS as supported.
- Argument validation: invalid count or type raises runtime errors.

## Return Values

Returns a Variant result. Depending on the operation, this can be String, Boolean, Number, Array, Dictionary/object handle, or Empty.

## Remarks

- Method names are case-insensitive.
- Prefer explicit variable assignment and defensive checks before using returned values.
- For object values, use Set when assigning the return value.

## Code Example

```asp
<%
Option Explicit
Dim obj, result
Set obj = Server.CreateObject("G3Mail")
result = obj.Send()
If IsObject(result) Then
    Response.Write "Object returned"
Else
    Response.Write CStr(result)
End If
Set obj = Nothing
%>
```



