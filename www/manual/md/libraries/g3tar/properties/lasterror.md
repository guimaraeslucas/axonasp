# LastError Property

## Overview
Gets the description of the last error encountered during a G3TAR operation.

## Syntax
```asp
Dim errorMsg
errorMsg = obj.LastError
```

## Parameters and Arguments
- Getter: None.

## Return Values
Returns a `String` detailing the most recent internal operation error.

## Remarks
- This property is read-only.
- Use this to inspect failures instead of relying solely on Boolean returns.

## Code Example
```asp
<%
Option Explicit
Dim obj, errorMsg
Set obj = Server.CreateObject("G3TAR")
If Not obj.Open("C:\temp\missing.tar") Then
    errorMsg = obj.LastError
    Response.Write errorMsg
End If
Set obj = Nothing
%>
```

