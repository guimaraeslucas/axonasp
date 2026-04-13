# GetError Method

## Overview

The **GetError** method retrieves the most recent error message resulting from database operations in G3Pix AxonASP.

## Syntax

```asp
result = obj.GetError()
```

## Parameters and Arguments

None. This method is also accessible through the alias **GetLastError**.

## Return Values

Returns a **String** containing the error description of the last failed operation. If no error has occurred, it returns an empty string.

## Remarks

- This method should be called immediately after a database operation returns **False** or an **Empty** value to determine the root cause of the failure.
- Typical errors returned include connection timeouts, SQL syntax errors, authentication failures, and driver-specific issues.
- The same error information can also be accessed directly via the **LastError** property.

## Code Example

```asp
<%
Dim db, res
Set db = Server.CreateObject("G3DB")

' Attempt to open with invalid parameters
If Not db.Open("mysql", "invalid_host:3306") Then
    Response.Write "Database Error: " & db.GetError()
End If

Set db = Nothing
%>
```
