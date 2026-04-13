# axlastmodified

## Overview

The `axlastmodified` method returns the Unix timestamp representing the last modification time of the current ASP page.

## Syntax

```asp
result = obj.axlastmodified()
```

## Parameters and Arguments

This method does not accept any parameters.

## Return Values

Returns an Integer representing the Unix timestamp of the current file's last modification. Returns `0` if the modification time cannot be determined.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It is useful for implementing caching headers or displaying "Last Updated" information on a page.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, ts
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

ts = ax.axlastmodified()

If ts > 0 Then
    Response.Write "Page last modified on: " & ax.axdate("Y-m-d H:i:s", ts)
Else
    Response.Write "Could not determine modification time."
End If

Set ax = Nothing
%>
```
