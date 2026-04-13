# axgetremotefile

## Overview

The `axgetremotefile` method performs a synchronous HTTP GET request to a specified URL and returns the response body.

## Syntax

```asp
result = obj.axgetremotefile(url)
```

## Parameters and Arguments

- **url** (String): The full URL of the remote file to fetch (must start with `http://` or `https://`).

## Return Values

Returns a String containing the response body if the request is successful (HTTP 200 OK). Returns the Boolean `False` if the request fails, times out, or returns a non-200 status code.

## Remarks

- This method is part of the G3Pix AxonASP library.
- The request has a default timeout of 10 seconds.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, content
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

content = ax.axgetremotefile("https://api.example.com/data.json")

If VarType(content) = vbString Then
    Response.Write "File content received: " & Server.HTMLEncode(content)
Else
    Response.Write "Failed to fetch the remote file."
End If

Set ax = Nothing
%>
```
