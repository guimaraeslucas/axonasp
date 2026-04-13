# axgetdefaultcss

## Overview

The `axgetdefaultcss` method retrieves the text content of the CSS file specified in the `axfunctions.ax_default_css_path` configuration setting of the `axonasp.toml` file.

## Syntax

```asp
result = obj.axgetdefaultcss()
```

## Parameters and Arguments

This method does not accept any parameters.

## Return Values

Returns a String containing the raw CSS content from the configured file. Returns an empty string if no path is configured or if the file cannot be read.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It allows ASP pages to dynamically include or inline the default system stylesheet without hardcoding file paths.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, css
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

css = ax.axgetdefaultcss()

If css <> "" Then
    Response.Write "<style>" & vbCrLf & css & vbCrLf & "</style>"
End If

Set ax = Nothing
%>
```
