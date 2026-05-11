# Read the Current Builder Content with ToString

## Overview
Use ToString to read the current complete text accumulated in G3STRINGBUILDER.

## Syntax
```asp
result = sb.ToString()
```

## Parameters and Arguments
- None.

## Return Values
Returns a String containing the current accumulated content from the internal builder.

## Remarks
- ToString does not clear the internal buffer.
- Multiple calls to ToString return the latest full content at the time of each call.

## Code Example
```asp
<%
Option Explicit

Dim sb
Dim result
Set sb = Server.CreateObject("G3STRINGBUILDER")

sb.Append "A"
sb.Append "B"
sb.Append "C"

result = sb.ToString()
Response.Write result

Set sb = Nothing
%>
```
