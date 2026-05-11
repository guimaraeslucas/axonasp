# Use the G3STRINGBUILDER Library

## Overview
The G3STRINGBUILDER library provides a native low-allocation string accumulation object for G3Pix AxonASP pages. It is designed for scenarios where repeated string concatenation would otherwise create high memory pressure and CPU overhead.

## Syntax
Set sb = Server.CreateObject("G3STRINGBUILDER")

## Parameters and Arguments
- Server.CreateObject input: String, required.
- Accepted ProgID: G3STRINGBUILDER.

## Return Values
Server.CreateObject returns a native object handle that exposes the Append and ToString methods.

## Remarks
- Use this object when building large responses, logs, JSON fragments, or templates in loops.
- The object stores content in an internal Go strings.Builder instance.
- Method names are case-insensitive.
- This object has no writable properties.

## Code Example
```asp
<%
Option Explicit

Dim sb
Set sb = Server.CreateObject("G3STRINGBUILDER")

sb.Append "Order #"
sb.Append "12345"
sb.Append " processed"

Response.Write sb.ToString()
Set sb = Nothing
%>
```
