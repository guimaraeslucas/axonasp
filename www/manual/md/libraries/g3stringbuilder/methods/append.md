# Append Text to the Builder

## Overview
Use Append to add text to the end of the current G3STRINGBUILDER buffer.

## Syntax
```asp
sb.Append text
```

## Parameters and Arguments
- text (String, Required): The value appended to the current buffer. Non-string values are coerced to String using VBScript conversion semantics.

## Return Values
Returns Empty.

## Remarks
- Each call appends content in sequence.
- Append does not return the builder object.
- Use ToString to read the full accumulated output.

## Code Example
```asp
<%
Option Explicit

Dim sb
Set sb = Server.CreateObject("G3STRINGBUILDER")

sb.Append "Hello"
sb.Append " "
sb.Append "AxonASP"

Response.Write sb.ToString()
Set sb = Nothing
%>
```
