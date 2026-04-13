# axw

## Overview
Writes HTML-escaped text to the current HTTP response in G3Pix AxonASP.

## Syntax
```asp
obj.axw(text)
```

## Parameters and Arguments
- **text** (String): The text to be written to the response.

## Return Values
Returns Empty.

## Remarks
This method is a secure alternative to `Response.Write` as it automatically escapes HTML characters, protecting against cross-site scripting (XSS) vulnerabilities. Use it when displaying data that may contain HTML tags that should be rendered as literal text.

## Code Example
```asp
<%
Option Explicit
Dim obj
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")

' This will render exactly as: <b>Secure Output</b>
obj.axw "<b>Secure Output</b>"

Set obj = Nothing
%>
```
