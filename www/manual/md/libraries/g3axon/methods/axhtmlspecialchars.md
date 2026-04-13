# Escape HTML Special Characters

## Overview

Escapes special characters in a string to their corresponding HTML entities.

## Syntax

```vbscript
strEscaped = obj.axhtmlspecialchars(str)
```

## Parameters

- **str** (String): The string to escape.

## Return Value

String. The HTML escaped string.

## Remarks

Critical for preventing Cross-Site Scripting (XSS) vulnerabilities when rendering user input directly in an HTML page.

## Code Example

```vbscript
Dim obj, strHTML
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strHTML = obj.axhtmlspecialchars("<script>alert(1);</script>")
Response.Write strHTML
```