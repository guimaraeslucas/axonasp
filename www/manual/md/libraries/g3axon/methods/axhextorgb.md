# Convert HTML Hexadecimal to RGB Color

## Overview

Converts an HTML hexadecimal color string into an equivalent RGB function format string.

## Syntax

```vbscript
strRgb = obj.axhextorgb(hex)
```

## Parameters

- **hex** (String): The hexadecimal color string.

## Return Value

String. The RGB representation.

## Remarks

If the hex code format is invalid, it defaults to returning `rgb(0,0,0)`.

## Code Example

```vbscript
Dim obj, strRgb
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strRgb = obj.axhextorgb("#0080FF")
Response.Write strRgb ' Outputs: rgb(0,128,255)
```