# Convert RGB to HTML Hexadecimal Color

## Overview

Converts individual Red, Green, and Blue color values into an HTML hexadecimal color string.

## Syntax

```vbscript
strHex = obj.axrgbtohex(r, g, b)
```

## Parameters

- **r** (Integer): The red color component (0-255).
- **g** (Integer): The green color component (0-255).
- **b** (Integer): The blue color component (0-255).

## Return Value

String. The HTML hexadecimal color code.

## Remarks

If fewer than three arguments are provided, it defaults to returning `#000000`.

## Code Example

```vbscript
Dim obj, strHex
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strHex = obj.axrgbtohex(255, 128, 0)
Response.Write strHex ' Outputs: #ff8000
```