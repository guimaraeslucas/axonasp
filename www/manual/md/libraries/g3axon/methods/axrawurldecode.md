# Decode Raw URL String

## Overview

Decodes a URL-encoded string by treating `+` as spaces before applying RFC 3986 unescaping.

## Syntax

```vbscript
strDecoded = obj.axrawurldecode(str)
```

## Parameters

- **str** (String): The raw URL-encoded string to decode.

## Return Value

String. The decoded string.

## Remarks

This format is slightly different from standard query strings, enforcing a specific handling of `+` symbols as spaces.

## Code Example

```vbscript
Dim obj, strDecode
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strDecode = obj.axrawurldecode("Hello+World")
Response.Write strDecode ' Outputs: Hello World
```