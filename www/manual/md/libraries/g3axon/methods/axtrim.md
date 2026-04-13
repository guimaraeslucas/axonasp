# Trim String Whitespace or Characters

## Overview

Removes whitespace or specified characters from the beginning and end of a string.

## Syntax

```vbscript
strResult = obj.axtrim(str[, chars])
```

## Parameters

- **str** (String): The string to trim.
- **chars** (String, Optional): Specific characters to remove. If omitted, defaults to standard whitespace characters.

## Return Value

String. The trimmed string.

## Remarks

Useful for cleaning up user input or formatted text data.

## Code Example

```vbscript
Dim obj, strTrimmed
Set obj = Server.CreateObject("G3AXON.FUNCTIONS")
strTrimmed = obj.axtrim("  Hello World  ")
Response.Write strTrimmed ' Outputs: Hello World
```