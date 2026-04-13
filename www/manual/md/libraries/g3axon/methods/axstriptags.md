# axstriptags

## Overview

The `axstriptags` method removes all HTML and XML tags from a provided string using regular expressions. This is useful for sanitizing input or preparing text for plain-text display.

## Syntax

```asp
result = obj.axstriptags(inputString)
```

## Parameters and Arguments

- **inputString** (String): The source string containing HTML or XML tags to be removed.

## Return Values

Returns a String containing the text with all tags removed. If the input is empty or no tags are found, it returns the original string.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It uses a regular expression to identify and remove anything within `<` and `>`.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, html, plainText
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

html = "<p>Hello <b>AxonASP</b>!</p><!-- comment -->"
plainText = ax.axstriptags(html)

' Output: Hello AxonASP!
Response.Write plainText

Set ax = Nothing
%>
```
