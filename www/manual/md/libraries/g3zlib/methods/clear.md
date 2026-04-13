# Clear G3ZLIB State

## Overview

Clears the current operation context, resetting the internal state and errors of the G3ZLIB object.

## Syntax

```asp
Call obj.Clear()
```

## Parameters and Arguments

This method takes no arguments.

## Return Values

This method does not return a value.

## Remarks

- Method names are case-insensitive.
- Use this method to reset the `LastError` property and internal buffers before initiating a new independent operation on the same object.

## Code Example

```asp
<%
Option Explicit
Dim obj
Set obj = Server.CreateObject("G3ZLIB")

Call obj.Clear()

Set obj = Nothing
%>
```



