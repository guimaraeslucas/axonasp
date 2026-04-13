# axgenerateguid

## Overview

The `axgenerateguid` method generates a cryptographically secure version 4 Globally Unique Identifier (GUID) string.

## Syntax

```asp
result = obj.axgenerateguid()
```

## Parameters and Arguments

This method does not accept any parameters.

## Return Values

Returns a String containing a unique 36-character GUID (e.g., `f47ac10b-58cc-4372-a567-0e02b2c3d479`).

## Remarks

- This method is part of the G3Pix AxonASP library.
- It uses a cryptographically secure random number generator to ensure uniqueness.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, guid
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

guid = ax.axgenerateguid()

Response.Write "Generated GUID: " & guid

Set ax = Nothing
%>
```
