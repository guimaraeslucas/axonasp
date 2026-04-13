# axtime

## Overview

The `axtime` method returns the current Unix timestamp, which is the number of seconds that have elapsed since January 1, 1970 (UTC).

## Syntax

```asp
result = obj.axtime()
```

## Parameters and Arguments

This method does not accept any parameters.

## Return Values

Returns an Integer representing the current Unix timestamp in seconds.

## Remarks

- This method is part of the G3Pix AxonASP library.
- The returned value is based on the system's current time and configured time zone.
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, ts
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

ts = ax.axtime()

Response.Write "Current Unix Timestamp: " & ts

Set ax = Nothing
%>
```
