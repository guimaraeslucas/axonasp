# axdate

## Overview

The `axdate` method formats a Unix timestamp (or the current time if omitted) into a readable string using PHP-like formatting tokens.

## Syntax

```asp
result = obj.axdate(format [, timestamp])
```

## Parameters and Arguments

- **format** (String): A string representing the desired date/time format. It supports PHP-like tokens (e.g., "Y-m-d H:i:s").
- **timestamp** (Integer, Optional): The Unix timestamp to format. If omitted, the current system time is used.

## Return Values

Returns a String containing the formatted date and time.

## Remarks

- This method is part of the G3Pix AxonASP library.
- It supports localized month and weekday names based on the VM configuration.
- Common tokens:
    - **Y**: 4-digit year.
    - **m**: Month with leading zeros (01-12).
    - **d**: Day of the month with leading zeros (01-31).
    - **H**: 24-hour format of an hour (00-23).
    - **i**: Minutes with leading zeros (00-59).
    - **s**: Seconds with leading zeros (00-59).
- Method names in G3Pix AxonASP are case-insensitive.

## Code Example

```asp
<%
Option Explicit
Dim ax, formattedDate
Set ax = Server.CreateObject("G3AXON.FUNCTIONS")

' Format current date
formattedDate = ax.axdate("Y-m-d H:i:s")
Response.Write "Current Date: " & formattedDate & "<br>"

' Format specific timestamp (e.g., 1735689600 for 2025-01-01)
formattedDate = ax.axdate("l, F j, Y", 1735689600)
Response.Write "Specific Date: " & formattedDate

Set ax = Nothing
%>
```
