# SetConnMaxIdleTime Method

## Overview

The **SetConnMaxIdleTime** method sets the maximum duration in seconds that a connection can remain idle in the pool before being closed in G3Pix AxonASP.

## Syntax

```asp
obj.SetConnMaxIdleTime(seconds)
```

## Parameters and Arguments

- **seconds** (Integer, Required): The maximum idle time for a connection in seconds. Use 0 for no limit.

## Return Values

Returns **Empty**.

## Remarks

- This setting helps to proactively reclaim system resources from database connections that are not actively in use.
- It is particularly useful for environments with intermittent database traffic.
- This method should be called before or after establishing a connection to configure the pool behavior.

## Code Example

```asp
<%
Dim db
Set db = Server.CreateObject("G3DB")

' Set maximum idle time to 120 seconds (2 minutes)
db.SetConnMaxIdleTime 120

If db.Open("mysql", "user:pass@tcp(localhost)/dbname") Then
    ' Database operations...
    db.Close
End If

Set db = Nothing
%>
```
