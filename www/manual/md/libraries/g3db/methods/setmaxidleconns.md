# SetMaxIdleConns Method

## Overview

The **SetMaxIdleConns** method sets the maximum number of idle database connections to be retained in the connection pool in G3Pix AxonASP.

## Syntax

```asp
obj.SetMaxIdleConns(n)
```

## Parameters and Arguments

- **n** (Integer, Required): The maximum number of idle connections to keep in the pool. A value of 0 or less will result in no idle connections being kept.

## Return Values

Returns **Empty**.

## Remarks

- This method allows for fine-tuning the resource balance between connection reuse and system overhead.
- Maintaining some idle connections can improve performance for subsequent requests by avoiding the cost of establishing new connections.
- If the number of idle connections exceeds this limit, the pool will close the extra connections.

## Code Example

```asp
<%
Dim db
Set db = Server.CreateObject("G3DB")

' Keep up to 10 idle connections in the pool
db.SetMaxIdleConns 10

If db.Open("mysql", "user:pass@tcp(localhost)/dbname") Then
    ' Database operations...
    db.Close
End If

Set db = Nothing
%>
```
