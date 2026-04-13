# SetMaxOpenConns Method

## Overview

The **SetMaxOpenConns** method sets the maximum number of open database connections in the connection pool in G3Pix AxonASP.

## Syntax

```asp
obj.SetMaxOpenConns(n)
```

## Parameters and Arguments

- **n** (Integer, Required): The maximum number of open connections to the database. Use 0 for no limit.

## Return Values

Returns **Empty**.

## Remarks

- This setting limits the total number of connections to the database from the G3Pix AxonASP server, including both in-use and idle connections.
- It is crucial for preventing the application from overwhelming the database server with too many concurrent connections.
- If this limit is reached, any new requests for a connection will wait until another connection is returned to the pool.

## Code Example

```asp
<%
Dim db
Set db = Server.CreateObject("G3DB")

' Limit to 50 concurrent open connections
db.SetMaxOpenConns 50

If db.Open("mysql", "user:pass@tcp(localhost)/dbname") Then
    ' Database operations...
    db.Close
End If

Set db = Nothing
%>
```
