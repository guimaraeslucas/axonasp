# Stats Method

## Overview

The **Stats** method retrieves runtime connection pool statistics for the current database connection in G3Pix AxonASP.

## Syntax

```asp
Set result = obj.Stats()
```

## Parameters and Arguments

None.

## Return Values

Returns a **Scripting.Dictionary** object containing detailed information about the current state of the database connection pool.

## Remarks

- The dictionary keys returned by this method include:
    - **MaxOpenConnections**: The maximum number of open connections allowed.
    - **OpenConnections**: The total number of established connections (in-use + idle).
    - **InUse**: The number of connections currently active in a transaction or query.
    - **Idle**: The number of idle connections sitting in the pool.
    - **WaitCount**: The total number of connections that had to wait before being granted.
    - **WaitDurationSeconds**: The total duration in seconds that callers blocked waiting for a connection.
    - **MaxIdleClosed**: The number of connections closed because of the **SetMaxIdleConns** limit.
    - **MaxIdleTimeClosed**: The number of connections closed because of the **SetConnMaxIdleTime** limit.
    - **MaxLifetimeClosed**: The number of connections closed because of the **SetConnMaxLifetime** limit.
- This method is useful for monitoring the performance and health of the database connection pool during runtime.

## Code Example

```asp
<%
Dim db, stats
Set db = Server.CreateObject("G3DB")

If db.Open("mysql", "user:pass@tcp(localhost)/dbname") Then
    Set stats = db.Stats()
    
    Response.Write "Active Connections: " & stats("InUse") & "<br>"
    Response.Write "Idle Connections: " & stats("Idle") & "<br>"
    Response.Write "Total Open: " & stats("OpenConnections") & "<br>"
    
    db.Close
End If

Set db = Nothing
%>
```
