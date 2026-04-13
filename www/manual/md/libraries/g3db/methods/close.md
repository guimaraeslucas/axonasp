# Close Method

## Overview

The **Close** method shuts down the database connection and releases the underlying connection pool in G3Pix AxonASP.

## Syntax

```asp
result = obj.Close()
```

## Parameters and Arguments

None.

## Return Values

Returns a **Boolean** value. It returns **True** if the connection was successfully closed or if no connection was currently open, and **False** if an error occurred during the closing process.

## Remarks

- This method should be called when database operations are complete to ensure system resources are properly released.
- After calling **Close**, the **IsOpen** property will return **False**.
- Re-opening a closed connection requires a new call to the **Open** or **OpenFromEnv** method.

## Code Example

```asp
<%
Dim db, isConnected
Set db = Server.CreateObject("G3DB")

If db.Open("sqlite", "local_data.db") Then
    ' Database operations...
    
    If db.Close() Then
        Response.Write "Connection closed successfully."
    Else
        Response.Write "Error closing connection: " & db.LastError
    End If
End If

Set db = Nothing
%>
```
