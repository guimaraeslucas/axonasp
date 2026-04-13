# Driver Property

## Overview

The **Driver** property gets or sets the database driver name for the G3Pix AxonASP database connection.

## Syntax

To set the property:
```asp
obj.Driver = "mysql"
```
To get the property:
```asp
result = obj.Driver
```

## Return Values

Returns a **String** representing the canonical driver name (e.g., "mysql", "postgres", "mssql", "sqlite", "oracle").

## Remarks

- This property can be set manually before calling the **Open** or **OpenFromEnv** methods.
- After a successful connection, this property returns the normalized driver name used for the current connection.
- Supported aliases are automatically converted to their canonical forms (e.g., "mariadb" becomes "mysql").

## Code Example

```asp
<%
Dim db, driverName
Set db = Server.CreateObject("G3DB")

' Pre-configure the driver
db.Driver = "mysql"

Response.Write "Configured Driver: " & db.Driver & "<br>"

' Attempt to open a connection
If db.OpenFromEnv() Then
    Response.Write "Connected using driver: " & db.Driver
    db.Close
End If

Set db = Nothing
%>
```
