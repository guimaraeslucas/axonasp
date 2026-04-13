# Open Method

## Overview

The **Open** method establishes a connection to a database using the specified driver and Data Source Name (DSN) in G3Pix AxonASP.

## Syntax

```asp
result = obj.Open(driver, dsn)
```

## Parameters and Arguments

- **driver** (String, Required): The name of the database driver. Supported values include "mysql", "postgres", "mssql", "sqlite", and "oracle".
- **dsn** (String, Required): The driver-specific connection string containing credentials, host, and database name.

## Return Values

Returns a **Boolean** value. It returns **True** if the connection was established and verified successfully, and **False** if the connection failed or if a connection is already open.

## Remarks

- The method performs an internal 5-second ping to verify the database availability before returning.
- If the connection fails, the error message can be retrieved using the **LastError** property or **GetError** method.
- Supported driver aliases:
    - MySQL: "mysql", "mariadb"
    - PostgreSQL: "postgres", "postgresql", "pgsql"
    - MS SQL Server: "mssql", "sqlserver"
    - SQLite: "sqlite", "sqlite3"
    - Oracle: "oracle", "ora", "oci"

## Code Example

```asp
<%
Dim db, isConnected
Set db = Server.CreateObject("G3DB")

' Example for MySQL
isConnected = db.Open("mysql", "user:password@tcp(127.0.0.1:3306)/my_database")

If isConnected Then
    Response.Write "Database connected successfully."
    db.Close
Else
    Response.Write "Connection failed: " & db.LastError
End If

Set db = Nothing
%>
```
