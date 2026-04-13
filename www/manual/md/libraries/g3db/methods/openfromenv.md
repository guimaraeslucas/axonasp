# OpenFromEnv Method

## Overview

The **OpenFromEnv** method establishes a connection to a database using settings from the G3Pix AxonASP configuration file (`axonasp.toml`) or environment variables.

## Syntax

```asp
result = obj.OpenFromEnv([driver])
```

## Parameters and Arguments

- **driver** (String, Optional): The name of the database driver to use. Defaults to "mysql" if not provided. Supported drivers include "mysql", "postgres", "mssql", "sqlite", and "oracle".

## Return Values

Returns a **Boolean** value. It returns **True** if the connection was established successfully, and **False** if the connection failed or if the required configuration settings are missing.

## Remarks

- Connection parameters such as host, port, user, password, and database name are read from the `[g3db]` section of the `axonasp.toml` configuration file.
- If environment variable support is enabled, these settings can be overridden by their corresponding environment variables.
- This method is useful for maintaining security and flexibility by separating connection details from the code.
- Like the **Open** method, it performs a ping to verify connectivity before returning.

## Code Example

```asp
<%
Dim db, isConnected
Set db = Server.CreateObject("G3DB")

' Attempt to open a connection using "mysql" settings from axonasp.toml
isConnected = db.OpenFromEnv("mysql")

If isConnected Then
    Response.Write "Database connected using environment configuration."
    db.Close
Else
    Response.Write "Failed to connect: " & db.LastError
End If

Set db = Nothing
%>
```
