# DSN Property

## Overview
Returns the Data Source Name (connection string) used to open the current database connection.

## Syntax
```asp
dsnString = db.DSN
```

## Return Values
Returns a **String** containing the active DSN. If the connection is not open, it returns an empty string.

## Remarks
The DSN contains the driver-specific parameters such as host, port, and database name. For security reasons, sensitive information like passwords may be masked depending on the underlying database driver's implementation.

## Code Example
```asp
<%
Dim db
Set db = Server.CreateObject("G3DB")
If db.Open("sqlite", "data.db") Then
    Response.Write "Active DSN: " & db.DSN
    db.Close
End If
Set db = Nothing
%>
```
