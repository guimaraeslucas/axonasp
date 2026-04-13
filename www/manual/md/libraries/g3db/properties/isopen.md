# IsOpen Property

## Overview
Indicates whether the database connection pool is currently open and active.

## Syntax
```asp
status = db.IsOpen
```

## Return Values
Returns a **Boolean** value. It returns **True** if the connection pool is open, and **False** if it is closed or has not yet been opened.

## Remarks
Use this property to verify the state of the database connection before executing queries or commands to avoid runtime errors.

## Code Example
```asp
<%
Dim db
Set db = Server.CreateObject("G3DB")
db.Open "mysql", "user:pass@tcp(localhost:3306)/db"

If db.IsOpen Then
    Response.Write "Database is connected"
Else
    Response.Write "Database is disconnected"
End If

db.Close
Set db = Nothing
%>
```
