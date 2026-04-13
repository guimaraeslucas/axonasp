# LastError Property

## Overview
Returns the most recent error message encountered by the database connection object.

## Syntax
```asp
errMsg = db.LastError
```

## Return Values
Returns a **String** containing the description of the last error. If no error has occurred, it returns an empty string.

## Remarks
This property is essential for debugging failed **Open**, **Exec**, or **Query** operations. It provides the specific error message returned by the underlying database driver.

## Code Example
```asp
<%
Dim db
Set db = Server.CreateObject("G3DB")

If Not db.Open("mysql", "invalid_dsn") Then
    Response.Write "Connection failed: " & db.LastError
End If

Set db = Nothing
%>
```
