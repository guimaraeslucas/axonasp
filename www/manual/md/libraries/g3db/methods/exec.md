# Exec Method

## Overview

The **Exec** method executes a SQL statement that does not return data rows, such as INSERT, UPDATE, or DELETE, in G3Pix AxonASP.

## Syntax

```asp
Set result = obj.Exec(sql [, params...])
```

## Parameters and Arguments

- **sql** (String, Required): The SQL statement to be executed.
- **params** (Variant, Optional): One or more values to be used as parameters in the SQL statement, replacing the `?` placeholders.

## Return Values

Returns a **G3DBResult** object. This object contains metadata about the operation, such as the number of rows affected and any last inserted ID.

## Remarks

- This method is designed for data modification operations where a record set is not expected.
- It automatically rewrites the `?` placeholders into the format required by the current database driver.
- The returned **G3DBResult** object provides **LastInsertId** and **RowsAffected** properties or methods.
- If the operation fails, the method returns an **Empty** value, and the error can be retrieved using the **LastError** property.

## Code Example

```asp
<%
Dim db, res
Set db = Server.CreateObject("G3DB")

If db.Open("mysql", "user:pass@tcp(localhost)/dbname") Then
    ' Execute an INSERT statement
    Set res = db.Exec("INSERT INTO users (username, status) VALUES (?, ?)", "john_doe", "active")

    If Not IsEmpty(res) Then
        Response.Write "Inserted ID: " & res.LastInsertId & "<br>"
        Response.Write "Rows affected: " & res.RowsAffected
    Else
        Response.Write "Error executing query: " & db.LastError
    End If

    db.Close
End If

Set db = Nothing
%>
```
