# Query Method

## Overview

The **Query** method executes a SQL query, such as a SELECT statement, and returns a result set object in G3Pix AxonASP.

## Syntax

```asp
Set result = obj.Query(sql [, params...])
```

## Parameters and Arguments

- **sql** (String, Required): The SQL query statement to be executed.
- **params** (Variant, Optional): One or more values to be used as parameters in the SQL statement, replacing the `?` placeholders.

## Return Values

Returns a **G3DBResultSet** object. This object provides a forward-only cursor to navigate and retrieve the data returned by the database.

## Remarks

- The method automatically rewrites the `?` placeholders into the format required by the current database driver (e.g., `$1, $2` for PostgreSQL or `@p1, @p2` for MS SQL Server).
- Parameterization is highly recommended to protect against SQL injection attacks.
- If the query fails, the method returns an **Empty** value, and the error description can be retrieved from the **LastError** property.
- The returned **G3DBResultSet** should be closed using its **Close** method when it is no longer needed.

## Code Example

```asp
<%
Dim db, rs
Set db = Server.CreateObject("G3DB")

If db.Open("mysql", "user:pass@tcp(localhost)/dbname") Then
    ' Simple query with parameters
    Set rs = db.Query("SELECT username, email FROM users WHERE id = ?", 123)

    If Not rs.EOF Then
        Response.Write "Username: " & rs("username") & "<br>"
        Response.Write "Email: " & rs("email")
    End If

    rs.Close
    db.Close
End If

Set db = Nothing
%>
```
