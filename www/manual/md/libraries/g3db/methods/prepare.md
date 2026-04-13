# Prepare Method

## Overview

The **Prepare** method creates a reusable prepared statement in G3Pix AxonASP.

## Syntax

```asp
Set result = obj.Prepare(sql)
```

## Parameters and Arguments

- **sql** (String, Required): The SQL statement to be prepared, containing `?` placeholders for parameters.

## Return Values

Returns a **G3DBStatement** object. This object represents the prepared statement and can be used to execute queries multiple times with different parameters.

## Remarks

- Prepared statements are pre-compiled by the database, which can result in better performance for statements executed repeatedly.
- The method automatically rewrites the `?` placeholders for the current database driver during the preparation phase.
- The returned **G3DBStatement** object exposes **Query**, **QueryRow**, and **Exec** methods for execution.
- It is important to call the **Close** method on the **G3DBStatement** object to release the prepared statement resource from the database server.

## Code Example

```asp
<%
Dim db, stmt, rs, i
Set db = Server.CreateObject("G3DB")

If db.Open("sqlite", "mydb.db") Then
    ' Prepare a reusable statement
    Set stmt = db.Prepare("SELECT name FROM users WHERE id = ?")

    ' Execute the prepared statement multiple times
    For i = 1 To 3
        Set rs = stmt.Query(i)
        If Not rs.EOF Then
            Response.Write "User " & i & ": " & rs("name") & "<br>"
        End If
        rs.Close
    Next

    ' Properly close the statement
    stmt.Close
    db.Close
End If

Set db = Nothing
%>
```
