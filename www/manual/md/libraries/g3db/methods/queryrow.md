# QueryRow Method

## Overview

The **QueryRow** method executes a SQL query that is expected to return a single row in G3Pix AxonASP.

## Syntax

```asp
Set result = obj.QueryRow(sql [, params...])
```

## Parameters and Arguments

- **sql** (String, Required): The SQL query statement to be executed.
- **params** (Variant, Optional): One or more values to be used as parameters in the SQL statement, replacing the `?` placeholders.

## Return Values

Returns a **G3DBRow** object. This object represents the single row returned by the query and provides methods to retrieve its data.

## Remarks

- This method is designed for efficiency when you expect exactly one row, such as fetching a record by its primary key.
- If the query returns multiple rows, only the first row will be available through the **G3DBRow** object.
- If no row is found, the **Scan** method of the returned **G3DBRow** will return an **Empty** value.
- Parameterized queries with `?` are supported and are automatically rewritten for the current database driver.

## Code Example

```asp
<%
Dim db, row, userName
Set db = Server.CreateObject("G3DB")

If db.Open("postgres", "host=localhost user=myuser dbname=mydb") Then
    ' Fetch a single user's name
    Set row = db.QueryRow("SELECT name FROM users WHERE id = ?", 1)

    ' Retrieve the name using Scan
    userName = row.Scan()

    If Not IsEmpty(userName) Then
        Response.Write "User Name: " & userName
    Else
        Response.Write "User not found."
    End If

    db.Close
End If

Set db = Nothing
%>
```
