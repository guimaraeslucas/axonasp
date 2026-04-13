# Begin Method

## Overview

The **Begin** method starts a new database transaction in G3Pix AxonASP.

## Syntax

```asp
Set result = obj.Begin()
```

## Parameters and Arguments

None. This method is also accessible through the aliases **BeginTrans** and **BeginTransaction**.

## Return Values

Returns a **G3DBTransaction** object. This object represents the active database transaction and provides methods to commit or roll back the changes.

## Remarks

- Transactions are essential for ensuring data integrity when performing multiple related database operations.
- All operations performed through the returned **G3DBTransaction** object are isolated until an explicit **Commit** is called.
- If the transaction is not committed by the time the script finishes execution, G3Pix AxonASP will automatically roll it back to prevent resource leaks and incomplete data updates.
- Use the **Rollback** method of the transaction object to explicitly cancel all pending changes.

## Code Example

```asp
<%
Dim db, tx
Set db = Server.CreateObject("G3DB")

If db.Open("mysql", "user:pass@tcp(localhost)/dbname") Then
    ' Start a transaction
    Set tx = db.Begin()

    On Error Resume Next
    tx.Exec "UPDATE accounts SET balance = balance - 100 WHERE id = 1"
    tx.Exec "UPDATE accounts SET balance = balance + 100 WHERE id = 2"

    If Err.Number = 0 Then
        tx.Commit
        Response.Write "Transaction completed successfully."
    Else
        tx.Rollback
        Response.Write "Transaction failed and was rolled back: " & Err.Description
    End If
    On Error GoTo 0

    db.Close
End If

Set db = Nothing
%>
```
