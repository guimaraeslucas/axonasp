# BeginTx Method

## Overview

The **BeginTx** method starts a database transaction with optional settings for timeout and read-only mode in G3Pix AxonASP.

## Syntax

```asp
Set result = obj.BeginTx([timeoutSeconds, readOnly])
```

## Parameters and Arguments

- **timeoutSeconds** (Integer, Optional): The maximum duration in seconds for the transaction before it automatically expires and becomes invalid.
- **readOnly** (Boolean, Optional): If set to **True**, the transaction will be optimized for read-only operations.

## Return Values

Returns a **G3DBTransaction** object. This object represents the active transaction and allows for grouping multiple operations into a single atomic unit.

## Remarks

- This method provides finer control over transaction behavior compared to the standard **Begin** method.
- Read-only transactions can reduce database locking and improve performance for SELECT operations.
- The transaction must be explicitly committed or rolled back using its respective methods.
- If the script execution ends without a commit, G3Pix AxonASP performs an automatic rollback.

## Code Example

```asp
<%
Dim db, tx, rs
Set db = Server.CreateObject("G3DB")

If db.Open("postgres", "host=localhost user=dbuser dbname=reports") Then
    ' Start a read-only transaction with a 30-second timeout
    Set tx = db.BeginTx(30, True)

    ' Execute queries within the transaction
    Set rs = tx.Query("SELECT id, value FROM statistics")
    
    ' Perform data processing...
    
    ' Explicitly close the transaction (Rollback is safe for read-only)
    tx.Rollback
    rs.Close
    db.Close
End If

Set db = Nothing
%>
```
