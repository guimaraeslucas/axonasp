# Use the G3DB Library

## Overview
Use **G3DB** to access SQL databases through the native G3Pix AxonASP runtime. The library manages a connection pool, supports parameterized SQL, and exposes result, statement, row, and transaction objects for structured workflows.

## Prerequisites
- Reachable database server or local SQLite file path.
- Valid driver and connection string, or configured environment-based settings.
- Create the object with the primary ProgID:

```asp
Dim db
Set db = Server.CreateObject("G3DB")
```
```javascript
var db = Server.CreateObject("G3DB");
```

## How It Works
- `Open` and `OpenFromEnv` initialize and validate a pooled connection.
- `Query`, `QueryRow`, `Exec`, and `Prepare` execute SQL against the current connection.
- `Begin` and `BeginTx` create transaction handles.
- Placeholder rewriting converts `?` markers to the active driver format when needed.
- Errors are stored in `LastError` and returned by `GetError`.

## API Reference

### Methods
- **Begin()**: Returns a `G3DBTransaction` object on success, otherwise **Empty**.
- **BeginTx([timeoutSeconds, readOnly])**: Returns a `G3DBTransaction` object on success, otherwise **Empty**.
- **Close()**: Returns **True** when close succeeds or when connection is already closed; returns **False** on close error.
- **Exec(sql[, params...])**: Returns a `G3DBResult` object on success, otherwise **Empty**.
- **GetError()**: Returns the current error string.
- **Open(driver, dsn)**: Returns **True** on successful open; **False** otherwise.
- **OpenFromEnv([driver])**: Returns **True** on successful open from config/env; **False** otherwise.
- **Prepare(sql)**: Returns a `G3DBStatement` object on success, otherwise **Empty**.
- **Query(sql[, params...])**: Returns a `G3DBResultSet` object on success, otherwise **Empty**.
- **QueryRow(sql[, params...])**: Returns a `G3DBRow` object on success, otherwise **Empty**.
- **SetConnMaxIdleTime(seconds)**: Returns **Empty**.
- **SetConnMaxLifetime(seconds)**: Returns **Empty**.
- **SetMaxIdleConns(count)**: Returns **Empty**.
- **SetMaxOpenConns(count)**: Returns **Empty**.
- **Stats()**: Returns a `Scripting.Dictionary` when connection is open; otherwise **Empty**.

### Properties
- **Driver** (String): Read/write.
- **DSN** (String): Read/write.
- **IsOpen** (Boolean): Read-only.
- **LastError** (String): Read-only.

## Example
```asp
<%
Dim db, rs
Set db = Server.CreateObject("G3DB")

If db.Open("mysql", "user:pass@tcp(127.0.0.1:3306)/appdb") Then
    Set rs = db.Query("SELECT id, username FROM users WHERE active = ?", 1)
    If Not IsEmpty(rs) Then
        Do While Not rs.EOF
            Response.Write rs("id") & " - " & rs("username") & "<br>"
            rs.MoveNext
        Loop
        rs.Close
    End If
    db.Close
Else
    Response.Write db.LastError
End If

Set db = Nothing
%>
```
