# Use the G3DB Library

## Overview
The **G3DB** library is a high-performance native database abstraction layer for G3Pix AxonASP. It provides a direct interface to Go's standard `database/sql` package, enabling efficient access to multiple database systems without the overhead of ADODB or external COM components. 

The library supports industry-standard database drivers, including:
- **MySQL / MariaDB**
- **PostgreSQL**
- **Microsoft SQL Server**
- **SQLite**
- **Oracle**

## Syntax
To instantiate the library, use the following syntax:
```asp
Set db = Server.CreateObject("G3DB")
```

## Prerequisites
Ensure that the target database server is accessible and that the appropriate connection parameters (host, port, user, password, database) are configured. For **SQLite**, the library requires write access to the database file path.

## How it Works
The G3DB object manages a pool of active database connections. When the **Open** method is called, the library establishes a connection and verifies it using a synchronous ping operation. Once open, you can execute queries that return forward-only results (**G3DBResultSet**), single rows (**G3DBRow**), or metadata for non-query commands (**G3DBResult**).

The library supports parameterized queries using the standard `?` placeholder, which it automatically rewrites to match the specific requirements of the underlying database driver (e.g., `$1` for PostgreSQL or `@p1` for SQL Server).

## API Reference

### Methods
- **Begin**: Starts a standard database transaction.
- **BeginTx**: Starts a transaction with optional timeout and read-only settings.
- **Close**: Closes the database connection pool and releases all resources.
- **Exec**: Executes a command (INSERT, UPDATE, DELETE) and returns a **G3DBResult** object.
- **GetError**: Retrieves the last error message recorded by the connection.
- **Open**: Opens a connection using an explicit driver name and DSN string.
- **OpenFromEnv**: Opens a connection using settings defined in `axonasp.toml` or environment variables.
- **Prepare**: Creates a **G3DBStatement** for efficient repeated execution.
- **Query**: Executes a SELECT query and returns a **G3DBResultSet**.
- **QueryRow**: Executes a SELECT query expected to return a single row.
- **SetConnMaxIdleTime**: Sets the maximum time a connection can remain idle before being closed.
- **SetConnMaxLifetime**: Sets the maximum time a connection can be reused.
- **SetMaxIdleConns**: Sets the maximum number of idle connections in the pool.
- **SetMaxOpenConns**: Sets the maximum number of simultaneous open connections.
- **Stats**: Returns a Dictionary containing runtime connection pool statistics.

### Properties
- **Driver**: Returns the name of the active database driver.
- **DSN**: Returns the connection string (Data Source Name) used to open the connection.
- **IsOpen**: Indicates whether the database connection is currently active.
- **LastError**: Returns the most recent error message encountered by the library.

## Code Example
The following example demonstrates how to connect to a MySQL database and retrieve a list of users.

```asp
<%
Dim db, rs, sql
Set db = Server.CreateObject("G3DB")

' Open connection to MySQL
If db.Open("mysql", "user:password@tcp(localhost:3306)/my_database?parseTime=true") Then
    sql = "SELECT id, username, email FROM users WHERE active = ?"
    
    ' Execute parameterized query
    Set rs = db.Query(sql, 1)
    
    Do While Not rs.EOF
        Response.Write "ID: " & rs("id") & " - User: " & rs("username") & "<br>"
        rs.MoveNext
    Loop
    
    rs.Close
    db.Close
Else
    Response.Write "Connection failed: " & db.LastError
End If

Set db = Nothing
%>
```
