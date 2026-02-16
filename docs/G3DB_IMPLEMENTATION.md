# G3DB Implementation Guide

## Overview

G3DB is a powerful database library for AxonASP that provides direct access to Go's `database/sql` functionality with full VBScript compatibility. It supports multiple database systems including MySQL, PostgreSQL, MS SQL Server, and SQLite.

## Key Features

- **Multiple Database Support**: MySQL, PostgreSQL, MS SQL Server, SQLite
- **Full VBScript Compatibility**: Automatic type conversion between Go and VBScript
- **Connection Pooling**: Configurable connection pool with statistics
- **Prepared Statements**: Secure parameterized queries
- **Transaction Support**: Full transaction control with commit/rollback
- **ResultSet Navigation**: Iterate through query results with familiar methods
- **Environment Configuration**: Load connection settings from .env file
- **Automatic Cleanup**: Resources are automatically released at script end

## Creating a G3DB Object

```vbscript
Dim db
Set db = Server.CreateObject("G3DB")
```

Alternative shorthand:
```vbscript
Set db = Server.CreateObject("DB")
```

## Connection Methods

### Open(driver, dsn)
Opens a database connection with the specified driver and connection string.

**Parameters:**
- `driver` (String): Database driver - "mysql", "postgres", "mssql", "sqlite"
- `dsn` (String): Data Source Name (connection string)

**Returns:** Boolean - True if successful, False otherwise

**Example - SQLite:**
```vbscript
db.Open("sqlite", "./database.db")
db.Open("sqlite", ":memory:")  ' In-memory database
```

**Example - MySQL:**
```vbscript
db.Open("mysql", "user:password@tcp(localhost:3306)/dbname?parseTime=true")
```

**Example - PostgreSQL:**
```vbscript
db.Open("postgres", "host=localhost port=5432 user=postgres password=pass dbname=test sslmode=disable")
```

**Example - MS SQL Server:**
```vbscript
db.Open("mssql", "server=localhost;port=1433;user id=sa;password=pass;database=test")
```

### OpenFromEnv([driver])
Opens a database connection using settings from the .env file.

**Parameters:**
- `driver` (String, Optional): Database driver - defaults to "mysql"

**Returns:** Boolean - True if successful, False otherwise

**Environment Variables Required:**
```env
# MySQL
MYSQL_HOST=localhost
MYSQL_PORT=3306
MYSQL_USER=root
MYSQL_PASS=password
MYSQL_DATABASE=dbname

# PostgreSQL
POSTGRES_HOST=localhost
POSTGRES_PORT=5432
POSTGRES_USER=postgres
POSTGRES_PASS=password
POSTGRES_DATABASE=dbname
POSTGRES_SSLMODE=disable

# MS SQL Server
MSSQL_HOST=localhost
MSSQL_PORT=1433
MSSQL_USER=sa
MSSQL_PASS=password
MSSQL_DATABASE=dbname

# SQLite
SQLITE_PATH=./database.db
```

**Example:**
```vbscript
' Uses MySQL configuration from .env
db.OpenFromEnv("mysql")

' Uses PostgreSQL configuration from .env
db.OpenFromEnv("postgres")

' Uses SQLite configuration from .env
db.OpenFromEnv("sqlite")
```

### Close()
Closes the database connection and releases resources.

**Returns:** Boolean - True if successful, False otherwise

**Example:**
```vbscript
db.Close()
```

## Query Methods

### Query(sql, [params...])
Executes a SELECT query and returns a ResultSet object.

**Parameters:**
- `sql` (String): SQL query with `?` placeholders for parameters
- `params...` (Variant): Optional parameters to bind to placeholders

**Returns:** G3DBResultSet object or Null on error

**Example:**
```vbscript
' Simple query
Set rs = db.Query("SELECT * FROM users")

' Parameterized query
Set rs = db.Query("SELECT * FROM users WHERE age > ?", 25)

' Multiple parameters
Set rs = db.Query("SELECT * FROM users WHERE age > ? AND city = ?", 25, "New York")
```

### QueryRow(sql, [params...])
Executes a query that returns a single row.

**Parameters:**
- `sql` (String): SQL query with `?` placeholders
- `params...` (Variant): Optional parameters to bind

**Returns:** G3DBRow object or Null on error

**Example:**
```vbscript
' Get single row
Set row = db.QueryRow("SELECT name, email FROM users WHERE id = ?", 1)
If Not IsNull(row) Then
    values = row.Scan(2)  ' 2 columns
    Response.Write "Name: " & values(0) & "<br>"
    Response.Write "Email: " & values(1) & "<br>"
End If

' Scan into dictionary
Set row = db.QueryRow("SELECT name, email FROM users WHERE id = ?", 1)
Set dict = row.ScanMap("name", "email")
Response.Write dict("name")
```

### Exec(sql, [params...])
Executes an INSERT, UPDATE, or DELETE statement.

**Parameters:**
- `sql` (String): SQL statement with `?` placeholders
- `params...` (Variant): Optional parameters to bind

**Returns:** G3DBResult object with LastInsertId and RowsAffected properties, or Null on error

**Example:**
```vbscript
' Insert
Set result = db.Exec("INSERT INTO users (name, email, age) VALUES (?, ?, ?)", "John Doe", "john@example.com", 30)
If Not IsNull(result) Then
    Response.Write "Last Insert ID: " & result.LastInsertId & "<br>"
    Response.Write "Rows Affected: " & result.RowsAffected & "<br>"
End If

' Update
Set result = db.Exec("UPDATE users SET age = ? WHERE name = ?", 31, "John Doe")
Response.Write "Updated " & result.RowsAffected & " rows<br>"

' Delete
Set result = db.Exec("DELETE FROM users WHERE id = ?", 5)
Response.Write "Deleted " & result.RowsAffected & " rows<br>"
```

## Prepared Statements

### Prepare(sql)
Prepares a SQL statement for repeated execution.

**Parameters:**
- `sql` (String): SQL statement with `?` placeholders

**Returns:** G3DBStatement object or Null on error

**Example:**
```vbscript
Set stmt = db.Prepare("INSERT INTO users (name, email) VALUES (?, ?)")
If Not IsNull(stmt) Then
    stmt.Exec("Alice", "alice@example.com")
    stmt.Exec("Bob", "bob@example.com")
    stmt.Exec("Carol", "carol@example.com")
    stmt.Close()
End If
```

### PrepareContext(timeout, sql)
Prepares a SQL statement with a timeout context.

**Parameters:**
- `timeout` (Integer): Timeout in seconds
- `sql` (String): SQL statement with `?` placeholders

**Returns:** G3DBStatement object or Null on error

**Example:**
```vbscript
' Prepare with 5 second timeout
Set stmt = db.PrepareContext(5, "SELECT * FROM large_table WHERE id = ?")
Set rs = stmt.Query(1000)
stmt.Close()
```

## G3DBStatement Methods

### Query([params...])
Executes the prepared statement as a query.

**Returns:** G3DBResultSet object

### QueryRow([params...])
Executes the prepared statement returning a single row.

**Returns:** G3DBRow object

### Exec([params...])
Executes the prepared statement as a command.

**Returns:** G3DBResult object

### Close()
Closes the prepared statement.

## Transaction Support

### Begin()
Starts a new transaction.

**Returns:** G3DBTransaction object or Null on error

**Example:**
```vbscript
Set tx = db.Begin()
If Not IsNull(tx) Then
    tx.Exec("INSERT INTO users (name) VALUES (?)", "Test User")
    tx.Exec("UPDATE accounts SET balance = balance - 100 WHERE user_id = ?", 1)
    
    If tx.Commit() Then
        Response.Write "Transaction committed successfully"
    Else
        Response.Write "Failed to commit"
    End If
End If
```

### BeginTx(timeout, readOnly)
Starts a transaction with options.

**Parameters:**
- `timeout` (Integer): Timeout in seconds (0 for no timeout)
- `readOnly` (Boolean): True for read-only transaction

**Returns:** G3DBTransaction object or Null on error

**Example:**
```vbscript
' Read-only transaction with 10 second timeout
Set tx = db.BeginTx(10, True)
Set rs = tx.Query("SELECT * FROM users")
tx.Commit()
```

## G3DBTransaction Methods

### Commit()
Commits the transaction.

**Returns:** Boolean - True if successful

### Rollback()
Rolls back the transaction.

**Returns:** Boolean - True if successful

### Query(sql, [params...])
Executes a query within the transaction.

### QueryRow(sql, [params...])
Executes a single-row query within the transaction.

### Exec(sql, [params...])
Executes a command within the transaction.

### Prepare(sql)
Prepares a statement within the transaction.

**Example:**
```vbscript
Set tx = db.Begin()

' Insert data
tx.Exec("INSERT INTO orders (user_id, total) VALUES (?, ?)", 1, 100.50)

' Query within transaction
Set rs = tx.Query("SELECT * FROM orders WHERE user_id = ?", 1)
Do While Not rs.EOF
    Response.Write rs("total") & "<br>"
    rs.MoveNext()
Loop
rs.Close()

' Rollback if needed
If someCondition Then
    tx.Rollback()
Else
    tx.Commit()
End If
```

## G3DBResultSet Object

### Properties

- `EOF` (Boolean): True if past the last record
- `BOF` (Boolean): True if before the first record
- `Fields` (FieldsCollection): Collection of field objects

### Methods

#### MoveNext()
Moves to the next record.

**Returns:** Boolean - True if successful

#### Close()
Closes the result set.

#### GetRows()
Returns all remaining rows as an array.

**Returns:** VBArray of dictionaries

**Example:**
```vbscript
Set rs = db.Query("SELECT name, age FROM users")

' Access current row fields
Response.Write rs("name") & " - " & rs("age") & "<br>"

' Iterate through results
Do While Not rs.EOF
    Response.Write rs("name") & " is " & rs("age") & " years old<br>"
    rs.MoveNext()
Loop

' Get all rows as array
Set rs = db.Query("SELECT * FROM users")
allRows = rs.GetRows()
For Each row In allRows
    Response.Write row("name") & "<br>"
Next

rs.Close()
```

#### Columns()
Returns an array of column names.

**Example:**
```vbscript
Set rs = db.Query("SELECT * FROM users")
columns = rs.Columns()
For Each col In columns
    Response.Write col & "<br>"
Next
```

### Accessing Fields

```vbscript
Set rs = db.Query("SELECT id, name, email FROM users")
If Not rs.EOF Then
    ' Direct field access
    Response.Write rs("name") & "<br>"
    
    ' Via Fields collection
    Set fields = rs.Fields
    For i = 0 To fields.Count - 1
        Set field = fields.Item(i)
        Response.Write field.Name & " = " & field.Value & "<br>"
    Next
End If
rs.Close()
```

## Connection Pool Configuration

### SetMaxOpenConns(n)
Sets the maximum number of open connections.

**Example:**
```vbscript
db.SetMaxOpenConns(25)
```

### SetMaxIdleConns(n)
Sets the maximum number of idle connections.

**Example:**
```vbscript
db.SetMaxIdleConns(10)
```

### SetConnMaxLifetime(seconds)
Sets the maximum lifetime of connections in seconds.

**Example:**
```vbscript
db.SetConnMaxLifetime(3600)  ' 1 hour
```

### SetConnMaxIdleTime(seconds)
Sets the maximum idle time for connections in seconds.

**Example:**
```vbscript
db.SetConnMaxIdleTime(600)  ' 10 minutes
```

### Stats()
Returns connection pool statistics as a Dictionary.

**Returns:** Dictionary with statistics

**Example:**
```vbscript
Set stats = db.Stats()
Response.Write "Open Connections: " & stats("OpenConnections") & "<br>"
Response.Write "In Use: " & stats("InUse") & "<br>"
Response.Write "Idle: " & stats("Idle") & "<br>"
Response.Write "Wait Count: " & stats("WaitCount") & "<br>"
Response.Write "Max Open: " & stats("MaxOpenConnections") & "<br>"
```

## Error Handling

### Properties

- `LastError` (String): Description of the last error
- `IsOpen` (Boolean): Connection status

**Example:**
```vbscript
If Not db.Open("mysql", connectionString) Then
    Response.Write "Error: " & db.LastError & "<br>"
    Response.End()
End If

Set result = db.Exec("INSERT INTO users (name) VALUES (?)", "Test")
If IsNull(result) Then
    Response.Write "Insert failed: " & db.LastError & "<br>"
End If
```

## Complete Examples

### Example 1: Basic CRUD Operations

```vbscript
<%
Dim db
Set db = Server.CreateObject("G3DB")

' Connect to SQLite
If Not db.Open("sqlite", ":memory:") Then
    Response.Write "Connection failed: " & db.LastError
    Response.End()
End If

' Create table
db.Exec("CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT, email TEXT, age INTEGER)")

' Insert records
Set result = db.Exec("INSERT INTO users (name, email, age) VALUES (?, ?, ?)", "Alice", "alice@example.com", 28)
Response.Write "Inserted ID: " & result.LastInsertId & "<br>"

db.Exec("INSERT INTO users (name, email, age) VALUES (?, ?, ?)", "Bob", "bob@example.com", 35)
db.Exec("INSERT INTO users (name, email, age) VALUES (?, ?, ?)", "Carol", "carol@example.com", 42)

' Query records
Set rs = db.Query("SELECT * FROM users WHERE age > ?", 30)
Response.Write "<h3>Users older than 30:</h3>"
Do While Not rs.EOF
    Response.Write rs("name") & " (" & rs("age") & ")<br>"
    rs.MoveNext()
Loop
rs.Close()

' Update record
Set result = db.Exec("UPDATE users SET age = ? WHERE name = ?", 29, "Alice")
Response.Write "Updated " & result.RowsAffected & " rows<br>"

' Delete record
Set result = db.Exec("DELETE FROM users WHERE name = ?", "Bob")
Response.Write "Deleted " & result.RowsAffected & " rows<br>"

db.Close()
%>
```

### Example 2: Using Transactions

```vbscript
<%
Dim db, tx
Set db = Server.CreateObject("G3DB")
db.Open("sqlite", "./mydb.db")

' Start transaction
Set tx = db.Begin()

' Execute multiple operations
tx.Exec("INSERT INTO accounts (user_id, balance) VALUES (?, ?)", 1, 1000)
tx.Exec("INSERT INTO accounts (user_id, balance) VALUES (?, ?)", 2, 500)
tx.Exec("UPDATE accounts SET balance = balance - 100 WHERE user_id = ?", 1)
tx.Exec("UPDATE accounts SET balance = balance + 100 WHERE user_id = ?", 2)

' Commit or rollback
If someErrorCondition Then
    tx.Rollback()
    Response.Write "Transaction rolled back"
Else
    tx.Commit()
    Response.Write "Transaction committed"
End If

db.Close()
%>
```

### Example 3: Prepared Statements for Batch Insert

```vbscript
<%
Dim db, stmt
Set db = Server.CreateObject("G3DB")
db.Open("sqlite", ":memory:")

db.Exec("CREATE TABLE products (id INTEGER PRIMARY KEY, name TEXT, price REAL)")

' Prepare statement
Set stmt = db.Prepare("INSERT INTO products (name, price) VALUES (?, ?)")

' Execute multiple times
stmt.Exec("Widget A", 19.99)
stmt.Exec("Widget B", 29.99)
stmt.Exec("Widget C", 39.99)

stmt.Close()

' Query results
Set rs = db.Query("SELECT * FROM products")
Do While Not rs.EOF
    Response.Write rs("name") & ": $" & rs("price") & "<br>"
    rs.MoveNext()
Loop
rs.Close()

db.Close()
%>
```

### Example 4: Using Environment Configuration

```vbscript
<%
' In Global.asa - Application_OnStart
Sub Application_OnStart
    ' Open database connection for application scope
    Dim db
    Set db = Server.CreateObject("G3DB")
    
    If db.OpenFromEnv("mysql") Then
        ' Configure connection pool
        db.SetMaxOpenConns(20)
        db.SetMaxIdleConns(5)
        db.SetConnMaxLifetime(3600)
        
        ' Store in Application scope
        Set Application("Database") = db
    Else
        ' Log error
        Application("DatabaseError") = db.LastError
    End If
End Sub

Sub Application_OnEnd
    ' Clean up
    If IsObject(Application("Database")) Then
        Application("Database").Close()
    End If
End Sub
%>

<%
' In your ASP pages
Dim db
Set db = Application("Database")

If IsObject(db) And db.IsOpen Then
    Set rs = db.Query("SELECT * FROM users")
    ' Process results...
    rs.Close()
Else
    Response.Write "Database not available: " & Application("DatabaseError")
End If
%>
```

### Example 5: Using GetRows for Data Export

```vbscript
<%
Dim db, rs, allRows
Set db = Server.CreateObject("G3DB")
db.Open("sqlite", "./data.db")

Set rs = db.Query("SELECT id, name, email FROM users")
allRows = rs.GetRows()
rs.Close()

' Process as JSON
Dim json
Set json = Server.CreateObject("G3JSON")
Response.ContentType = "application/json"
Response.Write json.Stringify(allRows)

db.Close()
%>
```

## Best Practices

1. **Always Close Resources**: While automatic cleanup occurs at script end, explicitly closing connections, result sets, and statements is good practice:
   ```vbscript
   rs.Close()
   stmt.Close()
   db.Close()
   ```

2. **Use Prepared Statements**: For repeated queries or user input, always use prepared statements to prevent SQL injection:
   ```vbscript
   Set stmt = db.Prepare("SELECT * FROM users WHERE email = ?")
   Set rs = stmt.Query(userEmail)
   ```

3. **Connection Pooling**: For application-level connections, configure appropriate pool settings:
   ```vbscript
   db.SetMaxOpenConns(25)
   db.SetMaxIdleConns(10)
   db.SetConnMaxLifetime(3600)
   ```

4. **Error Handling**: Always check return values and LastError:
   ```vbscript
   If Not db.Open("mysql", connString) Then
       Response.Write "Error: " & db.LastError
       Response.End()
   End If
   ```

5. **Use Transactions**: For multiple related operations, use transactions to ensure data consistency:
   ```vbscript
   Set tx = db.Begin()
   ' Multiple operations
   If success Then
       tx.Commit()
   Else
       tx.Rollback()
   End If
   ```

6. **Environment Configuration**: Store database credentials in .env file and use OpenFromEnv():
   ```vbscript
   db.OpenFromEnv("mysql")
   ```

## Supported Database Drivers

| Driver | Identifier | Default Port |
|--------|-----------|-------------|
| MySQL/MariaDB | `"mysql"` | 3306 |
| PostgreSQL | `"postgres"` or `"postgresql"` | 5432 |
| MS SQL Server | `"mssql"` or `"sqlserver"` | 1433 |
| SQLite | `"sqlite"` or `"sqlite3"` | N/A |

## Type Conversions

G3DB automatically converts between Go and VBScript types:

| Go Type | VBScript Type |
|---------|---------------|
| nil | Null |
| bool | Boolean |
| int, int32, int64 | Integer/Long |
| float32, float64 | Double |
| string | String |
| []byte | String |
| time.Time | Date |
| []interface{} | Array |
| map[string]interface{} | Dictionary |

## Performance Tips

1. **Reuse Connections**: Store database connections in Application scope for reuse across requests
2. **Use Prepared Statements**: For repeated queries, prepare once and execute multiple times
3. **Configure Pool Size**: Adjust connection pool based on your application's concurrency needs
4. **Use Transactions**: Batch multiple operations in a single transaction for better performance
5. **Close Early**: Close result sets and statements as soon as you're done with them
6. **Index Your Tables**: Proper database indexing is crucial for query performance

## Thread Safety

G3DB is thread-safe and can be safely used across multiple concurrent requests. Connection pooling is handled automatically by the underlying Go database/sql package.

## Limitations

1. No support for stored procedures (use ADODB.Connection for that)
2. Advanced features like cursors and bulk operations require direct SQL
3. Named parameters not supported (use `?` placeholders instead)

## See Also

- [ADODB_IMPLEMENTATION.md](ADODB_IMPLEMENTATION.md) - Traditional ADO database access
- [CUSTOM_FUNCTIONS.md](CUSTOM_FUNCTIONS.md) - Custom VBScript functions
- [G3JSON_IMPLEMENTATION.md](G3JSON_IMPLEMENTATION.md) - JSON handling for data export
