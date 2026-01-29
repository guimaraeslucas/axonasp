## ADODB.Connection - Access Database Support

### Overview

The AxonASP ADODB.Connection object now supports direct connections to Microsoft Access databases (both Jet and ACE formats) on Windows systems using OLE/OLEDB technology.

### Supported Formats

#### Microsoft Jet OLEDB 4.0 (Older Access)
```
Provider=Microsoft.Jet.OLEDB.4.0;Data Source=<path_to_database.mdb>
```

**Example:**
```asp
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../db/sample.mdb")
conn.Open
```

#### Microsoft ACE OLEDB 12.0 (Newer Access)
```
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=<path_to_database.accdb>
```

**Example:**
```asp
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("../db/mydata.accdb")
conn.Open
```

### Platform Support

Access database support **requires Windows**. When running on non-Windows platforms (Linux, macOS), attempting to connect to an Access database will:

1. Log a warning message to the console:
   ```
   Warning: Direct Access database support is only available on Windows. 
   Please use a different database system for cross-platform compatibility.
   ```

2. Fail gracefully with a null connection (State = 0)

### Usage Pattern

```asp
<%
Dim conn, rs, sql

' Create and open connection
Set conn = Server.CreateObject("ADODB.Connection")
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../db/sample.mdb")
conn.Open

' Check if connection opened successfully
If conn.State = 1 Then
    ' Execute a query
    Set rs = conn.Execute("SELECT * FROM Users WHERE ID = 1")
    
    ' Process results
    If Not rs.EOF Then
        response.write rs("UserName")
    End If
    
    ' Close recordset and connection
    rs.Close
    conn.Close
Else
    response.write "Connection failed"
End If

Set rs = Nothing
Set conn = Nothing
%>
```

### Implementation Details

**Windows Support:**
- Uses COM/OLE interface to ADODB.Connection
- Delegates to system OLEDB drivers (Jet or ACE)
- Supports both 32-bit (.mdb) and 64-bit (.accdb) Access formats

**Architecture:**
- Detection occurs in `openDatabase()` method
- Microsoft.Jet.OLEDB and Microsoft.ACE.OLEDB providers trigger OLE/COM path
- Stores OLE connection object in `oleConnection` field
- Other databases continue using standard SQL drivers (MySQL, PostgreSQL, SQLite, MSSQL)

**Error Handling:**
- Non-Windows platforms log warning to console
- Failed COM object creation results in silent failure (null connection)
- Failed database open results in silent failure

### Compatibility Notes

- **Data Types:** Native VBScript/ASP type conversion maintained
- **Path Resolution:** Use `Server.MapPath()` to resolve relative paths
- **Transactions:** BeginTrans, CommitTrans, RollbackTrans supported
- **Recordsets:** Full ADODB.Recordset functionality available
- **Connection String:** Exact case-insensitive format as shown above required

### Performance Considerations

- Access databases use file-based locking
- Single file connection limit: ~255 concurrent users
- Suitable for small to medium applications
- For high-traffic sites, consider migrating to SQLite, MySQL, or PostgreSQL

### Testing

Test file: `www/tests/test_access_db.asp`

This test file validates:
1. Jet OLEDB 4.0 connection
2. ACE OLEDB 12.0 connection
3. Query execution on Access databases

Access the test at: `http://localhost:4050/tests/test_access_db.asp`

### Troubleshooting

**"Cannot open Access database" warning:**
- Verify file path is correct
- Check that Access database file exists
- Ensure IUSR has read permissions on the file

**COM errors on Windows:**
- Access database support requires Windows COM infrastructure
- Ensure Office/Access OLEDB drivers are installed
- For ACE format, ensure "Microsoft Access Database Engine" is installed

**Non-Windows platform warning:**
- This is expected behavior
- Migrate to SQLite, MySQL, or PostgreSQL for cross-platform support
- Update connection strings accordingly
