<%@ Language=VBScript %>
<%
' AxonASP Server - G3DB Library Test Suite
' Copyright (C) 2026 G3pix Ltda. All rights reserved.
Option Explicit

Response.Write "<!DOCTYPE html>" & vbCrLf
Response.Write "<html><head>" & vbCrLf
Response.Write "<meta charset='utf-8'>" & vbCrLf
Response.Write "<title>G3DB Library Test Suite</title>" & vbCrLf
Response.Write "<style>" & vbCrLf
Response.Write "body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }" & vbCrLf
Response.Write "h1 { color: #333; border-bottom: 3px solid #4CAF50; padding-bottom: 10px; }" & vbCrLf
Response.Write "h2 { color: #666; margin-top: 30px; border-bottom: 2px solid #2196F3; padding-bottom: 5px; }" & vbCrLf
Response.Write ".test { background: white; margin: 10px 0; padding: 15px; border-left: 4px solid #2196F3; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }" & vbCrLf
Response.Write ".pass { border-left-color: #4CAF50; }" & vbCrLf
Response.Write ".fail { border-left-color: #f44336; background: #ffebee; }" & vbCrLf
Response.Write ".label { font-weight: bold; color: #333; }" & vbCrLf
Response.Write ".result { margin-left: 20px; color: #666; }" & vbCrLf
Response.Write ".error { color: #f44336; font-weight: bold; }" & vbCrLf
Response.Write ".success { color: #4CAF50; font-weight: bold; }" & vbCrLf
Response.Write "pre { background: #f9f9f9; padding: 10px; border: 1px solid #ddd; overflow-x: auto; }" & vbCrLf
Response.Write "</style>" & vbCrLf
Response.Write "</head><body>" & vbCrLf
Response.Write "<h1>G3DB Library Test Suite</h1>" & vbCrLf

Dim testCount, passCount, failCount
testCount = 0
passCount = 0
failCount = 0

Sub StartTest(testName)
    Response.Write "<div class='test'>"
    Response.Write "<div class='label'>Test: " & testName & "</div>"
    testCount = testCount + 1
End Sub

Sub EndTest()
    Response.Write "</div>"
End Sub

Sub TestPass(message)
    Response.Write "<div class='result success'>✓ PASS: " & message & "</div>"
    passCount = passCount + 1
End Sub

Sub TestFail(message)
    Response.Write "<div class='result error'>✗ FAIL: " & message & "</div>"
    failCount = failCount + 1
End Sub

Sub TestInfo(message)
    Response.Write "<div class='result'>" & message & "</div>"
End Sub

' =============================================================================
' Test 1: Create G3DB object
' =============================================================================
Response.Write "<h2>1. Object Creation Tests</h2>"

StartTest("Create G3DB object")
On Error Resume Next
Dim db
Set db = Server.CreateObject("G3DB")
If Err.Number <> 0 Then
    TestFail("Cannot create G3DB object: " & Err.Description)
    EndTest()
    Response.Write "</body></html>"
    Response.End()
Else
    TestPass("G3DB object created successfully")
End If
On Error Goto 0
EndTest()

' =============================================================================
' Test 2: SQLite Connection (simplest, no external dependency)
' =============================================================================
Response.Write "<h2>2. SQLite Connection Tests</h2>"

StartTest("Open SQLite connection")
Dim sqliteDb
Set sqliteDb = Server.CreateObject("G3DB")
Dim openResult
openResult = sqliteDb.Open("sqlite", ":memory:")
If openResult Then
    TestPass("SQLite in-memory database opened")
    TestInfo("IsOpen: " & sqliteDb.IsOpen)
Else
    TestFail("Failed to open SQLite database: " & sqliteDb.LastError)
    EndTest()
    Response.Write "</body></html>"
    Response.End()
End If
EndTest()

' =============================================================================
' Test 3: Create Table and Insert Data
' =============================================================================
Response.Write "<h2>3. Table Creation and Data Insertion</h2>"

StartTest("Create users table")
Dim result
Set result = sqliteDb.Exec("CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT, email TEXT, age INTEGER)")
If Not IsNull(result) Then
    TestPass("Table created successfully")
    TestInfo("Rows affected: " & result.RowsAffected)
Else
    TestFail("Failed to create table: " & sqliteDb.LastError)
End If
EndTest()

StartTest("Insert test data")
Set result = sqliteDb.Exec("INSERT INTO users (name, email, age) VALUES ('Alice Smith', 'alice@example.com', 28)")
If Not IsNull(result) Then
    TestPass("Record inserted")
    TestInfo("Last Insert ID: " & result.LastInsertId)
    TestInfo("Rows affected: " & result.RowsAffected)
Else
    TestFail("Failed to insert: " & sqliteDb.LastError)
End If
EndTest()

StartTest("Insert multiple records")
Set result = sqliteDb.Exec("INSERT INTO users (name, email, age) VALUES ('Bob Johnson', 'bob@example.com', 35)")
Set result = sqliteDb.Exec("INSERT INTO users (name, email, age) VALUES ('Carol White', 'carol@example.com', 42)")
Set result = sqliteDb.Exec("INSERT INTO users (name, email, age) VALUES ('David Brown', 'david@example.com', 31)")
If Not IsNull(result) Then
    TestPass("Multiple records inserted")
Else
    TestFail("Failed to insert multiple records")
End If
EndTest()

' =============================================================================
' Test 4: Query and ResultSet operations
' =============================================================================
Response.Write "<h2>4. Query and ResultSet Tests</h2>"

StartTest("Query all users")
Dim rs
Set rs = sqliteDb.Query("SELECT * FROM users ORDER BY name")
If Not IsNull(rs) Then
    TestPass("Query executed successfully")
    TestInfo("EOF: " & rs.EOF & ", BOF: " & rs.BOF)
    
    Dim rowCount
    rowCount = 0
    Response.Write "<pre>"
    Do While Not rs.EOF
        rowCount = rowCount + 1
        Response.Write "Row " & rowCount & ": "
        Response.Write "ID=" & rs("id") & ", "
        Response.Write "Name=" & rs("name") & ", "
        Response.Write "Email=" & rs("email") & ", "
        Response.Write "Age=" & rs("age") & vbCrLf
        rs.MoveNext()
    Loop
    Response.Write "</pre>"
    TestInfo("Total rows: " & rowCount)
    rs.Close()
Else
    TestFail("Query failed: " & sqliteDb.LastError)
End If
EndTest()

StartTest("Query with parameterized WHERE clause")
Dim stmt
Set stmt = sqliteDb.Prepare("SELECT * FROM users WHERE age > ?")
If Not IsNull(stmt) Then
    TestPass("Statement prepared")
    
    Set rs = stmt.Query(30)
    If Not IsNull(rs) Then
        Response.Write "<pre>"
        Do While Not rs.EOF
            Response.Write "Name: " & rs("name") & ", Age: " & rs("age") & vbCrLf
            rs.MoveNext()
        Loop
        Response.Write "</pre>"
        TestPass("Parameterized query executed")
        rs.Close()
    Else
        TestFail("Failed to execute parameterized query")
    End If
    stmt.Close()
Else
    TestFail("Failed to prepare statement")
End If
EndTest()

StartTest("QueryRow for single record")
Dim row
Set row = sqliteDb.QueryRow("SELECT name, email FROM users WHERE id = 1")
If Not IsNull(row) Then
    Dim values
    values = row.Scan(2) ' 2 columns
    If Not IsNull(values) Then
        TestPass("QueryRow executed successfully")
        TestInfo("Name: " & values(0))
        TestInfo("Email: " & values(1))
    Else
        TestFail("Failed to scan row")
    End If
Else
    TestFail("QueryRow failed")
End If
EndTest()

' =============================================================================
' Test 5: Update and Delete operations
' =============================================================================
Response.Write "<h2>5. Update and Delete Tests</h2>"

StartTest("Update record")
Set result = sqliteDb.Exec("UPDATE users SET age = 29 WHERE name = 'Alice Smith'")
If Not IsNull(result) Then
    TestPass("Record updated")
    TestInfo("Rows affected: " & result.RowsAffected)
Else
    TestFail("Update failed: " & sqliteDb.LastError)
End If
EndTest()

StartTest("Delete record")
Set result = sqliteDb.Exec("DELETE FROM users WHERE name = 'David Brown'")
If Not IsNull(result) Then
    TestPass("Record deleted")
    TestInfo("Rows affected: " & result.RowsAffected)
Else
    TestFail("Delete failed: " & sqliteDb.LastError)
End If
EndTest()

' =============================================================================
' Test 6: Transaction support
' =============================================================================
Response.Write "<h2>6. Transaction Tests</h2>"

StartTest("Begin transaction and commit")
Dim tx
Set tx = sqliteDb.Begin()
If Not IsNull(tx) Then
    TestPass("Transaction started")
    
    Set result = tx.Exec("INSERT INTO users (name, email, age) VALUES ('Eve Adams', 'eve@example.com', 27)")
    Set result = tx.Exec("INSERT INTO users (name, email, age) VALUES ('Frank Miller', 'frank@example.com', 33)")
    
    If tx.Commit() Then
        TestPass("Transaction committed successfully")
    Else
        TestFail("Failed to commit transaction")
    End If
Else
    TestFail("Failed to begin transaction")
End If
EndTest()

StartTest("Begin transaction and rollback")
Set tx = sqliteDb.Begin()
If Not IsNull(tx) Then
    Set result = tx.Exec("INSERT INTO users (name, email, age) VALUES ('Should Not Exist', 'test@example.com', 99)")
    
    If tx.Rollback() Then
        TestPass("Transaction rolled back successfully")
        
        ' Verify rollback worked
        Set rs = sqliteDb.Query("SELECT * FROM users WHERE name = 'Should Not Exist'")
        If rs.EOF Then
            TestPass("Rollback verified - record does not exist")
        Else
            TestFail("Rollback failed - record still exists")
        End If
        rs.Close()
    Else
        TestFail("Failed to rollback transaction")
    End If
Else
    TestFail("Failed to begin transaction")
End If
EndTest()

' =============================================================================
' Test 7: GetRows functionality
' =============================================================================
Response.Write "<h2>7. GetRows Test</h2>"

StartTest("GetRows - fetch all results as array")
Set rs = sqliteDb.Query("SELECT name, age FROM users ORDER BY name")
If Not IsNull(rs) Then
    Dim allRows
    allRows = rs.GetRows()
    
    If IsArray(allRows) Then
        TestPass("GetRows returned array")
        TestInfo("Array size: " & UBound(allRows) + 1)
        
        Response.Write "<pre>"
        Dim i
        For i = 0 To UBound(allRows)
            Response.Write "Row " & (i + 1) & ": " & allRows(i)("name") & ", Age: " & allRows(i)("age") & vbCrLf
        Next
        Response.Write "</pre>"
    Else
        TestFail("GetRows did not return array")
    End If
    rs.Close()
Else
    TestFail("Query failed")
End If
EndTest()

' =============================================================================
' Test 8: Connection pool settings
' =============================================================================
Response.Write "<h2>8. Connection Pool Configuration Tests</h2>"

StartTest("Set connection pool parameters")
sqliteDb.SetMaxOpenConns(10)
sqliteDb.SetMaxIdleConns(5)
sqliteDb.SetConnMaxLifetime(3600)
sqliteDb.SetConnMaxIdleTime(600)
TestPass("Connection pool parameters set successfully")
EndTest()

StartTest("Get connection statistics")
Dim stats
Set stats = sqliteDb.Stats()
If Not IsNull(stats) Then
    TestPass("Connection statistics retrieved")
    TestInfo("Max Open Connections: " & stats.Item("MaxOpenConnections"))
    TestInfo("Open Connections: " & stats.Item("OpenConnections"))
    TestInfo("In Use: " & stats.Item("InUse"))
    TestInfo("Idle: " & stats.Item("Idle"))
Else
    TestFail("Failed to get connection statistics")
End If
EndTest()

' =============================================================================
' Test 9: Fields collection access
' =============================================================================
Response.Write "<h2>9. Fields Collection Tests</h2>"

StartTest("Access Fields collection")
Set rs = sqliteDb.Query("SELECT id, name, email, age FROM users LIMIT 1")
If Not IsNull(rs) And Not rs.EOF Then
    Dim fields
    Set fields = rs.Fields
    TestPass("Fields collection accessed")
    TestInfo("Field count: " & fields.Count)
    
    Response.Write "<pre>"
    Dim j
    For j = 0 To fields.Count - 1
        Dim fld
        Set fld = fields.Item(j)
        Response.Write "Field " & j & ": Name=" & fld.Name & ", Value=" & fld.Value & vbCrLf
    Next
    Response.Write "</pre>"
    rs.Close()
Else
    TestFail("Failed to access Fields collection")
End If
EndTest()

' =============================================================================
' Test 10: Close connection
' =============================================================================
Response.Write "<h2>10. Connection Cleanup Tests</h2>"

StartTest("Close database connection")
If sqliteDb.Close() Then
    TestPass("Database connection closed successfully")
    TestInfo("IsOpen: " & sqliteDb.IsOpen)
Else
    TestFail("Failed to close database connection")
End If
EndTest()

' =============================================================================
' Test 11: OpenFromEnv (if environment is configured)
' =============================================================================
Response.Write "<h2>11. Environment Configuration Tests</h2>"

StartTest("OpenFromEnv with SQLite")
Dim envDb
Set envDb = Server.CreateObject("G3DB")
openResult = envDb.OpenFromEnv("sqlite")
If openResult Then
    TestPass("Database opened from environment configuration")
    TestInfo("Driver: " & envDb.Driver)
    envDb.Close()
Else
    TestInfo("OpenFromEnv test skipped or failed: " & envDb.LastError)
    TestInfo("(This is expected if .env is not configured)")
End If
EndTest()

' =============================================================================
' Summary
' =============================================================================
Response.Write "<h2>Test Summary</h2>"
Response.Write "<div class='test'>"
Response.Write "<div class='label'>Total Tests: " & testCount & "</div>"
Response.Write "<div class='result success'>Passed: " & passCount & "</div>"
Response.Write "<div class='result error'>Failed: " & failCount & "</div>"
If failCount = 0 Then
    Response.Write "<div class='result success' style='font-size: 1.2em; margin-top: 10px;'>✓ ALL TESTS PASSED!</div>"
Else
    Response.Write "<div class='result error' style='font-size: 1.2em; margin-top: 10px;'>✗ SOME TESTS FAILED</div>"
End If
Response.Write "</div>"

Response.Write "</body></html>"
%>
