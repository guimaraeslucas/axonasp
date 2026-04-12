<%
'        AxonASP Server
'        Copyright (C) 2026 G3pix Ltda. All rights reserved.
'        
'        Developed by Lucas Guimarães - G3pix Ltda
'        Contact: https://g3pix.com.br/
'        Project URL: https://g3pix.com.br/axonasp
'        
'        This Source Code Form is subject to the terms of the Mozilla Public
'        License, v. 2.0. If a copy of the MPL was not distributed with this
'        file, You can obtain one at https://mozilla.org/MPL/2.0/.
'        
'        Attribution Notice:
'        If this software is used in other projects, the name "AxonASP Server"
'        must be cited in the documentation or "About" section.
'        
'        Contribution Policy:
'        Modifications to the core source code of AxonASP Server must be
'        made available under this same license terms.


Response.ContentType = "text/plain"

Dim tableName
tableName = Request.QueryString("table")

If tableName = "" Then
    Response.Write "ERROR: No table specified"
    Response.End
End If

Dim accessPath
accessPath = Session("AccessPath")

If accessPath = "" Then
    Response.Write "ERROR: Session expired or Access path missing"
    Response.End
End If

' Target Configuration
Dim dbType, connMode, dbHost, dbPort, dbName, dbUser, dbPass
dbType = Session("DbType")
connMode = Session("ConnMode")
dbHost = Session("DbHost")
dbPort = Session("DbPort")
dbName = Session("DbName")
dbUser = Session("DbUser")
dbPass = Session("DbPass")

' Returns the correctly-quoted identifier for the target database dialect.
' SQLite/MSSQL use [brackets], MySQL uses `backticks`, PostgreSQL uses "double quotes".
Function QuoteIdent(name)
    Select Case dbType
        Case "mysql"
            QuoteIdent = "`" & Replace(name, "`", "``") & "`"
        Case "postgres"
            QuoteIdent = Chr(34) & Replace(name, Chr(34), Chr(34) & Chr(34)) & Chr(34)
        Case Else ' sqlite, mssql
            QuoteIdent = "[" & Replace(name, "]", "]]") & "]"
    End Select
End Function

' Maps an ADO field type constant to the correct SQL type for the target database.
' adSmallInt=2, adInteger=3, adTinyInt=16, adUnsignedTinyInt=17, adUnsignedSmallInt=18,
' adUnsignedInt=19, adBigInt=20, adUnsignedBigInt=21, adBoolean=11,
' adSingle=4, adDouble=5, adDecimal=14, adNumeric=131,
' adDate=7, adDBDate=133, adDBTime=134, adDBTimeStamp=135
Function MapType(adoType)
    Dim t
    Select Case adoType
        Case 2, 3, 16, 17, 18, 19, 20, 21, 11
            t = "INTEGER"
        Case 4, 5, 14, 131
            t = "FLOAT"
        Case 7, 133, 134, 135
            t = "DATETIME"
        Case Else
            t = "TEXT"
    End Select

    Select Case dbType
        Case "mysql"
            If t = "INTEGER" Then t = "INT"
            If t = "FLOAT" Then t = "DOUBLE"
            If t = "TEXT" Then t = "LONGTEXT"
        Case "postgres"
            If t = "INTEGER" Then t = "INTEGER"
            If t = "FLOAT" Then t = "DOUBLE PRECISION"
            If t = "DATETIME" Then t = "TIMESTAMP"
            If t = "TEXT" Then t = "TEXT"
        Case "mssql"
            If t = "INTEGER" Then t = "INT"
            If t = "FLOAT" Then t = "FLOAT"
            If t = "DATETIME" Then t = "DATETIME"
            If t = "TEXT" Then t = "NVARCHAR(MAX)"
        Case "sqlite"
            If t = "FLOAT" Then t = "REAL"
            If t = "DATETIME" Then t = "TEXT"
    End Select

    MapType = t
End Function

Dim connAccess, connTarget
Set connAccess = Server.CreateObject("ADODB.Connection")
Set connTarget = Server.CreateObject("ADODB.Connection")

' Open Access
On Error Resume Next
connAccess.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessPath
connAccess.Open
If Err.Number <> 0 Then
    Err.Clear
    connAccess.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & accessPath
    connAccess.Open
End If

If Err.Number <> 0 Then
    Response.Write "ERROR: Could not open Access DB: " & Err.Description
    Response.End
End If

' Build Target Connection String
Dim targetConnStr
Dim obj, result
Set obj = Server.CreateObject("G3AXON.Functions")

If connMode = "env" Then
    Select Case dbType
        Case "sqlite"
            targetConnStr = "Driver={SQLite3};Data Source=" & obj.axgetconfig("g3db.sqlite_path")
        Case "mysql"
            targetConnStr = "Driver={MySQL};Server=" & obj.AxGetConfig("g3db.mysql_host") & ";Port=" & obj.AxGetConfig("g3db.mysql_port") & ";Database=" & obj.AxGetConfig("g3db.mysql_database") & ";uid=" & obj.AxGetConfig("g3db.mysql_user") & ";pwd=" & obj.AxGetConfig("g3db.mysql_pass")
        Case "postgres"
            targetConnStr = "Driver={PostgreSQL};Server=" & obj.AxGetConfig("g3db.postgres_host") & ";Port=" & obj.AxGetConfig("g3db.postgres_port") & ";Database=" & obj.AxGetConfig("g3db.postgres_database") & ";uid=" & obj.AxGetConfig("g3db.postgres_user") & ";pwd=" & obj.AxGetConfig("g3db.postgres_pass")
        Case "mssql"
            targetConnStr = "Provider=SQLOLEDB;Server=" & obj.AxGetConfig("g3db.mssql_host") & "," & obj.AxGetConfig("g3db.mssql_port") & ";Database=" & obj.AxGetConfig("g3db.mssql_database") & ";uid=" & obj.AxGetConfig("g3db.mssql_user") & ";pwd=" & obj.AxGetConfig("g3db.mssql_pass")
    End Select
Else
    Select Case dbType
        Case "sqlite"
            targetConnStr = "Driver={SQLite3};Data Source=" & Server.MapPath(dbName)
        Case "mysql"
            If dbPort = "" Then dbPort = "3306"
            targetConnStr = "Driver={MySQL};Server=" & dbHost & ";Port=" & dbPort & ";Database=" & dbName & ";uid=" & dbUser & ";pwd=" & dbPass
        Case "postgres"
            If dbPort = "" Then dbPort = "5432"
            targetConnStr = "Driver={PostgreSQL};Server=" & dbHost & ";Port=" & dbPort & ";Database=" & dbName & ";uid=" & dbUser & ";pwd=" & dbPass
        Case "mssql"
            If dbPort = "" Then dbPort = "1433"
            targetConnStr = "Provider=SQLOLEDB;Server=" & dbHost & "," & dbPort & ";Database=" & dbName & ";uid=" & dbUser & ";pwd=" & dbPass
    End Select
End If

' Open Target
Err.Clear
connTarget.ConnectionString = targetConnStr
connTarget.Open

If Err.Number <> 0 Then
    Response.Write "ERROR: Could not open Target DB (" & dbType & "): " & Err.Description & " | ConnStr: " & targetConnStr
    connAccess.Close
    Response.End
End If
On Error Goto 0

' Get schema from Access
Dim rsAccess
Set rsAccess = Server.CreateObject("ADODB.Recordset")

On Error Resume Next
' Opens via OLE ADODB.Recordset with adUseClient=3 in the VM (see adodbRecordsetOpen).
' adUseClient forces the Microsoft Cursor Service to materialise full column metadata
' even when the table is empty, so Fields.Count is always correct.
rsAccess.Open "SELECT * FROM [" & tableName & "]", connAccess, 1, 1
If Err.Number <> 0 Then
    Response.Write "ERROR: Could not read table " & tableName & ": " & Err.Description
    connAccess.Close
    connTarget.Close
    Response.End
End If

If rsAccess.Fields.Count = 0 Then
    Response.Write "ERROR: Table " & tableName & " has no readable columns in source provider."
    If rsAccess.State = 1 Then rsAccess.Close
    connAccess.Close
    connTarget.Close
    Response.End
End If

' Build CREATE TABLE using dialect-correct quoting and type mapping
Dim sqlCreate, i, field, fieldType, fieldName, TypeName
sqlCreate = "CREATE TABLE " & QuoteIdent(tableName) & " ("

For i = 0 To rsAccess.Fields.Count - 1
    Set field = rsAccess.Fields.Item(i)
    fieldName = field.Name
    fieldType = field.Type

    sqlCreate = sqlCreate & QuoteIdent(fieldName) & " " & MapType(fieldType)

    If i < rsAccess.Fields.Count - 1 Then
        sqlCreate = sqlCreate & ", "
    End If
Next
sqlCreate = sqlCreate & ")"

' Execute DROP TABLE with dialect-correct syntax, then CREATE TABLE
Err.Clear
Select Case dbType
    Case "sqlite", "postgres"
        connTarget.Execute "DROP TABLE IF EXISTS " & QuoteIdent(tableName)
    Case "mysql"
        connTarget.Execute "DROP TABLE IF EXISTS " & QuoteIdent(tableName)
    Case "mssql"
        connTarget.Execute "IF OBJECT_ID(N'" & Replace(tableName, "'", "''") & "', N'U') IS NOT NULL DROP TABLE " & QuoteIdent(tableName)
End Select

connTarget.Execute sqlCreate
If Err.Number <> 0 Then
    Response.Write "ERROR: CREATE TABLE failed: " & Err.Description & " | SQL: " & sqlCreate
    rsAccess.Close
    connAccess.Close
    connTarget.Close
    Response.End
End If

' Build INSERT with dialect-correct quoting.
' All supported ODBC drivers accept ? as positional placeholder.
Dim sqlInsert, placeholders, values
sqlInsert = "INSERT INTO " & QuoteIdent(tableName) & " ("
placeholders = ""

For i = 0 To rsAccess.Fields.Count - 1
    sqlInsert = sqlInsert & QuoteIdent(rsAccess.Fields.Item(i).Name)
    placeholders = placeholders & "?"
    If i < rsAccess.Fields.Count - 1 Then
        sqlInsert = sqlInsert & ", "
        placeholders = placeholders & ", "
    End If
Next
sqlInsert = sqlInsert & ") VALUES (" & placeholders & ")"

' Loop through records and insert
Dim Count
Count = 0

connTarget.BeginTrans

Do While Not rsAccess.EOF
    ReDim values(rsAccess.Fields.Count - 1)
    For i = 0 To rsAccess.Fields.Count - 1
        values(i) = rsAccess.Fields.Item(i).Value
    Next

    connTarget.Execute sqlInsert, values
    If Err.Number <> 0 Then
        ' Log error and continue? Or fail? Let's fail for now to be safe.
        connTarget.RollbackTrans
        Response.Write "ERROR: INSERT failed at record " & (Count + 1) & ": " & Err.Description
        rsAccess.Close
        connAccess.Close
        connTarget.Close
        Response.End
    End If

    Count = Count + 1
    rsAccess.MoveNext
Loop

connTarget.CommitTrans

rsAccess.Close
connAccess.Close
connTarget.Close

Response.Write "SUCCESS: Imported " & Count & " records from " & tableName & " to " & dbType
%>
