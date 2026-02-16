<%
    ' AxonASP Database Export Worker
    ' Imports a single table from Access to another modern database

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
    If connMode = "env" Then
        Select Case dbType
            Case "sqlite"
                targetConnStr = "Driver={SQLite3};Data Source=" & axgetenv("SQLITE_PATH")
            Case "mysql"
                targetConnStr = "Driver={MySQL};Server=" & axgetenv("MYSQL_HOST") & ";Port=" & axgetenv("MYSQL_PORT") & ";Database=" & axgetenv("MYSQL_DATABASE") & ";uid=" & axgetenv("MYSQL_USER") & ";pwd=" & axgetenv("MYSQL_PASS")
            Case "postgres"
                targetConnStr = "Driver={PostgreSQL};Server=" & axgetenv("POSTGRES_HOST") & ";Port=" & axgetenv("POSTGRES_PORT") & ";Database=" & axgetenv("POSTGRES_DATABASE") & ";uid=" & axgetenv("POSTGRES_USER") & ";pwd=" & axgetenv("POSTGRES_PASS")
            Case "mssql"
                targetConnStr = "Provider=SQLOLEDB;Server=" & axgetenv("MSSQL_HOST") & "," & axgetenv("MSSQL_PORT") & ";Database=" & axgetenv("MSSQL_DATABASE") & ";uid=" & axgetenv("MSSQL_USER") & ";pwd=" & axgetenv("MSSQL_PASS")
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
    On Error GoTo 0

    ' Get schema from Access
    Dim rsAccess
    Set rsAccess = Server.CreateObject("ADODB.Recordset")
    
    On Error Resume Next
    rsAccess.Open "SELECT * FROM [" & tableName & "]", connAccess, 1, 1 
    If Err.Number <> 0 Then
        Response.Write "ERROR: Could not read table " & tableName & ": " & Err.Description
        connAccess.Close
        connTarget.Close
        Response.End
    End If

    ' Build CREATE TABLE
    Dim sqlCreate, i, field, fieldType, fieldName, typeName
    sqlCreate = "CREATE TABLE [" & tableName & "] ("
    
    For i = 0 To rsAccess.Fields.Count - 1
        Set field = rsAccess.Fields.Item(i)
        fieldName = field.Name
        fieldType = field.Type
        
        sqlCreate = sqlCreate & "[" & fieldName & "] "
        
        ' Map types
        Select Case fieldType
            Case 2, 3, 16, 17, 18, 19, 20, 21, 11
                typeName = "INTEGER"
            Case 4, 5, 14, 131 ' adSingle, adDouble, adDecimal, adNumeric
                typeName = "FLOAT"
            Case 7, 133, 134, 135 ' adDate, adDBDate, adDBTime, adDBTimeStamp
                typeName = "DATETIME"
            Case Else
                typeName = "TEXT"
        End Select
        
        ' Database specific type adjustments
        If dbType = "mysql" Then
            If typeName = "INTEGER" Then typeName = "INT"
            ElseIf typeName = "TEXT" Then typeName = "VARCHAR(255)"
        End If
        
        sqlCreate = sqlCreate & typeName
        
        If i < rsAccess.Fields.Count - 1 Then
            sqlCreate = sqlCreate & ", "
        End If
    Next
    sqlCreate = sqlCreate & ")"

    ' Execute CREATE TABLE (handle existing table)
    Err.Clear
    If dbType = "sqlite" Or dbType = "mysql" Or dbType = "postgres" Then
        connTarget.Execute "DROP TABLE IF EXISTS [" & tableName & "]"
    Else ' MSSQL
        connTarget.Execute "IF OBJECT_ID('[" & tableName & "]', 'U') IS NOT NULL DROP TABLE [" & tableName & "]"
    End If
    
    connTarget.Execute sqlCreate
    If Err.Number <> 0 Then
        Response.Write "ERROR: CREATE TABLE failed: " & Err.Description & " | SQL: " & sqlCreate
        rsAccess.Close
        connAccess.Close
        connTarget.Close
        Response.End
    End If

    ' Insert data
    Dim sqlInsert, placeholders, values
    sqlInsert = "INSERT INTO [" & tableName & "] ("
    placeholders = ""
    
    For i = 0 To rsAccess.Fields.Count - 1
        sqlInsert = sqlInsert & "[" & rsAccess.Fields.Item(i).Name & "]"
        placeholders = placeholders & "?"
        If i < rsAccess.Fields.Count - 1 Then
            sqlInsert = sqlInsert & ", "
            placeholders = placeholders & ", "
        End If
    Next
    sqlInsert = sqlInsert & ") VALUES (" & placeholders & ")"

    ' Loop through records and insert
    Dim count
    count = 0
    
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
            Response.Write "ERROR: INSERT failed at record " & (count + 1) & ": " & Err.Description
            rsAccess.Close
            connAccess.Close
            connTarget.Close
            Response.End
        End If
        
        count = count + 1
        rsAccess.MoveNext
    Loop
    
    connTarget.CommitTrans

    rsAccess.Close
    connAccess.Close
    connTarget.Close

    Response.Write "SUCCESS: Imported " & count & " records from " & tableName & " to " & dbType
%>
