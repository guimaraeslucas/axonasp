<%@ Language="VBScript" %>
<%
Response.ContentType = "text/html"
Response.Write "<html><body>"
Response.Write "<h1>Database Constants Count Diagnostic</h1>"

' Open database directly
Dim conn, rs
Set conn = Server.CreateObject("ADODB.Connection")

' Find the connection string used by QuickerSite
' Check if there's a db.mdb or similar
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Try to find the database
Dim dbPath
dbPath = Server.MapPath("/data/db.mdb")
Response.Write "<div>Looking for DB at: " & dbPath & "</div>"
Response.Write "<div>Exists: " & fso.FileExists(dbPath) & "</div>"

If Not fso.FileExists(dbPath) Then
    dbPath = Server.MapPath("/data/quickersite.mdb")
    Response.Write "<div>Looking for DB at: " & dbPath & "</div>"
    Response.Write "<div>Exists: " & fso.FileExists(dbPath) & "</div>"
End If

If Not fso.FileExists(dbPath) Then
    ' Look for any .mdb or .db file
    Response.Write "<div>Searching for database files in /data/...</div>"
    Dim dataFolder, f
    If fso.FolderExists(Server.MapPath("/data")) Then
        Set dataFolder = fso.GetFolder(Server.MapPath("/data"))
        For Each f In dataFolder.Files
            Response.Write "<div>Found file: " & f.Name & " (" & f.Size & " bytes)</div>"
        Next
    Else
        Response.Write "<div>/data folder does not exist</div>"
    End If
    
    ' Also check root
    Response.Write "<div>Searching root web folder for .mdb files...</div>"
    Dim rootFolder
    Set rootFolder = fso.GetFolder(Server.MapPath("/"))
    For Each f In rootFolder.Files
        If LCase(Right(f.Name, 4)) = ".mdb" Or LCase(Right(f.Name, 3)) = ".db" Then
            Response.Write "<div>Found: " & f.Name & " (" & f.Size & " bytes)</div>"
        End If
    Next
End If

' Try to use the same connection as QuickerSite
' Check Application for connection string
Response.Write "<h2>Application connection info</h2>"
Dim k
For Each k In Application.Contents
    If InStr(1, k, "conn", 1) > 0 Or InStr(1, k, "database", 1) > 0 Or InStr(1, k, "db", 1) > 0 Then
        Response.Write "<div>" & k & " = "
        If IsObject(Application.Contents(k)) Then
            Response.Write "OBJECT (" & TypeName(Application.Contents(k)) & ")"
        Else
            Response.Write Application.Contents(k)
        End If
        Response.Write "</div>"
    End If
Next

' Try a direct query  
On Error Resume Next
If fso.FileExists(dbPath) Then
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
    If Err.Number <> 0 Then
        ' Try SQLite
        conn.Open "Driver={SQLite3};Data Source=" & dbPath
    End If
    
    If Err.Number = 0 Then
        Set rs = conn.Execute("SELECT COUNT(*) FROM tblConstant")
        Response.Write "<h2>tblConstant count: " & rs(0) & "</h2>"
        rs.Close
        
        Set rs = conn.Execute("SELECT iId, sConstant, iCustomerID FROM tblConstant ORDER BY sConstant")
        Response.Write "<table border='1'><tr><th>iId</th><th>sConstant</th><th>iCustomerID</th></tr>"
        Do While Not rs.EOF
            Response.Write "<tr><td>" & rs("iId") & "</td><td>" & rs("sConstant") & "</td><td>" & rs("iCustomerID") & "</td></tr>"
            rs.MoveNext
        Loop
        rs.Close
        Response.Write "</table>"
        conn.Close
    Else
        Response.Write "<div>Cannot open database: " & Err.Description & "</div>"
    End If
Else
    Response.Write "<div>No database file found to query directly</div>"
End If
On Error GoTo 0

Set fso = Nothing
Response.Write "</body></html>"
%>
