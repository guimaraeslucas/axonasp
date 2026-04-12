<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"

' Test 1: Server.MapPath with virtual path
Dim dbPath
dbPath = Server.MapPath("/db/data_jj2ar6as.mdb")
Response.Write "Test 1 - MapPath(/db/data_jj2ar6as.mdb): " & dbPath & vbCrLf

' Test 2: Check FSO FileExists
Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
Response.Write "Test 2 - FSO FileExists: " & fso.FileExists(dbPath) & vbCrLf

' Test 3: Try to open the database
On Error Resume Next
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
Dim connStr
connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath
Response.Write "Test 3 - Connection string: " & connStr & vbCrLf

conn.Open connStr
If Err.Number <> 0 Then
    Response.Write "Test 3 - FAILED to open DB: " & Err.Description & vbCrLf
    Err.Clear
Else
    Response.Write "Test 3 - Database opened OK. State=" & conn.State & vbCrLf
    
    ' Test 4: Simple query
    Dim rs
    Set rs = conn.Execute("SELECT COUNT(*) AS cnt FROM tblPages")
    If Err.Number <> 0 Then
        Response.Write "Test 4 - FAILED query: " & Err.Description & vbCrLf
        Err.Clear
    Else
        If Not rs.EOF Then
            Response.Write "Test 4 - tblPages count: " & rs("cnt") & vbCrLf
        Else
            Response.Write "Test 4 - No results" & vbCrLf
        End If
        rs.Close
        Set rs = Nothing
    End If
    
    conn.Close
End If
Set conn = Nothing
On Error GoTo 0

Response.Write "Done." & vbCrLf
%>
