<%@ Language="VBScript" %>
<%
On Error Resume Next
Response.ContentType = "text/html"
Response.Write "<html><body>"
Response.Write "<h1>Direct DB Constants Query</h1>"

Dim dbPath
dbPath = Server.MapPath("/db/data_jj2ar6as.mdb")
Response.Write "<div>DB Path: " & dbPath & "</div>"

Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
Response.Write "<div>Connection created: " & TypeName(conn) & "</div>"

conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
If Err.Number <> 0 Then
    Response.Write "<div>Connection error: " & Err.Number & " - " & Err.Description & "</div>"
    Err.Clear
End If
Response.Write "<div>Connection state: " & conn.State & "</div>"

If conn.State = 1 Then
    Dim rs
    Set rs = conn.Execute("SELECT COUNT(*) as cnt FROM tblConstant")
    If Err.Number <> 0 Then
        Response.Write "<div>Query error: " & Err.Number & " - " & Err.Description & "</div>"
        Err.Clear
    End If
    If Not rs Is Nothing Then
        Response.Write "<h2>Total constants in DB: " & rs(0) & "</h2>"
        rs.Close
    Else
        Response.Write "<div>Recordset is Nothing!</div>"
    End If
    
    Set rs = conn.Execute("SELECT iId, sConstant, iCustomerID FROM tblConstant ORDER BY sConstant")
    If Not rs Is Nothing Then
        Response.Write "<table border='1'><tr><th>iId</th><th>sConstant</th><th>iCustomerID</th></tr>"
        Do While Not rs.EOF
            Response.Write "<tr><td>" & rs("iId") & "</td><td>" & rs("sConstant") & "</td><td>" & rs("iCustomerID") & "</td></tr>"
            rs.MoveNext
        Loop
        rs.Close
        Response.Write "</table>"
    Else
        Response.Write "<div>Recordset is Nothing for SELECT!</div>"
    End If
    conn.Close
Else
    Response.Write "<div>Connection not open</div>"
End If

Set conn = Nothing
Response.Write "</body></html>"
%>
