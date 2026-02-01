<%@ Language="VBScript" %>
<%
Option Explicit
Response.ContentType = "text/html"
Dim conn, rs, i, fld, val, vt

' Open database connection
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Server.MapPath("/db/sample.mdb")

' Execute query
Set rs = conn.Execute("SELECT TOP 3 id, text, number, boolean FROM sample")

Response.Write "<h1>JSON Debug Test</h1>"
Response.Write "<h2>TypeName(rs): " & TypeName(rs) & "</h2>"
Response.Write "<h2>rs.Fields.Count: " & rs.Fields.Count & "</h2>"

Response.Write "<table border='1'>"
Response.Write "<tr><th>#</th><th>Name</th><th>Value</th><th>VarType</th><th>TypeName</th><th>CLng</th></tr>"

Do While Not rs.EOF
    For i = 0 To rs.Fields.Count - 1
        Set fld = rs.Fields(i)
        val = fld.Value
        vt = VarType(val)
        
        Response.Write "<tr>"
        Response.Write "<td>" & i & "</td>"
        Response.Write "<td>" & fld.Name & "</td>"
        Response.Write "<td>" & val & "</td>"
        Response.Write "<td>" & vt & "</td>"
        Response.Write "<td>" & TypeName(val) & "</td>"
        
        On Error Resume Next
        If vt = 2 Or vt = 3 Then
            Response.Write "<td>" & CLng(val) & "</td>"
        Else
            Response.Write "<td>N/A</td>"
        End If
        If Err.Number <> 0 Then
            Response.Write "<td>ERROR: " & Err.Description & "</td>"
            Err.Clear
        End If
        On Error GoTo 0
        
        Response.Write "</tr>"
    Next
    rs.MoveNext
    Response.Write "<tr><td colspan='6'><hr></td></tr>"
Loop

Response.Write "</table>"

rs.Close
conn.Close
Set rs = Nothing
Set conn = Nothing
%>
