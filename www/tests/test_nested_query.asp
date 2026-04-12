<%@ Language="VBScript" %>
<%
On Error Resume Next
Response.ContentType = "text/html"
Response.Write "<html><body>"
Response.Write "<h1>Constants Loading Simulation</h1>"

' Open database
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/db/data_jj2ar6as.mdb")
Response.Write "<div>Connection state: " & conn.State & "</div>"

If conn.State <> 1 Then
    Response.Write "<div>Cannot connect to DB</div>"
    Response.End
End If

' 1. Test basic recordset iteration with nested queries
Response.Write "<h2>Test 1: Nested query iteration</h2>"
Dim rs1
Set rs1 = conn.Execute("SELECT iId FROM tblConstant WHERE iCustomerID=73 ORDER BY sConstant")
If Err.Number <> 0 Then
    Response.Write "<div>Error: " & Err.Description & "</div>"
    Err.Clear
End If

Dim count1
count1 = 0
Do While Not rs1.EOF
    Dim innerRS
    Set innerRS = conn.Execute("SELECT sConstant FROM tblConstant WHERE iId=" & rs1("iId"))
    If Not innerRS Is Nothing Then
        If Not innerRS.EOF Then
            Response.Write "<div>" & count1 & ": iId=" & rs1("iId") & " sConstant=" & innerRS("sConstant") & "</div>"
        End If
        innerRS.Close
        Set innerRS = Nothing
    End If
    count1 = count1 + 1
    rs1.MoveNext
Loop
rs1.Close
Set rs1 = Nothing
Response.Write "<div><b>Total iterated: " & count1 & "</b></div>"

' 2. Test Dictionary with class-like objects
Response.Write "<h2>Test 2: Dictionary with class instances</h2>"
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")

Dim rs2
Set rs2 = conn.Execute("SELECT iId, sConstant FROM tblConstant WHERE iCustomerID=73 ORDER BY sConstant")
Dim count2
count2 = 0
Do While Not rs2.EOF
    Dim cid2, cname2
    cid2 = rs2("iId")
    cname2 = rs2("sConstant")
    dict.Add cid2, cname2
    If Err.Number <> 0 Then
        Response.Write "<div>Error adding key " & cid2 & ": " & Err.Description & "</div>"
        Err.Clear
    End If
    count2 = count2 + 1
    rs2.MoveNext
Loop
rs2.Close
Set rs2 = Nothing
Response.Write "<div><b>Dict count: " & dict.Count & " (iterated: " & count2 & ")</b></div>"

' List dictionary contents
Dim dkey
For Each dkey In dict
    Response.Write "<div>Key=" & dkey & " Value=" & dict(dkey) & "</div>"
Next

' 3. Test nested queries breaking outer recordset
Response.Write "<h2>Test 3: Does inner Execute affect outer rs?</h2>"
Dim rs3
Set rs3 = conn.Execute("SELECT iId FROM tblConstant WHERE iCustomerID=73 ORDER BY sConstant")
Dim count3
count3 = 0
Dim firstId, firstEof
firstId = rs3("iId")
Response.Write "<div>First record iId: " & firstId & " EOF: " & rs3.EOF & "</div>"

' Execute another query
Dim dummy
Set dummy = conn.Execute("SELECT * FROM tblConstant WHERE iId=" & firstId)
If Not dummy Is Nothing Then
    Response.Write "<div>Inner query returned sConstant: " & dummy("sConstant") & "</div>"
    dummy.Close
    Set dummy = Nothing
End If

' Check if outer rs is still valid
Response.Write "<div>After inner query - outer rs3.EOF: " & rs3.EOF & "</div>"
If Not rs3.EOF Then
    Response.Write "<div>Current rs3 iId: " & rs3("iId") & "</div>"
    rs3.MoveNext
    If Not rs3.EOF Then
        Response.Write "<div>After MoveNext rs3 iId: " & rs3("iId") & "</div>"
    Else
        Response.Write "<div>After MoveNext rs3 is EOF!</div>"
    End If
End If

' Continue counting
Do While Not rs3.EOF
    count3 = count3 + 1
    rs3.MoveNext
Loop
rs3.Close
Response.Write "<div><b>Remaining records after inner query + 1 movenext: " & count3 & "</b></div>"

conn.Close
Set conn = Nothing

Response.Write "</body></html>"
%>
