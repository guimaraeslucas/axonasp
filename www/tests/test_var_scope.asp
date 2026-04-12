<%@ Language="VBScript" %>
<%
On Error Resume Next
Response.ContentType = "text/html"
Response.Write "<html><body>"
Response.Write "<h1>Variable Scoping Test</h1>"

' Open database
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/db/data_jj2ar6as.mdb")

Class InnerClass
    Public iId
    Public sConstant
    
    Public Function Pick(id)
        ' This uses DIM to declare local RS - should NOT affect outer RS
        Dim sql, RS
        sql = "SELECT * FROM tblConstant WHERE iCustomerID=73 AND iId=" & id
        Set RS = conn.Execute(sql)
        If Not RS.EOF Then
            iId = RS("iId")
            sConstant = RS("sConstant")
        End If
        RS.Close
        Set RS = Nothing
    End Function
End Class

' Test: Does inner function's DIM RS shadow outer RS?
Response.Write "<h2>Test 1: Variable scoping with DIM</h2>"
Dim rs, constant
Set rs = conn.Execute("SELECT iId FROM tblConstant WHERE iCustomerID=73 ORDER BY sConstant")

Dim count1
count1 = 0
Do While Not rs.EOF
    Set constant = New InnerClass
    constant.Pick(rs("iId"))
    Response.Write "<div>" & count1 & ": outer_rs_iId=" & rs("iId") & " inner_sConstant=" & constant.sConstant & "</div>"
    If Err.Number <> 0 Then
        Response.Write "<div>ERROR at " & count1 & ": " & Err.Number & " - " & Err.Description & "</div>"
        Err.Clear
    End If
    Set constant = Nothing
    count1 = count1 + 1
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Response.Write "<div><b>Total outer iterations: " & count1 & "</b></div>"

' Test 2: Without DIM in inner function (accessing outer scope variable)  
Response.Write "<h2>Test 2: Outer variables check</h2>"
Dim testVar
testVar = "OUTER"

Sub InnerSub()
    Dim testVar
    testVar = "INNER"
End Sub

InnerSub
Response.Write "<div>testVar after InnerSub: " & testVar & " (should be OUTER)</div>"

Sub InnerSub2()
    ' No DIM - should this modify outer scope?
    testVar = "MODIFIED"
End Sub

InnerSub2
Response.Write "<div>testVar after InnerSub2: " & testVar & " (should be MODIFIED)</div>"

conn.Close
Set conn = Nothing
Response.Write "</body></html>"
%>
