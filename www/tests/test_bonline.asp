<%@ Language="VBScript" %>
<%
On Error Resume Next
Response.ContentType = "text/html"
Response.Write "<html><body>"
Response.Write "<h1>Variable Scope Debug</h1>"

' Check what global variables are accessible
Response.Write "<h2>Global Variables</h2>"

Response.Write "<div>cId: [" 
Response.Write cId
Response.Write "] (TypeName: " & TypeName(cId) & ")</div>"
If Err.Number <> 0 Then Response.Write "<div style='color:red'>cId ERR: " & Err.Description & "</div>" : Err.Clear

Response.Write "<div>Application('QS_CMS_iCustomerID'): [" 
Response.Write Application("QS_CMS_iCustomerID")
Response.Write "]</div>"
If Err.Number <> 0 Then Response.Write "<div style='color:red'>App ERR: " & Err.Description & "</div>" : Err.Clear

' Check db object
Response.Write "<div>db is Nothing: "
If db Is Nothing Then
    Response.Write "True"
Else
    Response.Write "False"
End If
Response.Write "</div>"
If Err.Number <> 0 Then Response.Write "<div style='color:red'>db ERR: " & Err.Description & "</div>" : Err.Clear

' Check customer
Response.Write "<div>customer: "
Response.Write TypeName(customer) 
Response.Write "</div>"
If Err.Number <> 0 Then Response.Write "<div style='color:red'>customer ERR: " & Err.Description & "</div>" : Err.Clear

Response.Write "<div>customer.iId: "
Response.Write customer.iId
Response.Write "</div>"
If Err.Number <> 0 Then Response.Write "<div style='color:red'>customer.iId ERR: " & Err.Description & "</div>" : Err.Clear

' Check QS_CMS_arrconstants variable 
Response.Write "<h2>Cache Key</h2>"
Response.Write "<div>QS_CMS_arrconstants var: "
Response.Write QS_CMS_arrconstants
Response.Write "</div>"
If Err.Number <> 0 Then Response.Write "<div style='color:red'>QS_CMS_arr ERR: " & Err.Description & "</div>" : Err.Clear

' Now directly try what pick does with cid
Response.Write "<h2>Pick Query Test</h2>"
Dim tconn
Set tconn = Server.CreateObject("ADODB.Connection")
tconn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("/db/data_jj2ar6as.mdb")

' Test with hardcoded cid=73
Dim trs1
Set trs1 = tconn.Execute("select iId, sConstant from tblConstant where iCustomerID=73 and iId=1141")
If Not trs1 Is Nothing Then
    If Not trs1.EOF Then
        Response.Write "<div>cid=73, iId=1141: sConstant=" & trs1("sConstant") & "</div>"
    Else
        Response.Write "<div>cid=73, iId=1141: EOF (no results)</div>"
    End If
End If

' Test with actual cid value
Dim trs2, tcid
tcid = cId
If Err.Number <> 0 Then tcid = 0 : Err.Clear
Response.Write "<div>Testing with cid=" & tcid & "</div>"
Set trs2 = tconn.Execute("select iId, sConstant from tblConstant where iCustomerID=" & tcid & " and iId=1141")
If Err.Number <> 0 Then Response.Write "<div style='color:red'>Query ERR: " & Err.Description & "</div>" : Err.Clear
If Not trs2 Is Nothing Then
    If Not trs2.EOF Then
        Response.Write "<div>cid=" & tcid & ", iId=1141: sConstant=" & trs2("sConstant") & "</div>"
    Else
        Response.Write "<div style='color:red'>cid=" & tcid & ", iId=1141: EOF (NO RESULTS!)</div>"
    End If
Else
    Response.Write "<div style='color:red'>trs2 is Nothing</div>"
End If

tconn.Close
Set tconn = Nothing

Response.Write "</body></html>"
%>
