<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
' Test Dim Hoisting and Database Connection
' This test verifies the fix for QuickerSite's Dim hoisting pattern

' --- Test 1: Dim hoisting (assign before Dim) ---
testVar1 = "HoistOK"
Dim testVar1
Response.Write "Test1-DimHoist: " & testVar1 & vbCrLf

' --- Test 2: Multiple Dim hoisting ---
varA = 10
varB = 20
varC = 30
Dim varA, varB, varC
Response.Write "Test2-MultiDim: " & varA & "," & varB & "," & varC & vbCrLf

' --- Test 3: Server.MapPath ---
Dim dbPath
dbPath = Server.MapPath("/db/data_jj2ar6as.mdb")
Response.Write "Test3-MapPath: " & dbPath & vbCrLf

' --- Test 4: Database Connection ---
On Error Resume Next
Dim db2
Set db2 = Server.CreateObject("ADODB.Connection")
db2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
If Err.Number <> 0 Then
    Response.Write "Test4-DBOpen: FAIL - " & Err.Description & vbCrLf
    Err.Clear
Else
    Response.Write "Test4-DBOpen: OK" & vbCrLf
    ' --- Test 5: Query ---
    Dim rs2
    Set rs2 = db2.Execute("SELECT count(*) as cnt FROM tblCustomer")
    If Err.Number <> 0 Then
        Response.Write "Test5-Query: FAIL - " & Err.Description & vbCrLf
    Else
        If Not rs2.EOF Then
            Response.Write "Test5-Query: OK count=" & rs2("cnt") & vbCrLf
        End If
    End If
    db2.Close
End If
On Error Goto 0

' --- Test 6: Class with Dim hoisting pattern ---
execBeforePageLoad = "/common/before.asp"
QS_DBS = 1
Dim execBeforePageLoad, QS_DBS
Response.Write "Test6-ConfigVars: execBPL=" & execBeforePageLoad & " QS_DBS=" & QS_DBS & vbCrLf

' --- Test 7: Class Property access ---
Class cls_test
    Public myProp
    Private Sub Class_Initialize
        myProp = "ClassOK"
    End Sub
End Class
Dim obj
Set obj = New cls_test
Response.Write "Test7-ClassProp: " & obj.myProp & vbCrLf

Response.Write "ALL TESTS COMPLETE" & vbCrLf
%>
