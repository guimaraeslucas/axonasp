<%@ Language="VBScript" %>
<%
' Comprehensive QuickerSite diagnostic test
' Tests the exact patterns used by QuickerSite

Response.ContentType = "text/plain"

' ============================================
' TEST 1: Application variable assignment
' ============================================
Response.Write "=== TEST 1: Application Variables ===" & vbCrLf
Application("TEST_VAR_1") = "hello"
Application("TEST_VAR_2") = 73
Response.Write "Application(TEST_VAR_1) = " & Application("TEST_VAR_1") & vbCrLf
Response.Write "Application(TEST_VAR_2) = " & Application("TEST_VAR_2") & vbCrLf

' ============================================
' TEST 2: Database connection with Access
' ============================================
Response.Write vbCrLf & "=== TEST 2: Database Connection ===" & vbCrLf
dim dbPath
dbPath = Server.MapPath("/QuickerSite-test/db/data_jj2ar6as.mdb")
Response.Write "DB Path: " & dbPath & vbCrLf

dim conn
set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
if Err.Number <> 0 then
    Response.Write "Connection Error: " & Err.Description & vbCrLf
    Err.Clear
else
    Response.Write "Connection State: " & conn.State & vbCrLf
end if
On Error Goto 0

' ============================================
' TEST 3: Recordset field access rs("field")
' ============================================
Response.Write vbCrLf & "=== TEST 3: Recordset Field Access ===" & vbCrLf
if conn.State = 1 then
    dim rs
    On Error Resume Next
    set rs = conn.Execute("SELECT * FROM tblCustomer WHERE iId=73")
    if Err.Number <> 0 then
        Response.Write "Execute Error: " & Err.Description & vbCrLf
        Err.Clear
    end if
    On Error Goto 0
    
    if not rs is nothing then
        Response.Write "RS is valid object: YES" & vbCrLf
        Response.Write "RS.EOF = " & rs.EOF & vbCrLf
        if not rs.EOF then
            Response.Write "rs(""sName"") = [" & rs("sName") & "]" & vbCrLf
            Response.Write "rs(""sURL"") = [" & rs("sURL") & "]" & vbCrLf
            Response.Write "rs(""iId"") = [" & rs("iId") & "]" & vbCrLf
            
            ' Also test rs.Fields("field").Value
            Response.Write "rs.Fields(""sName"").Value = [" & rs.Fields("sName").Value & "]" & vbCrLf
            
            ' Test numeric index
            Response.Write "rs(0) = [" & rs(0) & "]" & vbCrLf
            Response.Write "Fields.Count = " & rs.Fields.Count & vbCrLf
        else
            Response.Write "RS is EOF - no records found!" & vbCrLf
        end if
    else
        Response.Write "RS is Nothing!" & vbCrLf
    end if
else
    Response.Write "SKIPPED - No database connection" & vbCrLf
end if

' ============================================
' TEST 4: Class with Pick method (simulating customer)
' ============================================
Response.Write vbCrLf & "=== TEST 4: Class Method with RS Access ===" & vbCrLf

Class cls_test_customer
    Private p_name
    Private p_url
    Private p_id
    
    Public Property Get siteName
        siteName = p_name
    End Property
    
    Public Property Get siteUrl
        siteUrl = p_url
    End Property
    
    Public Property Get custId
        custId = p_id
    End Property
    
    Public Sub Pick(id)
        dim sql
        sql = "SELECT * FROM tblCustomer WHERE iId=" & id
        dim rsLocal
        On Error Resume Next
        set rsLocal = conn.Execute(sql)
        if Err.Number <> 0 then
            Response.Write "  Pick Execute Error: " & Err.Description & vbCrLf
            Err.Clear
            On Error Goto 0
            Exit Sub
        end if
        On Error Goto 0
        
        if not rsLocal is nothing then
            if not rsLocal.EOF then
                p_name = rsLocal("sName")
                p_url = rsLocal("sURL")
                p_id = rsLocal("iId")
                Response.Write "  Inside Pick - p_name=[" & p_name & "]" & vbCrLf
                Response.Write "  Inside Pick - p_url=[" & p_url & "]" & vbCrLf
                Response.Write "  Inside Pick - p_id=[" & p_id & "]" & vbCrLf
            else
                Response.Write "  Pick: RS is EOF" & vbCrLf
            end if
        else
            Response.Write "  Pick: RS is Nothing" & vbCrLf
        end if
        set rsLocal = nothing
    End Sub
End Class

if conn.State = 1 then
    dim testCust
    set testCust = new cls_test_customer
    testCust.Pick 73
    Response.Write "testCust.siteName = [" & testCust.siteName & "]" & vbCrLf
    Response.Write "testCust.siteUrl = [" & testCust.siteUrl & "]" & vbCrLf
    Response.Write "testCust.custId = [" & testCust.custId & "]" & vbCrLf
else
    Response.Write "SKIPPED - No database connection" & vbCrLf
end if

' ============================================
' TEST 5: 2D Array storage in Application
' ============================================
Response.Write vbCrLf & "=== TEST 5: 2D Array in Application ===" & vbCrLf
dim arr2d
ReDim arr2d(2, 3)
arr2d(0, 0) = "MENU"
arr2d(1, 0) = "<nav>Menu Content</nav>"
arr2d(2, 0) = ""
arr2d(0, 1) = "SITEINFO"
arr2d(1, 1) = "My Site Info"
arr2d(2, 1) = ""
arr2d(0, 2) = "LOGO"
arr2d(1, 2) = "<img src='logo.png'/>"
arr2d(2, 2) = ""

Application("TEST_CONSTANTS") = arr2d
dim retrieved
retrieved = Application("TEST_CONSTANTS")
Response.Write "Type of retrieved: " & TypeName(retrieved) & vbCrLf

dim canIterate
canIterate = false
On Error Resume Next
dim ub2
ub2 = UBound(retrieved, 2)
if Err.Number = 0 then
    canIterate = true
    Response.Write "UBound(retrieved, 2) = " & ub2 & vbCrLf
else
    Response.Write "Error getting UBound(,2): " & Err.Description & vbCrLf
    Err.Clear
end if
On Error Goto 0

if canIterate then
    dim i
    for i = 0 to ub2
        Response.Write "  Constant(" & i & "): name=[" & retrieved(0, i) & "] value=[" & retrieved(1, i) & "]" & vbCrLf
    next
end if

' ============================================
' TEST 6: Blank lines test  
' ============================================
Response.Write vbCrLf & "=== TEST 6: Include Path Resolution ===" & vbCrLf
Response.Write "Server.MapPath("""") = " & Server.MapPath("") & vbCrLf
Response.Write "Server.MapPath(""/"") = " & Server.MapPath("/") & vbCrLf
Response.Write "Server.MapPath(""/db/data_jj2ar6as.mdb"") = " & Server.MapPath("/db/data_jj2ar6as.mdb") & vbCrLf
Response.Write "Server.MapPath(""/QuickerSite-test/db/data_jj2ar6as.mdb"") = " & Server.MapPath("/QuickerSite-test/db/data_jj2ar6as.mdb") & vbCrLf

' ============================================
' TEST 7: convertGetal function   
' ============================================
Response.Write vbCrLf & "=== TEST 7: Type Conversion ===" & vbCrLf
Response.Write "CInt(""73"") = " & CInt("73") & vbCrLf
Response.Write "CLng(""73"") = " & CLng("73") & vbCrLf

' ============================================
' TEST 8: Set FunctionName = value pattern
' ============================================
Response.Write vbCrLf & "=== TEST 8: Set FunctionName = value ===" & vbCrLf

Class cls_test_db
    Public Function Execute(sql)
        On Error Resume Next
        Set Execute = conn.Execute(sql)
        On Error Goto 0
    End Function
End Class

if conn.State = 1 then
    dim testDb
    set testDb = new cls_test_db
    dim testRs
    set testRs = testDb.Execute("SELECT * FROM tblCustomer WHERE iId=73")
    if not testRs is nothing then
        Response.Write "testDb.Execute returned: " & TypeName(testRs) & vbCrLf
        Response.Write "testRs.EOF = " & testRs.EOF & vbCrLf
        if not testRs.EOF then
            Response.Write "testRs(""sName"") = [" & testRs("sName") & "]" & vbCrLf
        end if
    else
        Response.Write "testDb.Execute returned Nothing!" & vbCrLf
    end if
else
    Response.Write "SKIPPED - No database connection" & vbCrLf
end if

' ============================================
' CLEANUP
' ============================================
if conn.State = 1 then
    conn.Close
end if
set conn = nothing

Response.Write vbCrLf & "=== ALL TESTS COMPLETE ===" & vbCrLf
%>
