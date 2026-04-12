<%@ Language="VBScript" %>
<%
Response.ContentType = "text/plain"
On Error Resume Next

' Test 1: Check if Application state has constants cached
Dim arrConst
arrConst = Application("QS_CMS_constants_1")
If IsEmpty(arrConst) Or IsNull(arrConst) Then
    Response.Write "Test 1 - Application constants: NOT CACHED" & vbCrLf
Else
    If IsArray(arrConst) Then
        Response.Write "Test 1 - Application constants: CACHED, IsArray=True, UBound=" & UBound(arrConst, 2) & vbCrLf
    Else
        Response.Write "Test 1 - Application constants: CACHED but not array, type=" & TypeName(arrConst) & vbCrLf
    End If
End If

' Test 2: Check customer object
Response.Write "Test 2 - customer type: " & TypeName(customer) & vbCrLf
If IsObject(customer) Then
    Response.Write "Test 2a - customer.siteName: '" & customer.siteName & "'" & vbCrLf
    Response.Write "Test 2b - customer.sSiteSlogan: '" & customer.sSiteSlogan & "'" & vbCrLf
    Response.Write "Test 2c - customer.cId: '" & customer.cId & "'" & vbCrLf
End If
If Err.Number <> 0 Then
    Response.Write "Test 2 - Error: " & Err.Description & vbCrLf
    Err.Clear
End If

' Test 3: Test DB access directly
Dim conn2, rs2
Set conn2 = Server.CreateObject("ADODB.Connection")
Dim dbPath2
dbPath2 = Server.MapPath("/db/data_jj2ar6as.mdb")
Response.Write "Test 3 - DB Path: " & dbPath2 & vbCrLf

conn2.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath2
If Err.Number <> 0 Then
    Response.Write "Test 3 - DB Open Error: " & Err.Description & vbCrLf
    Err.Clear
Else
    Response.Write "Test 3 - DB Open OK, State=" & conn2.State & vbCrLf
    
    ' Test 4: Query tblConstant 
    Set rs2 = conn2.Execute("SELECT sConstant, sValue FROM tblConstant WHERE sConstant='SITEINFO'")
    If Err.Number <> 0 Then
        Response.Write "Test 4 - Query Error: " & Err.Description & vbCrLf
        Err.Clear
    Else
        If rs2.EOF Then
            Response.Write "Test 4 - No SITEINFO constant found in tblConstant" & vbCrLf
        Else
            Response.Write "Test 4 - SITEINFO sConstant: '" & rs2("sConstant") & "'" & vbCrLf
            Response.Write "Test 4 - SITEINFO sValue length: " & Len(rs2("sValue") & "") & vbCrLf
            Response.Write "Test 4 - SITEINFO sValue first 200: '" & Left(rs2("sValue") & "", 200) & "'" & vbCrLf
        End If
        If Err.Number <> 0 Then
            Response.Write "Test 4 - Field Error: " & Err.Description & vbCrLf
            Err.Clear
        End If
    End If
    
    ' Test 5: Query tblPage 
    Set rs2 = conn2.Execute("SELECT iId, sTitle, sValue FROM tblPage WHERE iId=1")
    If Err.Number <> 0 Then
        Response.Write "Test 5 - Query Error: " & Err.Description & vbCrLf
        Err.Clear
    Else
        If rs2.EOF Then
            Response.Write "Test 5 - No page with iId=1" & vbCrLf
        Else
            Response.Write "Test 5 - Page iId: '" & rs2("iId") & "'" & vbCrLf
            Response.Write "Test 5 - Page sTitle: '" & rs2("sTitle") & "'" & vbCrLf
            Response.Write "Test 5 - Page sValue len: " & Len(rs2("sValue") & "") & vbCrLf
        End If
        If Err.Number <> 0 Then
            Response.Write "Test 5 - Field Error: " & Err.Description & vbCrLf
            Err.Clear
        End If
    End If
    
    conn2.Close
End If

Response.Write vbCrLf & "Done." & vbCrLf
%>
