<!-- #include file="asp/includes/constants.asp" -->
<!-- #include file="asp/includes/database.asp" -->
<!-- #include file="asp/includes/encryption.asp" -->
<!-- #include file="asp/includes/functions.asp" -->
<!-- #include file="asp/includes/page.asp" -->
<%
ON Error Resume Next

Response.Write "Test Page.Pick Flow" & vbCrLf
Response.Write "====================" & vbCrLf & vbCrLf

Dim selectedPage
Set selectedPage = New cls_page

Dim testId
testId = Request("iId")

Response.Write "Step 1: Request(""iId""): [" & testId & "]" & vbCrLf

If IsEmpty(testId) Or Len(testId) = 0 Then
    Response.Write "Step 2: No iId parameter found" & vbCrLf
Else
    Response.Write "Step 2: Parameter found, calling decrypt..." & vbCrLf
    
    Dim decrypted
    decrypted = decrypt(Left(testId, 40))
    
    Response.Write "Step 3: Decrypted value: [" & decrypted & "]" & vbCrLf
    Response.Write "Step 4: Calling selectedPage.pick(" & decrypted & ")..." & vbCrLf
    
    selectedPage.pick(decrypted)
    
    If Err.Number <> 0 Then
        Response.Write "ERROR after pick: " & Err.Description & " (Error " & Err.Number & ")" & vbCrLf
        Err.Clear
    End If
    
    Response.Write "Step 5: After pick(), selectedPage.iId = " & selectedPage.iId & vbCrLf
    
    If IsNull(selectedPage.iId) Then
        Response.Write "Step 6: iId is NULL, would call pickByCode" & vbCrLf
    End If
End If

Response.Write vbCrLf & "Test complete!" & vbCrLf

ON Error Goto 0
%>
