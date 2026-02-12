<%
Response.Write "Test On Error Resume Next in Class Methods" & vbCrLf
Response.Write "============================================" & vbCrLf & vbCrLf

Class TestClass
    Public Sub TestMethod()
        Response.Write "Step 1: Inside TestMethod, before OERN" & vbCrLf
        
        On Error Resume Next
        
        Response.Write "Step 2: OERN active" & vbCrLf
        
        ' Try to access .eof on a nil recordset
        Dim rs
        Set rs = Nothing
        
        Response.Write "Step 3: About to access rs.eof (rs is Nothing)" & vbCrLf
        
        Dim testEof
        testEof = rs.eof
        
        If Err.Number <> 0 Then
            Response.Write "Step 4: Error caught: " & Err.Description & " (Error " & Err.Number & ")" & vbCrLf
            Err.Clear
        Else
            Response.Write "Step 4: NO ERROR - testEof = " & testEof & vbCrLf
        End If
        
        Response.Write "Step 5: After error, continuing..." & vbCrLf
        
        On Error Goto 0
        
        Response.Write "Step 6: Exiting TestMethod successfully" & vbCrLf
    End Sub
End Class

Dim testObj
Set testObj = New TestClass

Response.Write vbCrLf & "Calling class method..." & vbCrLf
testObj.TestMethod()

Response.Write vbCrLf & "Back in main script - test complete!" & vbCrLf
%>
