<%
    ' Test VBScript VMENGINE global constant
    Response.Write "VBScript Engine ID: " & VMENGINE & "<br/>"
    
    ' Verify it's the correct string
    If VMENGINE = "G3pix AxonASP VBScript Engine" Then
        Response.Write "✓ VBScript VMENGINE constant works correctly<br/>"
    Else
        Response.Write "✗ VBScript VMENGINE returned unexpected value: " & VMENGINE & "<br/>"
    End If
%>