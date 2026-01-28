<%
Option Explicit
debug_asp_code = "TRUE"
Response.ContentType = "text/html"

Response.Write "<h3>MD5 Step-by-Step Debug</h3>"

On Error Resume Next

' Step 1: Create a simple class with function
Class SimpleTest
    Public Function GetValue()
        GetValue = "Hello"
    End Function
    
    Private Function PrivateHelper()
        PrivateHelper = "World"
    End Function
    
    Public Function CallPrivate()
        CallPrivate = PrivateHelper()
    End Function
End Class

Dim st
Set st = New SimpleTest
Response.Write "SimpleTest.GetValue(): " & st.GetValue() & "<br>"
Response.Write "SimpleTest.CallPrivate(): " & st.CallPrivate() & "<br>"

' Step 2: Test WordToHex-like logic
Class HashTest
    Private Function ByteToHex(b)
        ByteToHex = Right("0" & Hex(b), 2)
    End Function
    
    Public Function TestWordToHex()
        Dim result
        result = ""
        result = result & ByteToHex(72)
        result = result & ByteToHex(101)
        TestWordToHex = result
    End Function
End Class

Dim ht
Set ht = New HashTest
Response.Write "HashTest.TestWordToHex(): " & ht.TestWordToHex() & " (expected: 4865)<br>"

' Step 3: Test actual MD5 plugin
Response.Write "<br><b>Testing MD5 Plugin:</b><br>"
Dim md5Obj
Set md5Obj = aspL.plugin("md5")

If Err.Number <> 0 Then
    Response.Write "ERROR loading: " & Err.Description & "<br>"
    Err.Clear
End If

Dim hash
hash = md5Obj.md5("test", 32)

If Err.Number <> 0 Then
    Response.Write "ERROR calling md5: " & Err.Description & "<br>"
    Err.Clear
Else
    Response.Write "MD5 Result: '" & hash & "' (Length: " & Len(hash) & ")<br>"
    Response.Write "Expected: '098f6bcd4621d373cade4e832627b4f6' (32 chars)<br>"
End If

On Error Goto 0
%>
