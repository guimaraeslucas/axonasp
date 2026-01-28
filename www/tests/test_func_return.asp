<%
Option Explicit
Response.ContentType = "text/plain"

Class ReturnTest
    Public Function GetNumber()
        GetNumber = 42
    End Function
    
    Public Function GetString()
        Dim temp
        temp = "Hello"
        GetString = temp
    End Function
    
    Public Function CallHelper()
        CallHelper = HelperFunc()
    End Function
    
    Private Function HelperFunc()
        HelperFunc = "From Helper"
    End Function
End Class

Dim rt
Set rt = New ReturnTest

Response.Write "GetNumber: " & rt.GetNumber() & vbCrLf
Response.Write "GetString: " & rt.GetString() & vbCrLf
Response.Write "CallHelper: " & rt.CallHelper() & vbCrLf
%>
