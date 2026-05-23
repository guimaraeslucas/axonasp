<%@ Language="VBScript" %>
<%
' Advanced logging function combining all features
Function LogMessage(ByVal level As String, _
                    ByVal msg As String, _
                    Optional timestamp As String = "now", _
                    ParamArray tags())
    Dim result, i
    result = "[" & level & "] " & msg
    
    If timestamp <> "now" Then
        result = result & " @" & timestamp
    End If
    
    If IsArray(tags) And UBound(tags) >= LBound(tags) Then
        result = result & " {"
        For i = LBound(tags) To UBound(tags)
            If i > LBound(tags) Then result = result & ", "
            result = result & tags(i)
        Next
        result = result & "}"
    End If
    
    LogMessage = result
End Function

' Examples
Response.Write LogMessage("INFO", "Application started") & "<br>"
Response.Write LogMessage("WARN", "High memory usage", "now", "server1", "memory") & "<br>"
Response.Write LogMessage("ERROR", "Connection failed", "12:00:00", "critical", "db", "timeout")
%>