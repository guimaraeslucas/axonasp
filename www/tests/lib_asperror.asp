<%
' Support library for enhanced ASP functionality
' This provides asperror and other helper functions

' Simple asperror function for error reporting
Public Function asperror(optional path)
    On Error Resume Next
    If Err.Number <> 0 Then
        Response.Write("<!-- Error: " & Err.Description)
        If Len(path) > 0 Then
            Response.Write(" (Path: " & path & ")")
        End If
        Response.Write(" -->" & vbCrLf)
        Err.Clear()
    End If
    On Error GoTo 0
End Function

' Alternative version with detailed error info
Public Function aspError_Detailed(optional path)
    On Error Resume Next
    If Err.Number <> 0 Then
        Response.Write("<!-- ")
        Response.Write("Error Number: " & Err.Number & ", ")
        Response.Write("Description: " & Err.Description & ", ")
        Response.Write("Source: " & Err.Source)
        If Len(path) > 0 Then
            Response.Write(", File: " & path)
        End If
        Response.Write(" -->" & vbCrLf)
        Err.Clear()
    End If
    On Error GoTo 0
End Function
%>
