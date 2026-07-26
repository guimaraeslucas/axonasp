<%
' ============================================================
' lib/helpers.asp - Utility functions
' ============================================================

' JSEsc escapes a string for safe use inside JavaScript single-quoted strings.
Function JSEsc(ByVal s)
    s = Replace(s, "\", "\\")
    s = Replace(s, "'", "\'")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbCr, " ")
    JSEsc = s
End Function

' FmtSize formats a byte count into a human-readable string.
Function FmtSize(ByVal bytes)
    If bytes < 1024 Then
        FmtSize = bytes & " B"
    ElseIf bytes < 1048576 Then
        FmtSize = Round(bytes / 1024, 1) & " KB"
    ElseIf bytes < 1073741824 Then
        FmtSize = Round(bytes / 1048576, 1) & " MB"
    Else
        FmtSize = Round(bytes / 1073741824, 2) & " GB"
    End If
End Function

' IsAudioFile returns True if the file extension is a supported audio format.
Function IsAudioFile(ByVal fso, ByVal fileName)
    Dim ext
    ext = LCase(fso.GetExtensionName(fileName))
    IsAudioFile = (ext = "mp3" Or ext = "wav" Or ext = "ogg" _
        Or ext = "flac" Or ext = "m4a" Or ext = "wma" Or ext = "aac")
End Function

' HtmlEsc escapes a string for safe HTML output.
Function HtmlEsc(ByVal s)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HtmlEsc = s
End Function
%>
