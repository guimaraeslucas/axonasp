<%@ Language="VBScript" %>
<%
' Diagnostic: check what Application(QS_CMS_arrconstants) contains
Response.Write "<html><body>"
Response.Write "<h1>QuickerSite Constants Cache Diagnostic</h1>"

' Check common QuickerSite Application keys
Dim keys(20)
keys(0) = "QS_CMS_arrconstants"
keys(1) = "qs_cms_arrconstants"
keys(2) = "arrconstants"

Response.Write "<h2>Application Keys Check</h2>"

Dim i, val
For i = 0 To 2
    val = Application(keys(i))
    If IsEmpty(val) Or IsNull(val) Then
        Response.Write "<div>Application(""" & keys(i) & """) = EMPTY/NULL</div>"
    ElseIf IsArray(val) Then
        Response.Write "<div style='color:green'>Application(""" & keys(i) & """) = ARRAY</div>"
        Response.Write "<div>  LBound(1)=" & LBound(val, 1) & " UBound(1)=" & UBound(val, 1) & "</div>"
        Response.Write "<div>  LBound(2)=" & LBound(val, 2) & " UBound(2)=" & UBound(val, 2) & "</div>"
        Dim j
        For j = LBound(val, 2) To UBound(val, 2)
            Response.Write "<div>  [" & j & "] name=" & val(0, j) & " | value_len=" & Len(val(1, j)) & "</div>"
        Next
    Else
        Response.Write "<div>Application(""" & keys(i) & """) = " & TypeName(val) & ": " & val & "</div>"
    End If
Next

' Check all Application keys by examining common QS prefixes
Response.Write "<h2>All QS Application Variables</h2>"
Dim appContents
For Each appContents In Application.Contents
    If Left(LCase(appContents), 3) = "qs_" Or Left(LCase(appContents), 4) = "arr_" Or InStr(LCase(appContents), "constant") > 0 Or InStr(LCase(appContents), "arr") > 0 Then
        Response.Write "<div>" & appContents & " = "
        Dim appVal
        appVal = Application(appContents)
        If IsEmpty(appVal) Or IsNull(appVal) Then
            Response.Write "EMPTY"
        ElseIf IsArray(appVal) Then
            Response.Write "ARRAY"
        ElseIf IsObject(appVal) Then
            Response.Write "OBJECT: " & TypeName(appVal)
        Else
            Dim strVal
            strVal = CStr(appVal)
            If Len(strVal) > 100 Then strVal = Left(strVal, 100) & "..."
            Response.Write Server.HTMLEncode(strVal)
        End If
        Response.Write "</div>"
    End If
Next

Response.Write "</body></html>"
%>
