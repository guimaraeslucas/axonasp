<%@ Language=VBScript %>
<html>
<head>
    <title>Platform Parity Test</title>
    <meta charset="utf-8">
</head>
<body>
<%
Sub WriteCheck(name, passed, details)
    If passed Then
        Response.Write "PASS - " & name
    Else
        Response.Write "FAIL - " & name
    End If
    If details <> "" Then
        Response.Write " : " & Server.HTMLEncode(details)
    End If
    Response.Write "<br>"
End Sub

Dim rndText, dotPos, decLen
rndText = CStr(Rnd)
dotPos = InStr(1, rndText, ".")
decLen = 0
If dotPos > 0 Then
    decLen = Len(rndText) - dotPos
End If
Call WriteCheck("Rnd max 7 decimals", decLen <= 7, "value=" & rndText)

Dim dv
dv = CStr(DateValue("12/25/2025"))
Call WriteCheck("DateValue 4-digit year", dv = "12/25/2025", "value=" & dv)

Session.LCID = ""
Call WriteCheck("Session.LCID default", CStr(Session.LCID) = "1033", "value=" & CStr(Session.LCID))

Session.Timeout = ""
Call WriteCheck("Session.Timeout default", CStr(Session.Timeout) = "20", "value=" & CStr(Session.Timeout))

Dim fc, fn
fc = CStr(FormatCurrency(1234.5))
fn = CStr(FormatNumber(1234.5, 2))
Call WriteCheck("FormatCurrency Windows-like", fc = "$1234.50", "value=" & fc)
Call WriteCheck("FormatNumber Windows-like", fn = "1234.50", "value=" & fn)
%>
</body>
</html>
