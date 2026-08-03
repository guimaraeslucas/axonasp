<%@ Language="VBScript" CodePage="65001" %>
<%
Response.CharSet = "UTF-8"
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>VBScript Constants Test / VBScript 常量测试</title>
<style>
body { font-family: Consolas, "Courier New", monospace; background: #1e1e1e; color: #d4d4d4; margin: 20px; }
h1 { color: #569cd6; border-bottom: 2px solid #569cd6; padding-bottom: 10px; }
h2 { color: #4ec9b0; margin-top: 30px; background: #2d2d2d; padding: 8px 12px; border-left: 4px solid #4ec9b0; }
table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
th { background: #094771; color: #fff; text-align: left; padding: 8px 12px; }
td { padding: 6px 12px; border-bottom: 1px solid #333; }
tr:hover { background: #2a2d2e; }
.name { color: #9cdcfe; font-weight: bold; }
.val { color: #b5cea8; }
.status-ok { color: #4ec9b0; }
.status-fail { color: #f44747; font-weight: bold; }
.summary { background: #094771; padding: 15px; margin-top: 30px; border-radius: 4px; font-size: 1.1em; }
.pass { color: #4ec9b0; font-weight: bold; }
.fail { color: #f44747; font-weight: bold; }
.note { color: #6a9955; font-style: italic; }
</style>
</head>
<body>

<h1>VBScript Built-in Constants Test / VBScript 内置常量测试</h1>
<p class="note">
This page tests all 79 VBScript built-in constants from the Script 5.6 CHM knowledge base.<br>
本页测试 Script 5.6 CHM 知识库中全部 79 个 VBScript 内置常量。
</p>

<%
Dim totalTests, passCount, failCount
totalTests = 0
passCount = 0
failCount = 0

Sub TestConst(constName, expected, actual)
    totalTests = totalTests + 1
    Dim status, cssClass
    If CStr(expected) = CStr(actual) Then
        status = "PASS / 通过"
        cssClass = "status-ok"
        passCount = passCount + 1
    Else
        status = "FAIL / 失败"
        cssClass = "status-fail"
        failCount = failCount + 1
    End If
    Response.Write "<tr>"
    Response.Write "<td class=""name"">" & constName & "</td>"
    Response.Write "<td class=""val"">" & expected & "</td>"
    Response.Write "<td class=""val"">" & actual & "</td>"
    Response.Write "<td class=""" & cssClass & """>" & status & "</td>"
    Response.Write "</tr>" & vbCrLf
End Sub

Sub TestConstHex(constName, expectedHex, actual)
    totalTests = totalTests + 1
    Dim expectedDec, status, cssClass
    expectedDec = CLng("&H" & Replace(expectedHex, "&h", ""))
    If expectedDec = CLng(actual) Then
        status = "PASS / 通过"
        cssClass = "status-ok"
        passCount = passCount + 1
    Else
        status = "FAIL / 失败"
        cssClass = "status-fail"
        failCount = failCount + 1
    End If
    Response.Write "<tr>"
    Response.Write "<td class=""name"">" & constName & "</td>"
    Response.Write "<td class=""val"">" & expectedHex & " (" & expectedDec & ")</td>"
    Response.Write "<td class=""val"">" & actual & "</td>"
    Response.Write "<td class=""" & cssClass & """>" & status & "</td>"
    Response.Write "</tr>" & vbCrLf
End Sub
%>

<!-- ============================================================ -->
<h2>1. String Constants / 字符串常量 (9) — vsconstring.htm</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConst "vbCr", Asc(Chr(13)), Asc(vbCr)
TestConst "vbCrLf", Chr(13) & Chr(10), vbCrLf
TestConst "vbFormFeed", Asc(Chr(12)), Asc(vbFormFeed)
TestConst "vbLf", Asc(Chr(10)), Asc(vbLf)
' vbNewLine: platform-dependent, just check it's not empty
totalTests = totalTests + 1
If Len(vbNewLine) > 0 Then
    passCount = passCount + 1
    Response.Write "<tr><td class=""name"">vbNewLine</td><td class=""val"">Len>0</td><td class=""val"">Len=" & Len(vbNewLine) & " (Asc=" & Asc(vbNewLine) & ")</td><td class=""status-ok"">PASS / 通过</td></tr>" & vbCrLf
Else
    failCount = failCount + 1
    Response.Write "<tr><td class=""name"">vbNewLine</td><td class=""val"">Len>0</td><td class=""val"">Len=0</td><td class=""status-fail"">FAIL / 失败</td></tr>" & vbCrLf
End If
TestConst "vbNullChar", Asc(Chr(0)), Asc(vbNullChar)
TestConst "vbNullString", "", vbNullString
TestConst "vbTab", Asc(Chr(9)), Asc(vbTab)
TestConst "vbVerticalTab", Asc(Chr(11)), Asc(vbVerticalTab)
%>
</table>

<!-- ============================================================ -->
<h2>2. Tristate Constants / 三态常量 (3) — vscontristate.htm</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConst "vbUseDefault", -2, vbUseDefault
TestConst "vbTrue", -1, vbTrue
TestConst "vbFalse", 0, vbFalse
%>
</table>

<!-- ============================================================ -->
<h2>3. Color Constants / 颜色常量 (8) — vsconcolor.htm</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConstHex "vbBlack", "&h00", vbBlack
TestConstHex "vbRed", "&hFF", vbRed
TestConstHex "vbGreen", "&hFF00", vbGreen
TestConstHex "vbYellow", "&hFFFF", vbYellow
TestConstHex "vbBlue", "&hFF0000", vbBlue
TestConstHex "vbMagenta", "&hFF00FF", vbMagenta
TestConstHex "vbCyan", "&hFFFF00", vbCyan
TestConstHex "vbWhite", "&hFFFFFF", vbWhite
%>
</table>

<!-- ============================================================ -->
<h2>4. Comparison Constants / 比较常量 (2) — vsconcompare.htm</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConst "vbBinaryCompare", 0, vbBinaryCompare
TestConst "vbTextCompare", 1, vbTextCompare
%>
</table>

<!-- ============================================================ -->
<h2>5. Date and Time Constants / 日期和时间常量 (11) — vscondatetime.htm</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConst "vbUseSystemDayOfWeek", 0, vbUseSystemDayOfWeek
TestConst "vbSunday", 1, vbSunday
TestConst "vbMonday", 2, vbMonday
TestConst "vbTuesday", 3, vbTuesday
TestConst "vbWednesday", 4, vbWednesday
TestConst "vbThursday", 5, vbThursday
TestConst "vbFriday", 6, vbFriday
TestConst "vbSaturday", 7, vbSaturday
TestConst "vbFirstJan1", 1, vbFirstJan1
TestConst "vbFirstFourDays", 2, vbFirstFourDays
TestConst "vbFirstFullWeek", 3, vbFirstFullWeek
%>
</table>

<!-- ============================================================ -->
<h2>6. Date Format Constants / 日期格式常量 (5) — vscondateformat.htm</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConst "vbGeneralDate", 0, vbGeneralDate
TestConst "vbLongDate", 1, vbLongDate
TestConst "vbShortDate", 2, vbShortDate
TestConst "vbLongTime", 3, vbLongTime
TestConst "vbShortTime", 4, vbShortTime
%>
</table>

<!-- ============================================================ -->
<h2>7. Miscellaneous Constants / 杂项常量 (1) — vsconmisc.htm</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConst "vbObjectError", -2147221504, vbObjectError
%>
</table>

<!-- ============================================================ -->
<h2>8. MsgBox Parameter Constants / MsgBox 参数常量 (16) — vsconmsgbox.htm Table 1</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
' Button types
TestConst "vbOKOnly", 0, vbOKOnly
TestConst "vbOKCancel", 1, vbOKCancel
TestConst "vbAbortRetryIgnore", 2, vbAbortRetryIgnore
TestConst "vbYesNoCancel", 3, vbYesNoCancel
TestConst "vbYesNo", 4, vbYesNo
TestConst "vbRetryCancel", 5, vbRetryCancel
' Icon types
TestConst "vbCritical", 16, vbCritical
TestConst "vbQuestion", 32, vbQuestion
TestConst "vbExclamation", 48, vbExclamation
TestConst "vbInformation", 64, vbInformation
' Default buttons
TestConst "vbDefaultButton1", 0, vbDefaultButton1
TestConst "vbDefaultButton2", 256, vbDefaultButton2
TestConst "vbDefaultButton3", 512, vbDefaultButton3
TestConst "vbDefaultButton4", 768, vbDefaultButton4
' Modal
TestConst "vbApplicationModal", 0, vbApplicationModal
TestConst "vbSystemModal", 4096, vbSystemModal
%>
</table>

<!-- ============================================================ -->
<h2>9. MsgBox Return Value Constants / MsgBox 返回值常量 (7) — vsconmsgbox.htm Table 2</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConst "vbOK", 1, vbOK
TestConst "vbCancel", 2, vbCancel
TestConst "vbAbort", 3, vbAbort
TestConst "vbRetry", 4, vbRetry
TestConst "vbIgnore", 5, vbIgnore
TestConst "vbYes", 6, vbYes
TestConst "vbNo", 7, vbNo
%>
</table>

<!-- ============================================================ -->
<h2>10. VarType Constants / VarType 常量 (17) — vsconvartype.htm</h2>
<table>
<tr><th>Constant / 常量</th><th>Expected / 期望值</th><th>Actual / 实际值</th><th>Result / 结果</th></tr>
<%
TestConst "vbEmpty", 0, vbEmpty
TestConst "vbNull", 1, vbNull
TestConst "vbInteger", 2, vbInteger
TestConst "vbLong", 3, vbLong
TestConst "vbSingle", 4, vbSingle
TestConst "vbDouble", 5, vbDouble
TestConst "vbCurrency", 6, vbCurrency
TestConst "vbDate", 7, vbDate
TestConst "vbString", 8, vbString
TestConst "vbObject", 9, vbObject
TestConst "vbError", 10, vbError
TestConst "vbBoolean", 11, vbBoolean
TestConst "vbVariant", 12, vbVariant
TestConst "vbDataObject", 13, vbDataObject
TestConst "vbDecimal", 14, vbDecimal
TestConst "vbByte", 17, vbByte
TestConst "vbArray", 8192, vbArray
%>
</table>

<!-- ============================================================ -->
<div class="summary">
<h2 style="background:none;border:none;padding:0;margin:0;color:#fff;">
Summary / 测试汇总
</h2>
<p>Total Tests / 总测试数: <strong><%= totalTests %></strong></p>
<p>Passed / 通过: <span class="pass"><%= passCount %></span></p>
<p>Failed / 失败: <span class="fail"><%= failCount %></span></p>
<%
If failCount = 0 Then
    Response.Write "<p style=""font-size:1.3em;"">✅ All constants match! / 所有常量值一致！</p>"
Else
    Response.Write "<p style=""font-size:1.3em;"">❌ " & failCount & " constant(s) mismatch! Please check above. / 有 " & failCount & " 个常量不一致，请检查上方标红项！</p>"
End If
%>
</div>

<p class="note" style="margin-top:30px;">
Source / 来源: Script 5.6 CHM (en/ + cn/ combined) — 9 constant pages under <code>vscon*.htm</code><br>
Total: 9 + 3 + 8 + 2 + 11 + 5 + 1 + 16 + 7 + 17 = <strong>79</strong>
</p>

</body>
</html>
