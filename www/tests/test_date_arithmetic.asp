<%@ Language=VBScript %>
<%
' Test: arithmetic and numeric conversion on Date values use the day serial
' (days since 12/30/1899, time of day as the fraction), as on IIS.
' Expected values measured on IIS 10.0 with saborrie/iis-oracle, runs
' 45592970246, 45594216582, 45595451489 and 45595820636 (issue #236).
' Each check compares TypeName and the serial, so it does not depend on
' the server's date format.
Option Explicit

Dim testCount, passCount
testCount = 0
passCount = 0

Sub Check(name, actual, expectedType, expectedSerial)
	Dim actualType, ok
	testCount = testCount + 1
	actualType = TypeName(actual)
	ok = (actualType = expectedType)
	If ok Then ok = (Abs(CDbl(actual) - expectedSerial) < 0.000001)
	If ok Then
		passCount = passCount + 1
		Response.Write "<div class='pass'>PASS: " & Server.HTMLEncode(name) & "</div>"
	Else
		Response.Write "<div class='fail'>FAIL: " & Server.HTMLEncode(name) & " — expected " & expectedType & " " & expectedSerial & ", got " & actualType & " " & Server.HTMLEncode(CStr(actual)) & "</div>"
	End If
End Sub

Sub CheckBool(name, actual, expected)
	testCount = testCount + 1
	If actual = expected Then
		passCount = passCount + 1
		Response.Write "<div class='pass'>PASS: " & Server.HTMLEncode(name) & "</div>"
	Else
		Response.Write "<div class='fail'>FAIL: " & Server.HTMLEncode(name) & " — expected " & expected & ", got " & actual & "</div>"
	End If
End Sub

Dim d, e, p
d = #2027-01-01 10:30:00#
e = #2026-12-31#
p = #1800-01-01 06:00:00#
%>
<!DOCTYPE html>
<html>
    <head>
        <title>Date Arithmetic Test</title>
        <style>
            body { font-family: monospace; background: #1e1e1e; color: #d4d4d4; padding: 20px; }
            .pass { color: #4ec9b0; }
            .fail { color: #f44747; }
            .summary { margin-top: 20px; font-weight: bold; font-size: 1.2em; }
            .summary.pass { color: #4ec9b0; }
            .summary.fail { color: #f44747; }
        </style>
    </head>
    <body>
        <h1>Date Arithmetic Test (#236)</h1>
<%
' Date +/- number is a Date; Date - Date is a Double.
Check "d + 1", d + 1, "Date", 46389.4375
Check "1 + d", 1 + d, "Date", 46389.4375
Check "d - 1", d - 1, "Date", 46387.4375
Check "d + 0.5", d + 0.5, "Date", 46388.9375
Check "1 - d", 1 - d, "Date", -46387.4375
Check "d + ""1""", d + "1", "Date", 46389.4375
Check "d - e", d - e, "Double", 1.4375
Check "d + 3000000 (out of Date range)", d + 3000000, "Double", 3046388.4375
' *, / give Doubles; \, Mod give whole numbers.
Check "d * 1", d * 1, "Double", 46388.4375
Check "d / 1", d / 1, "Double", 46388.4375
CheckBool "d \ 1 = 46388", (d \ 1) = 46388, True
CheckBool "d Mod 7 = 6", (d Mod 7) = 6, True
' Unary minus, Int, Fix, Round and Abs keep the Date subtype.
Check "-d", -d, "Date", -46388.4375
Check "Int(d)", Int(d), "Date", 46388
Check "Fix(d)", Fix(d), "Date", 46388
Check "Round(d)", Round(d), "Date", 46388
Check "Abs(-d)", Abs(-d), "Date", 46388.4375
' Conversions read the serial.
Check "CDbl(d)", CDbl(d), "Double", 46388.4375
CheckBool "CLng(e) = 46387", CLng(e) = 46387, True
Check "CDbl(#1899-12-30#)", CDbl(#1899-12-30#), "Double", 0
Check "CDate(d + 1)", CDate(d + 1), "Date", 46389.4375
' Before 1899-12-30 the serial is negative and the fraction is still the time of day.
Check "CDbl(p)", CDbl(p), "Double", -36522.25
Check "p + 0.5", p + 0.5, "Date", -36521.75
Check "Int(p)", Int(p), "Date", -36523
' Comparisons and date functions.
CheckBool "d < 46389", d < 46389, True
CheckBool "d = e + 1.4375", d = e + 1.4375, True
CheckBool "d + 1 - 1 = d", d + 1 - 1 = d, True
CheckBool "DateDiff(""d"", d, d + 1) = 1", DateDiff("d", d, d + 1) = 1, True
CheckBool "Year(d + 1) = 2027", Year(d + 1) = 2027, True
CheckBool "Hour(d + 0.5) = 22", Hour(d + 0.5) = 22, True
CheckBool "TypeName(Date - 1) = ""Date""", TypeName(Date - 1) = "Date", True
CheckBool "Int(Now + 1) = Date + 1", Int(Now + 1) = Date + 1, True
%>
        <div class='summary <% If passCount = testCount Then Response.Write "pass" Else Response.Write "fail" %>'>
            <%= passCount %> / <%= testCount %> tests passed
        </div>
    </body>
</html>
