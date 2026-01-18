<%@ LANGUAGE="VBScript" %>
<% debug_asp_code = "FALSE" %>
<!DOCTYPE html>
<html>
<head>
    <title>VBScript Functions Test</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .section { margin: 20px 0; padding: 10px; border: 1px solid #ccc; }
        .pass { color: green; }
        .fail { color: red; }
        h2 { border-bottom: 2px solid #333; padding-bottom: 5px; }
        pre { background: #f4f4f4; padding: 10px; overflow-x: auto; }
    </style>
</head>
<body>
    <h1>VBScript Functions Test Suite</h1>
    
    <div class="section">
        <h2>String Functions</h2>
        <pre>
<% 
    Response.Write "LEN('hello'): " & Len("hello") & " (expected: 5)" & vbCrLf
    Response.Write "LEFT('hello', 3): " & Left("hello", 3) & " (expected: hel)" & vbCrLf
    Response.Write "RIGHT('hello', 2): " & Right("hello", 2) & " (expected: lo)" & vbCrLf
    Response.Write "MID('hello', 2, 3): " & Mid("hello", 2, 3) & " (expected: ell)" & vbCrLf
    Response.Write "INSTR('hello world', 'world'): " & InStr("hello world", "world") & " (expected: 7)" & vbCrLf
    Response.Write "INSTRREV('hello hello', 'hello'): " & InStrRev("hello hello", "hello") & " (expected: 7)" & vbCrLf
    Response.Write "REPLACE('hello world', 'world', 'there'): " & Replace("hello world", "world", "there") & vbCrLf
    Response.Write "TRIM('  hello  '): '" & Trim("  hello  ") & "' (expected: 'hello')" & vbCrLf
    Response.Write "LTRIM('  hello'): '" & LTrim("  hello") & "' (expected: 'hello')" & vbCrLf
    Response.Write "RTRIM('hello  '): '" & RTrim("hello  ") & "' (expected: 'hello')" & vbCrLf
    Response.Write "LCASE('HELLO'): " & LCase("HELLO") & " (expected: hello)" & vbCrLf
    Response.Write "UCASE('hello'): " & UCase("hello") & " (expected: HELLO)" & vbCrLf
    Response.Write "SPACE(5): '" & Space(5) & "' (expected: 5 spaces)" & vbCrLf
    Response.Write "STRING(3, 'x'): " & String(3, "x") & " (expected: xxx)" & vbCrLf
    Response.Write "STRREVERSE('hello'): " & StrReverse("hello") & " (expected: olleh)" & vbCrLf
    Response.Write "STRCOMP('abc', 'abc'): " & StrComp("abc", "abc") & " (expected: 0)" & vbCrLf
    Response.Write "STRCOMP('abc', 'def'): " & StrComp("abc", "def") & " (expected: -1)" & vbCrLf
    Response.Write "ASC('A'): " & Asc("A") & " (expected: 65)" & vbCrLf
    Response.Write "CHR(65): " & Chr(65) & " (expected: A)" & vbCrLf
    Response.Write "HEX(255): " & Hex(255) & " (expected: FF)" & vbCrLf
    Response.Write "OCT(8): " & Oct(8) & " (expected: 10)" & vbCrLf
%>
        </pre>
    </div>
    
    <div class="section">
        <h2>Math Functions</h2>
        <pre>
<% 
    Response.Write "ABS(-42): " & Abs(-42) & " (expected: 42)" & vbCrLf
    Response.Write "SQR(16): " & Sqr(16) & " (expected: 4)" & vbCrLf
    Response.Write "ROUND(3.7): " & Round(3.7) & " (expected: 4)" & vbCrLf
    Response.Write "ROUND(3.14159, 2): " & Round(3.14159, 2) & " (expected: 3.14)" & vbCrLf
    Response.Write "INT(3.9): " & Int(3.9) & " (expected: 3)" & vbCrLf
    Response.Write "INT(-3.9): " & Int(-3.9) & " (expected: -4)" & vbCrLf
    Response.Write "FIX(3.9): " & Fix(3.9) & " (expected: 3)" & vbCrLf
    Response.Write "FIX(-3.9): " & Fix(-3.9) & " (expected: -3)" & vbCrLf
    Response.Write "SGN(-5): " & Sgn(-5) & " (expected: -1)" & vbCrLf
    Response.Write "SGN(0): " & Sgn(0) & " (expected: 0)" & vbCrLf
    Response.Write "SGN(5): " & Sgn(5) & " (expected: 1)" & vbCrLf
    Response.Write "SIN(0): " & Sin(0) & " (expected: 0)" & vbCrLf
    Response.Write "COS(0): " & Cos(0) & " (expected: 1)" & vbCrLf
    Response.Write "TAN(0): " & Tan(0) & " (expected: 0)" & vbCrLf
    Response.Write "LOG(1): " & Log(1) & " (expected: 0)" & vbCrLf
    Response.Write "EXP(0): " & Exp(0) & " (expected: 1)" & vbCrLf
    Response.Write "RND(): " & (RND() >= 0 And RND() <= 1) & " (expected: True)" & vbCrLf
    Response.Write "ATN(0): " & Atn(0) & " (expected: 0)" & vbCrLf
%>
        </pre>
    </div>
    
    <div class="section">
        <h2>Date/Time Functions</h2>
        <pre>
<% 
    Response.Write "YEAR(NOW()): " & Year(Now()) & vbCrLf
    Response.Write "MONTH(NOW()): " & Month(Now()) & vbCrLf
    Response.Write "DAY(NOW()): " & Day(Now()) & vbCrLf
    Response.Write "HOUR(NOW()): " & Hour(Now()) & vbCrLf
    Response.Write "MINUTE(NOW()): " & Minute(Now()) & vbCrLf
    Response.Write "SECOND(NOW()): " & Second(Now()) & vbCrLf
    Response.Write "WEEKDAY(NOW()): " & Weekday(Now()) & " (1=Sunday to 7=Saturday)" & vbCrLf
    Response.Write "TIMER(): " & Timer() & " (seconds since midnight)" & vbCrLf
    Response.Write "WEEKDAYNAME(1): " & WeekdayName(1) & " (expected: Sunday)" & vbCrLf
    Response.Write "MONTHNAME(1): " & MonthName(1) & " (expected: January)" & vbCrLf
    Response.Write "TIME(): " & Time() & vbCrLf
    Response.Write "DATE(): " & Date() & vbCrLf
%>
        </pre>
    </div>
    
    <div class="section">
        <h2>Type Conversion Functions</h2>
        <pre>
<% 
    Response.Write "CINT('42'): " & CInt("42") & " (expected: 42)" & vbCrLf
    Response.Write "CDBL('3.14'): " & CDbl("3.14") & " (expected: 3.14)" & vbCrLf
    Response.Write "CSTR(42): " & CStr(42) & " (expected: 42)" & vbCrLf
    Response.Write "CBOOL(1): " & CBool(1) & " (expected: True)" & vbCrLf
    Response.Write "CBOOL(0): " & CBool(0) & " (expected: False)" & vbCrLf
    Response.Write "CBYTE(255): " & CByte(255) & " (expected: 255)" & vbCrLf
    Response.Write "CCUR(19.99): " & CCur(19.99) & " (expected: 19.99)" & vbCrLf
    Response.Write "CLNG(42): " & CLng(42) & " (expected: 42)" & vbCrLf
    Response.Write "CSNG(3.14): " & CSng(3.14) & vbCrLf
%>
        </pre>
    </div>
    
    <div class="section">
        <h2>Type Checking Functions</h2>
        <pre>
<% 
    Response.Write "ISEMPTY(Empty): " & IsEmpty(Empty) & " (expected: True)" & vbCrLf
    Response.Write "ISNUMERIC('42'): " & IsNumeric("42") & " (expected: True)" & vbCrLf
    Response.Write "ISNUMERIC('hello'): " & IsNumeric("hello") & " (expected: False)" & vbCrLf
    Response.Write "ISDATE('01/01/2020'): " & IsDate("01/01/2020") & " (expected: True)" & vbCrLf
    Response.Write "ISARRAY(ARRAY(1,2,3)): " & IsArray(Array(1,2,3)) & " (expected: True)" & vbCrLf
    Response.Write "ISOBJECT(CreateObject('G3JSON')): " & IsObject(CreateObject("G3JSON")) & " (expected: True)" & vbCrLf
%>
        </pre>
    </div>
    
    <div class="section">
        <h2>Formatting Functions</h2>
        <pre>
<% 
    Response.Write "FORMATCURRENCY(19.99): " & FormatCurrency(19.99) & vbCrLf
    Response.Write "FORMATCURRENCY(19.9, 1): " & FormatCurrency(19.9, 1) & vbCrLf
    Response.Write "FORMATNUMBER(3.14159, 2): " & FormatNumber(3.14159, 2) & vbCrLf
    Response.Write "FORMATPERCENT(0.75, 0): " & FormatPercent(0.75, 0) & vbCrLf
%>
        </pre>
    </div>
    
    <div class="section">
        <h2>Color Function</h2>
        <pre>
<% 
    Response.Write "RGB(255, 0, 0): " & RGB(255, 0, 0) & " (expected: 255 for red)" & vbCrLf
    Response.Write "RGB(0, 255, 0): " & RGB(0, 255, 0) & " (expected: 65280 for green)" & vbCrLf
    Response.Write "RGB(0, 0, 255): " & RGB(0, 0, 255) & " (expected: 16711680 for blue)" & vbCrLf
%>
        </pre>
    </div>
    
    <hr>
    <p><small>Test completed at: <% Response.Write Now() %></small></p>
</body>
</html>
