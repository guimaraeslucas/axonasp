<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Standard Functions Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background: #f5f5f5; color: #333; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Standard VBScript Functions Test</h1>
        <div class="intro">
            <p>Tests date/time, string, math, type conversion and formatting functions.</p>
        </div>

<% Response.Write("<h2>Data/Hora</h2>")
Response.Write("Now: " & Now() & "<br>")
Response.Write("Date: " & Date() & "<br>")
Response.Write("Time: " & Time() & "<br>")
Response.Write("Year: " & Year(Now()) & "<br>")
Response.Write("Month: " & Month(Date()) & "<br>")
Response.Write("Day: " & Day(Date()) & "<br>")
Response.Write("Hour: " & Hour(Time()) & "<br>")
Response.Write("Minute: " & Minute(Time()) & "<br>")
Response.Write("Second: " & Second(Time()) & "<br>")
Response.Write("Weekday: " & Weekday(2) & "<br>")
Response.Write("WeekdayName: " & WeekdayName(1) & "<br>")
Response.Write("MonthName: " & MonthName(1) & "<br>")
Response.Write("Timer: " & Timer & "<br>")
Response.Write("DateValue('01/01/2025'): " & DateValue("01/01/2025") & "<br>")
Response.Write("TimeValue('12:30:45'): " & TimeValue("12:30:45") & "<br>")
Response.Write("DateAdd('d', 5, '01/01/2025'): " & DateAdd("d", 5, "01/01/2025") & "<br>")
Response.Write("DateDiff('d', '01/01/2025', '01/31/2025'): " & DateDiff("d", "01/01/2025", "01/31/2025") & "<br>")
Response.Write("DatePart('m', '03/15/2025'): " & DatePart("m", "03/15/2025") & "<br>")
Response.Write("DateSerial(2025, 6, 15): " & DateSerial(2025, 6, 15) & "<br>")
Response.Write("TimeSerial(14, 30, 45): " & TimeSerial(14, 30, 45) & "<br>")
Response.Write("<br>")

' String Functions
Response.Write("<h2>String</h2>")
Response.Write("Len('Hello'): " & Len("Hello") & "<br>")
Response.Write("Left('Hello', 2): " & Left("Hello", 2) & "<br>")
Response.Write("Right('Hello', 2): " & Right("Hello", 2) & "<br>")
Response.Write("Mid('Hello', 2, 3): " & Mid("Hello", 2, 3) & "<br>")
Response.Write("InStr('Hello World', 'World'): " & InStr("Hello World", "World") & "<br>")
Response.Write("InStrRev('Hello World', 'o'): " & InStrRev("Hello World", "o") & "<br>")
Response.Write("Replace('Hello', 'l', 'x'): " & Replace("Hello", "l", "x") & "<br>")
Response.Write("Trim('  Hello  '): [" & Trim("  Hello  ") & "]<br>")
Response.Write("LTrim('  Hello'): [" & LTrim("  Hello") & "]<br>")
Response.Write("RTrim('Hello  '): [" & RTrim("Hello  ") & "]<br>")
Response.Write("LCase('HELLO'): " & LCase("HELLO") & "<br>")
Response.Write("UCase('hello'): " & UCase("hello") & "<br>")
Response.Write("Space(5): [" & Space(5) & "]<br>")
Response.Write("String(3, 'a'): " & String(3, "a") & "<br>")
Response.Write("StrReverse('Hello'): " & StrReverse("Hello") & "<br>")
Response.Write("StrComp('abc', 'abc'): " & StrComp("abc", "abc") & "<br>")

' Math Functions
Response.Write("<h2>Matemática</h2>")
Response.Write("Abs(-5): " & Abs(-5) & "<br>")
Response.Write("Sqr(9): " & Sqr(9) & "<br>")
Response.Write("Rnd: " & Rnd & "<br>")
Response.Write("Round(3.7): " & Round(3.7) & "<br>")
Response.Write("Int(3.7): " & Int(3.7) & "<br>")
Response.Write("Sin(0): " & Sin(0) & "<br>")
Response.Write("Cos(0): " & Cos(0) & "<br>")
Response.Write("Tan(0): " & Tan(0) & "<br>")
Response.Write("Atn(1): " & Atn(1) & "<br>")
Response.Write("Log(2.718): " & Log(2.718) & "<br>")
Response.Write("Exp(1): " & Exp(1) & "<br>")
Response.Write("Sgn(-5): " & Sgn(-5) & "<br>")
Response.Write("Sgn(0): " & Sgn(0) & "<br>")
Response.Write("Sgn(5): " & Sgn(5) & "<br>")
Response.Write("Fix(-3.7): " & Fix(-3.7) & "<br>")

' Conversion Functions
Response.Write("<h2>Conversão</h2>")
Response.Write("CInt('42'): " & CInt("42") & "<br>")
Response.Write("CDbl('3.14'): " & CDbl("3.14") & "<br>")
Response.Write("CStr(42): " & CStr(42) & "<br>")
Response.Write("CBool(1): " & CBool(1) & "<br>")
Response.Write("CDate('01/01/2025'): " & CDate("01/01/2025") & "<br>")
Response.Write("Asc('A'): " & Asc("A") & "<br>")
Response.Write("Chr(65): " & Chr(65) & "<br>")
Response.Write("Hex(255): " & Hex(255) & "<br>")
Response.Write("Oct(8): " & Oct(8) & "<br>")
Response.Write("CByte(200): " & CByte(200) & "<br>")
Response.Write("CCur(100.5): " & CCur(100.5) & "<br>")
Response.Write("CLng(42): " & CLng(42) & "<br>")
Response.Write("CSng(3.14): " & CSng(3.14) & "<br>")

' Format Functions
Response.Write("<h2>Formatação</h2>")
Response.Write("FormatCurrency(1234.5): " & FormatCurrency(1234.5) & "<br>")
Response.Write("FormatNumber(1234.5678, 2): " & FormatNumber(1234.5678, 2) & "<br>")
Response.Write("FormatPercent(0.25, 1): " & FormatPercent(0.25, 1) & "<br>")

' Array Functions
Response.Write("<h2>Array</h2>")
Dim arr
arr = Array("apple", "banana", "cherry")
Response.Write("Array created<br>")
Response.Write("UBound(arr): " & UBound(arr) & "<br>")
Response.Write("LBound(arr): " & LBound(arr) & "<br>")
Response.Write("IsArray(arr): " & IsArray(arr) & "<br>")
Response.Write("Join(arr, ', '): " & Join(arr, ", ") & "<br>")

Dim str
str = "apple,banana,cherry"
Dim parts
parts = Split(str, ",")
Response.Write("Split('apple,banana,cherry', ','): count=" & UBound(parts) + 1 & "<br>")

Response.Write("<h3 style='margin-top:20px'>Testes de Funções Completados!</h3>")
%>
    </div>
</body>
</html>
