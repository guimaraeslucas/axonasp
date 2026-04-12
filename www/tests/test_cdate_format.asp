<%
@ Language = "VBScript"
%>
<%
Response.ContentType = "text/html; charset=utf-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"
%>
<!DOCTYPE html>
<html>
    <head>
        <title>CDate Format Test</title>
        <meta charset="utf-8" />
    </head>
    <body>
        <h1>CDate Format Test</h1>

        <%
        ' Test 1: Basic CDate conversion and display
        Dim d
        d = "01/02/2026"
        Response.Write "<h2>Test 1: Basic CDate Conversion</h2>"
        Response.Write "Input string: '01/02/2026'<br>"
        Response.Write "CDate(d) result: " & CDate(d) & "<br>"
        Response.Write "Expected: Date only, no time (locale dependent)<br>"
        Response.Write "<hr>"

        ' Test 2: CStr on a date
        Response.Write "<h2>Test 2: CStr on Date</h2>"
        Dim dateVal
        dateVal = CDate("15/03/2026")
        Response.Write "CStr(CDate('15/03/2026')): " & CStr(dateVal) & "<br>"
        Response.Write "Expected: Date only without time<br>"
        Response.Write "<hr>"

        ' Test 3: Concatenation with &
        Response.Write "<h2>Test 3: String Concatenation with &</h2>"
        Dim anotherDate
        anotherDate = CDate("2026-05-20")
        Response.Write "Concatenation: 'Today is: ' & CDate('2026-05-20')<br>"
        Response.Write "Result: Today is: " & anotherDate & "<br>"
        Response.Write "Expected: Should show date only, no time<br>"
        Response.Write "<hr>"

        ' Test 4: Now() function
        Response.Write "<h2>Test 4: Now() Function</h2>"
        Dim nowVal
        nowVal = Now()
        Response.Write "Now(): " & nowVal & "<br>"
        Response.Write "Expected: Date and time (contains time)<br>"
        Response.Write "<hr>"

        ' Test 5: Date() function
        Response.Write "<h2>Test 5: Date() Function</h2>"
        Dim dateOnlyVal
        dateOnlyVal = Date()
        Response.Write "Date(): " & dateOnlyVal & "<br>"
        Response.Write "Expected: Date only, no time<br>"
        Response.Write "<hr>"

        %>

        <p><a href="/default.asp">Back to Home</a></p>
    </body>
</html>
