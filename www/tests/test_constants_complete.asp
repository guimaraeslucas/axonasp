<%
@ Language = "VBScript"
%>
<%
' Test suite for VBScript constants
' Verifies vb* constants have the expected VBScript values

Dim failCount, passCount
failCount = 0
passCount = 0

Function DescribeString(value)
    Dim index, currentChar, currentCode, rendered
    rendered = ""

    For index = 1 To Len(value)
        currentChar = Mid(value, index, 1)
        currentCode = Asc(currentChar)

        Select Case currentCode
            Case 0
                rendered = rendered & "\0"
            Case 9
                rendered = rendered & "\t"
            Case 10
                rendered = rendered & "\n"
            Case 13
                rendered = rendered & "\r"
            Case Else
                rendered = rendered & currentChar
        End Select
    Next

    DescribeString = Chr(34) & rendered & Chr(34)
End Function

Function TestNumericConstant(name, actual, expected, comment)
    If CLng(actual) = CLng(expected) Then
        Response.Write "<div style='color:green;'>[PASS] " & name & " = " & actual & " (" & comment & ")</div>"
        passCount = passCount + 1
    Else
        Response.Write "<div style='color:red;'>[FAIL] " & name & " - Expected: " & expected & ", Got: " & actual & " (" & comment & ")</div>"
        failCount = failCount + 1
    End If
End Function

Function TestStringConstant(name, actual, expected, comment)
    If CStr(actual) = CStr(expected) Then
        Response.Write "<div style='color:green;'>[PASS] " & name & " = " & DescribeString(actual) & " (" & comment & ")</div>"
        passCount = passCount + 1
    Else
        Response.Write "<div style='color:red;'>[FAIL] " & name & " - Expected: " & DescribeString(expected) & ", Got: " & DescribeString(actual) & " (" & comment & ")</div>"
        failCount = failCount + 1
    End If
End Function

%>

<html>
    <head>
        <title>VBScript Constants Complete Test</title>
        <style>
            body {
                font-family: Tahoma, Verdana;
                background-color: #ece9d8;
                margin: 20px;
            }
            h1 {
                color: #003399;
                border-bottom: 2px solid #335ea8;
                padding-bottom: 10px;
            }
            .section {
                margin: 20px 0;
                border: 1px solid #808080;
                padding: 10px;
                background-color: #f0f0f0;
            }
            h2 {
                background-color: #e0e0e0;
                padding: 5px;
                margin: 0;
            }
            .results {
                margin-top: 10px;
            }
            .summary {
                background-color: #d0e0ff;
                padding: 10px;
                margin-top: 20px;
                border: 1px solid #335ea8;
            }
        </style>
    </head>
    <body>
        <h1>VBScript Constants Compliance Test Suite</h1>
        <p>
            Testing implementation of all VBScript vb* constants per Microsoft
            VBScript 5.8 specification.
        </p>

        <div class="section">
            <h2>Color Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbBlack", vbBlack, 0, "Black color"
                TestNumericConstant "vbRed", vbRed, 16711680, "Red color"
                TestNumericConstant "vbGreen", vbGreen, 65280, "Green color"
                TestNumericConstant "vbYellow", vbYellow, 16776960, "Yellow color"
                TestNumericConstant "vbBlue", vbBlue, 255, "Blue color"
                TestNumericConstant "vbMagenta", vbMagenta, 16711935, "Magenta color"
                TestNumericConstant "vbCyan", vbCyan, 65535, "Cyan color"
                TestNumericConstant "vbWhite", vbWhite, 16777215, "White color"
                %>
            </div>
        </div>

        <div class="section">
            <h2>Boolean Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbTrue", vbTrue, -1, "Boolean True"
                TestNumericConstant "vbFalse", vbFalse, 0, "Boolean False"
                %>
            </div>
        </div>

        <div class="section">
            <h2>Comparison Mode Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbBinaryCompare", vbBinaryCompare, 0, "Binary string comparison"
                TestNumericConstant "vbTextCompare", vbTextCompare, 1, "Text string comparison"
                TestNumericConstant "vbDatabaseCompare", vbDatabaseCompare, 2, "Database comparison"
                %>
            </div>
        </div>

        <div class="section">
            <h2>Date Format Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbGeneralDate", vbGeneralDate, 0, "General date format"
                TestNumericConstant "vbLongDate", vbLongDate, 1, "Long date format"
                TestNumericConstant "vbShortDate", vbShortDate, 2, "Short date format"
                TestNumericConstant "vbLongTime", vbLongTime, 3, "Long time format"
                TestNumericConstant "vbShortTime", vbShortTime, 4, "Short time format"
                %>
            </div>
        </div>

        <div class="section">
            <h2>Date/Time Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbSunday", vbSunday, 1, "Sunday"
                TestNumericConstant "vbMonday", vbMonday, 2, "Monday"
                TestNumericConstant "vbTuesday", vbTuesday, 3, "Tuesday"
                TestNumericConstant "vbWednesday", vbWednesday, 4, "Wednesday"
                TestNumericConstant "vbThursday", vbThursday, 5, "Thursday"
                TestNumericConstant "vbFriday", vbFriday, 6, "Friday"
                TestNumericConstant "vbSaturday", vbSaturday, 7, "Saturday"
                TestNumericConstant "vbJanuary", vbJanuary, 1, "January"
                TestNumericConstant "vbFebruary", vbFebruary, 2, "February"
                TestNumericConstant "vbMarch", vbMarch, 3, "March"
                TestNumericConstant "vbApril", vbApril, 4, "April"
                TestNumericConstant "vbMay", vbMay, 5, "May"
                TestNumericConstant "vbJune", vbJune, 6, "June"
                TestNumericConstant "vbJuly", vbJuly, 7, "July"
                TestNumericConstant "vbAugust", vbAugust, 8, "August"
                TestNumericConstant "vbSeptember", vbSeptember, 9, "September"
                TestNumericConstant "vbOctober", vbOctober, 10, "October"
                TestNumericConstant "vbNovember", vbNovember, 11, "November"
                TestNumericConstant "vbDecember", vbDecember, 12, "December"
                %>
            </div>
        </div>

        <div class="section">
            <h2>First Week/Day Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbFirstJan1", vbFirstJan1, 1, "Week containing Jan 1"
                TestNumericConstant "vbFirstFourDays", vbFirstFourDays, 2, "Week containing 4 days"
                TestNumericConstant "vbFirstFullWeek", vbFirstFullWeek, 3, "First complete week"
                TestNumericConstant "vbUseSystemDayOfWeek", vbUseSystemDayOfWeek, 0, "System default"
                %>
            </div>
        </div>

        <div class="section">
            <h2>Encoding Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbASCII", vbASCII, 0, "ASCII format"
                TestNumericConstant "vbUnicode", vbUnicode, 64, "StrConv Unicode conversion flag"
                TestNumericConstant "vbFromUnicode", vbFromUnicode, 128, "StrConv ANSI conversion flag"
                %>
            </div>
        </div>

        <div class="section">
            <h2>File Input/Output Constants</h2>
            <div class="results">
                <%
                TestStringConstant "vbCrLf", vbCrLf, Chr(13) & Chr(10), "Carriage Return + Line Feed"
                TestStringConstant "vbCr", vbCr, Chr(13), "Carriage Return"
                TestStringConstant "vbLf", vbLf, Chr(10), "Line Feed"
                TestStringConstant "vbNewLine", vbNewLine, Chr(13) & Chr(10), "New line character"
                TestStringConstant "vbTab", vbTab, Chr(9), "Tab character"
                TestStringConstant "vbNullChar", vbNullChar, Chr(0), "Null character"
                TestStringConstant "vbNullString", vbNullString, "", "Null-length string"
                %>
            </div>
        </div>

        <div class="section">
            <h2>Misc Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbEmpty", vbEmpty, 0, "Empty variant type"
                TestNumericConstant "vbNull", vbNull, 1, "Null variant type"
                TestNumericConstant "vbObject", vbObject, 9, "Object variant type"
                TestNumericConstant "vbCurrency", vbCurrency, 6, "Currency variant type"
                TestNumericConstant "vbDate", vbDate, 7, "Date variant type"
                TestNumericConstant "vbString", vbString, 8, "String variant type"
                TestNumericConstant "vbError", vbError, 10, "Error variant type"
                TestNumericConstant "vbBoolean", vbBoolean, 11, "Boolean variant type"
                TestNumericConstant "vbVariant", vbVariant, 12, "Variant variant type"
                TestNumericConstant "vbDataObject", vbDataObject, 13, "Data object variant type"
                TestNumericConstant "vbDecimal", vbDecimal, 14, "Decimal variant type"
                TestNumericConstant "vbByte", vbByte, 17, "Byte variant type"
                TestNumericConstant "vbLong", vbLong, 3, "Long variant type"
                TestNumericConstant "vbSingle", vbSingle, 4, "Single variant type"
                TestNumericConstant "vbDouble", vbDouble, 5, "Double variant type"
                TestNumericConstant "vbInteger", vbInteger, 2, "Integer variant type"
                %>
            </div>
        </div>

        <div class="section">
            <h2>MsgBox Constants and Return Values</h2>
            <div class="results">
                <%
                ' MsgBox buttons
                TestNumericConstant "vbOKOnly", vbOKOnly, 0, "MsgBox OK button"
                TestNumericConstant "vbOKCancel", vbOKCancel, 1, "MsgBox OK/Cancel buttons"
                TestNumericConstant "vbAbortRetryIgnore", vbAbortRetryIgnore, 2, "MsgBox Abort/Retry/Ignore buttons"
                TestNumericConstant "vbYesNoCancel", vbYesNoCancel, 3, "MsgBox Yes/No/Cancel buttons"
                TestNumericConstant "vbYesNo", vbYesNo, 4, "MsgBox Yes/No buttons"
                TestNumericConstant "vbRetryCancel", vbRetryCancel, 5, "MsgBox Retry/Cancel buttons"

                ' MsgBox return values
                TestNumericConstant "vbOK", vbOK, 1, "MsgBox return OK"
                TestNumericConstant "vbCancel", vbCancel, 2, "MsgBox return Cancel"
                TestNumericConstant "vbAbort", vbAbort, 3, "MsgBox return Abort"
                TestNumericConstant "vbRetry", vbRetry, 4, "MsgBox return Retry"
                TestNumericConstant "vbIgnore", vbIgnore, 5, "MsgBox return Ignore"
                TestNumericConstant "vbYes", vbYes, 6, "MsgBox return Yes"
                TestNumericConstant "vbNo", vbNo, 7, "MsgBox return No"
                %>
            </div>
        </div>

        <div class="section">
            <h2>StrConv Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbUpperCase", vbUpperCase, 1, "Convert to uppercase"
                TestNumericConstant "vbLowerCase", vbLowerCase, 2, "Convert to lowercase"
                TestNumericConstant "vbProperCase", vbProperCase, 3, "Convert to proper case"
                TestNumericConstant "vbWide", vbWide, 4, "Convert single-byte to double-byte"
                TestNumericConstant "vbNarrow", vbNarrow, 8, "Convert double-byte to single-byte"
                TestNumericConstant "vbKatakana", vbKatakana, 16, "Convert Hiragana to Katakana"
                TestNumericConstant "vbHiragana", vbHiragana, 32, "Convert Katakana to Hiragana"
                TestNumericConstant "vbUnicode", vbUnicode, 64, "Convert to Unicode"
                TestNumericConstant "vbFromUnicode", vbFromUnicode, 128, "Convert from Unicode"
                %>
            </div>
        </div>

        <div class="section">
            <h2>VarType Return Values</h2>
            <div class="results">
                <%
                TestNumericConstant "vbEmpty", vbEmpty, 0, "Uninitialized"
                TestNumericConstant "vbNull", vbNull, 1, "Null"
                TestNumericConstant "vbInteger", vbInteger, 2, "Integer"
                TestNumericConstant "vbLong", vbLong, 3, "Long"
                TestNumericConstant "vbSingle", vbSingle, 4, "Single-precision floating-point"
                TestNumericConstant "vbDouble", vbDouble, 5, "Double-precision floating-point"
                TestNumericConstant "vbCurrency", vbCurrency, 6, "Currency"
                TestNumericConstant "vbDate", vbDate, 7, "Date"
                TestNumericConstant "vbString", vbString, 8, "String"
                TestNumericConstant "vbObject", vbObject, 9, "Object"
                TestNumericConstant "vbError", vbError, 10, "Error"
                TestNumericConstant "vbBoolean", vbBoolean, 11, "Boolean"
                TestNumericConstant "vbVariant", vbVariant, 12, "Variant"
                TestNumericConstant "vbDataObject", vbDataObject, 13, "Data object"
                TestNumericConstant "vbDecimal", vbDecimal, 14, "Decimal"
                TestNumericConstant "vbByte", vbByte, 17, "Byte"
                TestNumericConstant "vbArray", vbArray, 8192, "Array"
                %>
            </div>
        </div>

        <div class="section">
            <h2>Tristate Constants</h2>
            <div class="results">
                <%
                TestNumericConstant "vbUseDefault", vbUseDefault, -2, "Use system default"
                TestNumericConstant "vbTrue", vbTrue, -1, "True state"
                TestNumericConstant "vbFalse", vbFalse, 0, "False state"
                %>
            </div>
        </div>

        <div class="summary">
            <h3>Test Summary</h3>
            <p>
                <strong>Passed:</strong>
                <span style="color: green"
                    ><%= passCount %></span
                >
            </p>
            <p>
                <strong>Failed:</strong>
                <span style="color: red"><%= failCount %></span>
            </p>
            <p>
                <strong>Total:</strong>
                <%= passCount + failCount %>
            </p>
            <%
            If failCount = 0 Then
                Response.Write "<p style='color:green; font-weight:bold;'> All constants are properly implemented!</p>"
            Else
                Response.Write "<p style='color:red; font-weight:bold;'> " & failCount & " constant(s) need attention.</p>"
            End If
            %>
        </div>
    </body>
</html>
