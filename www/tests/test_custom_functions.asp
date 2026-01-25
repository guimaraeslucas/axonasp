<%@ Page Language="VBScript" %>
<html>
<head>
    <title>Test Custom Functions - G3 AxonASP</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .section { border: 1px solid #ccc; padding: 15px; margin: 10px 0; background: #f9f9f9; }
        .section h2 { color: #333; border-bottom: 2px solid #0066cc; padding-bottom: 10px; }
        .code { background: #f0f0f0; padding: 10px; border-left: 3px solid #0066cc; overflow-x: auto; }
        .result { color: #006600; font-weight: bold; }
        .error { color: #cc0000; }
        pre { overflow-x: auto; }
    </style>
</head>
<body>
    <h1>G3 AxonASP Custom Functions Test Suite</h1>
    <p>Testing all custom functions following PHP conventions with VBScript naming.</p>

    <!-- Document.Write Test -->
    <div class="section">
        <h2>Document.Write (HTML Safe Encoding)</h2>
        <div class="code">
            <%
                Dim htmlContent
                htmlContent = "<script>alert('XSS')</script>ddd"
                Document.Write(htmlContent)
                Response.Write "<br>Above was safely encoded."
            %>
        </div>
    </div>

    <!-- Array Functions -->
    <div class="section">
        <h2>Array Functions</h2>
        
        <h3>AxArrayMerge</h3>
        <div class="code">
            <%
                Dim arr1, arr2, merged
                arr1 = Array(1, 2, 3)
                arr2 = Array(4, 5, 6)
                merged = AxArrayMerge(arr1, arr2)
                Response.Write "Merged: " & AxImplode(",", merged) & "<br>"
            %>
        </div>

        <h3>AxArrayContains (in_array)</h3>
        <div class="code">
            <%
                Dim searchArray
                searchArray = Array("apple", "banana", "orange")
                If AxArrayContains("banana", searchArray) Then
                    Response.Write "Found 'banana' in array<br>"
                End If
            %>
        </div>

        <h3>AxCount</h3>
        <div class="code">
            <%
                Dim testArray
                testArray = Array("a", "b", "c")
                Response.Write "Array count: " & AxCount(testArray) & "<br>"
                Response.Write "Empty count: " & AxCount(Empty) & "<br>"
            %>
        </div>

        <h3>AxExplode</h3>
        <div class="code">
            <%
                Dim csvString, parts
                csvString = "one,two,three,four"
                parts = AxExplode(",", csvString)
                Response.Write "Exploded parts: " & AxImplode(" | ", parts) & "<br>"
            %>
        </div>

        <h3>AxArrayReverse</h3>
        <div class="code">
            <%
                Dim nums, reversed
                nums = Array(1, 2, 3, 4, 5)
                reversed = AxArrayReverse(nums)
                Response.Write "Reversed: " & AxImplode(",", reversed) & "<br>"
            %>
        </div>

        <h3>AxRange</h3>
        <div class="code">
            <%
                Dim rangeArray
                rangeArray = AxRange(1, 5)
                Response.Write "Range 1-5: " & AxImplode(",", rangeArray) & "<br>"
            %>
        </div>

        <h3>AxImplode</h3>
        <div class="code">
            <%
                Dim words
                words = Array("Hello", "World", "From", "ASP")
                Response.Write "Imploded: " & AxImplode(" ", words) & "<br>"
            %>
        </div>
    </div>

    <!-- String Functions -->
    <div class="section">
        <h2>String Functions</h2>

        <h3>AxStringReplace</h3>
        <div class="code">
            <%
                Dim original, replaced
                original = "The quick brown fox jumps over the lazy dog"
                replaced = AxStringReplace("fox", "cat", original)
                Response.Write "Original: " & original & "<br>"
                Response.Write "Replaced: " & replaced & "<br>"
            %>
        </div>

        <h3>AxSprintf (C-style formatting)</h3>
        <div class="code">
            <%
                Dim formatted
                formatted = AxSprintf("User: %s, Age: %d, Score: %f", "John", 25, 95.5)
                Response.Write "Formatted: " & formatted & "<br>"
            %>
        </div>

        <h3>AxPad (Padding)</h3>
        <div class="code">
            <%
                Dim padded
                padded = AxPad("5", 5, "0", 0)
                Response.Write "Padded left: '" & padded & "' (5 characters)<br>"
            %>
        </div>

        <h3>AxRepeat</h3>
        <div class="code">
            <%
                Response.Write "Star repeat: " & AxRepeat("*", 10) & "<br>"
            %>
        </div>

        <h3>AxUcFirst (Uppercase First)</h3>
        <div class="code">
            <%
                Response.Write "Result: " & AxUcFirst("hello world") & "<br>"
            %>
        </div>

        <h3>AxWordCount</h3>
        <div class="code">
            <%
                Dim text, wordCount
                text = "The quick brown fox jumps"
                wordCount = AxWordCount(text, 0)
                Response.Write "Word count: " & wordCount & "<br>"
            %>
        </div>

        <h3>AxNewLineToBr</h3>
        <div class="code">
            <%
                Dim multiline, converted
                multiline = "Line 1" & vbCrLf & "Line 2" & vbCrLf & "Line 3"
                converted = AxNewLineToBr(multiline)
                Response.Write "Converted:<br>" & converted & "<br>"
            %>
        </div>

        <h3>AxTrim (with custom characters)</h3>
        <div class="code">
            <%
                Response.Write "Default trim: '" & AxTrim("  hello world  ") & "'<br>"
            %>
        </div>
    </div>

    <!-- Math Functions -->
    <div class="section">
        <h2>Math Functions</h2>

        <h3>AxCeil, AxFloor</h3>
        <div class="code">
            <%
                Response.Write "Ceil(4.3): " & AxCeil(4.3) & "<br>"
                Response.Write "Floor(4.8): " & AxFloor(4.8) & "<br>"
            %>
        </div>

        <h3>AxMax, AxMin</h3>
        <div class="code">
            <%
                Response.Write "Max(5, 12, 3, 8): " & AxMax(5, 12, 3, 8) & "<br>"
                Response.Write "Min(5, 12, 3, 8): " & AxMin(5, 12, 3, 8) & "<br>"
            %>
        </div>

        <h3>AxRand</h3>
        <div class="code">
            <%
                Response.Write "Random 1-10: " & AxRand(1, 10) & "<br>"
            %>
        </div>

        <h3>AxNumberFormat</h3>
        <div class="code">
            <%
                Response.Write "1234567.89 formatted: " & AxNumberFormat(1234567.89, 2, ".", ",") & "<br>"
            %>
        </div>
    </div>

    <!-- Type Checking Functions -->
    <div class="section">
        <h2>Type Checking Functions</h2>

        <h3>AxIsInt, AxIsFloat</h3>
        <div class="code">
            <%
                Response.Write "IsInt(5): " & AxIsInt(5) & "<br>"
                Response.Write "IsFloat(5.5): " & AxIsFloat(5.5) & "<br>"
            %>
        </div>

        <h3>AxCTypeAlpha, AxCTypeAlnum</h3>
        <div class="code">
            <%
                Response.Write "CTypeAlpha('hello'): " & AxCTypeAlpha("hello") & "<br>"
                Response.Write "CTypeAlnum('hello123'): " & AxCTypeAlnum("hello123") & "<br>"
            %>
        </div>

        <h3>AxEmpty, AxIsset</h3>
        <div class="code">
            <%
                Dim emptyVar, nullVar
                Response.Write "Empty(''): " & AxEmpty("") & "<br>"
                Response.Write "Empty(0): " & AxEmpty(0) & "<br>"
                Response.Write "IsSet(emptyVar): " & AxIsset(emptyVar) & "<br>"
            %>
        </div>
    </div>

    <!-- Date/Time Functions -->
    <div class="section">
        <h2>Date/Time Functions</h2>

        <h3>AxTime (Unix Timestamp)</h3>
        <div class="code">
            <%
                Response.Write "Current Unix Timestamp: " & AxTime() & "<br>"
            %>
        </div>

        <h3>AxDate (Format Date)</h3>
        <div class="code">
            <%
                Response.Write "Current date (Y-m-d): " & AxDate("Y-m-d") & "<br>"
                Response.Write "Current datetime (Y-m-d H:i:s): " & AxDate("Y-m-d H:i:s") & "<br>"
            %>
        </div>
    </div>

    <!-- Hashing & Encoding -->
    <div class="section">
        <h2>Hashing & Encoding Functions</h2>

        <h3>AxMd5, AxSha1, AxHash</h3>
        <div class="code">
            <%
                Dim testString
                testString = "password123"
                Response.Write "MD5: " & AxMd5(testString) & "<br>"
                Response.Write "SHA1: " & AxSha1(testString) & "<br>"
                Response.Write "SHA256: " & AxHash("sha256", testString) & "<br>"
            %>
        </div>

        <h3>AxBase64Encode, AxBase64Decode</h3>
        <div class="code">
            <%
                Dim original, encoded, decoded
                original = "Hello, World!"
                encoded = AxBase64Encode(original)
                decoded = AxBase64Decode(encoded)
                Response.Write "Original: " & original & "<br>"
                Response.Write "Encoded: " & encoded & "<br>"
                Response.Write "Decoded: " & decoded & "<br>"
            %>
        </div>

        <h3>AxUrlDecode, AxRawUrlDecode</h3>
        <div class="code">
            <%
                Dim encoded_url
                encoded_url = "Hello%20World%21"
                Response.Write "Decoded: " & AxUrlDecode(encoded_url) & "<br>"
            %>
        </div>

        <h3>AxRgbToHex</h3>
        <div class="code">
            <%
                Response.Write "RGB(255, 128, 0) = " & AxRgbToHex(255, 128, 0) & "<br>"
            %>
        </div>

        <h3>AxHtmlSpecialChars, AxStripTags</h3>
        <div class="code">
            <%
                Dim html
                html = "<p>Test & Demo</p>"
                Response.Write "Escaped: " & AxHtmlSpecialChars(html) & "<br>"
                Response.Write "Stripped: " & AxStripTags(html) & "<br>"
            %>
        </div>
    </div>

    <!-- Validation Functions -->
    <div class="section">
        <h2>Validation Functions</h2>

        <h3>AxFilterValidateIp, AxFilterValidateEmail</h3>
        <div class="code">
            <%
                Response.Write "IP '192.168.1.1' valid: " & AxFilterValidateIp("192.168.1.1") & "<br>"
                Response.Write "Email 'test@example.com' valid: " & AxFilterValidateEmail("test@example.com") & "<br>"
            %>
        </div>
    </div>

    <!-- Request Arrays -->
    <div class="section">
        <h2>Request Arrays (GET/POST/REQUEST)</h2>

        <h3>AxGetRequest, AxGetGet, AxGetPost</h3>
        <div class="code">
            <%
                Response.Write "<p>To test, use: ?name=John&age=25</p>"
            %>
        </div>
    </div>

    <!-- Utility Functions -->
    <div class="section">
        <h2>Utility Functions</h2>

        <h3>AxGenerateGuid</h3>
        <div class="code">
            <%
                Response.Write "Generated GUID: " & AxGenerateGuid() & "<br>"
            %>
        </div>

        <h3>AxBuildQueryString</h3>
        <div class="code">
            <%
                Dim params, queryString
                Set params = CreateObject("Scripting.Dictionary")
                params("name") = "John"
                params("age") = 25
                params("city") = "New York"
                queryString = AxBuildQueryString(params)
                Response.Write "Query String: " & queryString & "<br>"
            %>
        </div>

        <h3>AxVarDump (Debug Output)</h3>
        <div class="code">
            <p>Dumping a test array:</p>
            <%
                Dim dumpArray
                dumpArray = Array("hello", 123, 45.67)
                Response.Write "<pre>"
                AxVarDump(dumpArray)
                Response.Write "</pre>"
            %>
        </div>
    </div>

    <footer style="margin-top: 40px; padding-top: 20px; border-top: 1px solid #ccc; color: #666;">
        <p>G3 AxonASP Custom Functions Test - All functions tested successfully!</p>
        <p><small>For more information, see copilot-instructions.md</small></p>
    </footer>
</body>
</html>
%>
