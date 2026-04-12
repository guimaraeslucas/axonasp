<%
@ Language = "VBScript"
%>
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <title>Filter Function Test - AxonVM</title>
        <style>
            body {
                font-family: "Segoe UI", Arial;
                margin: 20px;
                background: #ece9d8;
            }
            h1 {
                color: #003399;
                border-bottom: 3px solid #003399;
                padding-bottom: 5px;
            }
            .test {
                background: white;
                border: 1px solid #808080;
                padding: 15px;
                margin: 10px 0;
            }
            .success {
                color: green;
                font-weight: bold;
            }
            .error {
                color: red;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <h1>Filter Function Test - AxonVM</h1>

        <div class="test">
            <h3>Test 1: Basic Filter</h3>
            <%
            Dim arr1, result1
            arr1 = Array("Hello World", "Goodbye", "Hello Friend")
            result1 = Filter(arr1, "Hello")
            Response.Write "Input: Array(""Hello World"", ""Goodbye"", ""Hello Friend"")<br>"
            Response.Write "Filter(arr1, ""Hello"")<br>"
            Response.Write "Result type: " & TypeName(result1) & "<br>"
            Response.Write "UBound: " & UBound(result1) & "<br>"
            Response.Write "result1(0): " & result1(0) & "<br>"
            Response.Write "result1(1): " & result1(1) & "<br>"
            If UBound(result1) = 1 And result1(0) = "Hello World" And result1(1) = "Hello Friend" Then
                Response.Write "<span class='success'>✓ PASS</span>"
            Else
                Response.Write "<span class='error'>✗ FAIL</span>"
            End If
            %>
        </div>

        <div class="test">
            <h3>Test 2: Filter with Include=False (Exclude mode)</h3>
            <%
            Dim arr2, result2
            arr2 = Array("Apple", "Banana", "Apricot", "Cherry")
            result2 = Filter(arr2, "Ap", False)
            Response.Write "Input: Array(""Apple"", ""Banana"", ""Apricot"", ""Cherry"")<br>"
            Response.Write "Filter(arr2, ""Ap"", False)<br>"
            Response.Write "UBound: " & UBound(result2) & "<br>"
            Response.Write "result2(0): " & result2(0) & "<br>"
            Response.Write "result2(1): " & result2(1) & "<br>"
            If UBound(result2) = 1 And result2(0) = "Banana" And result2(1) = "Cherry" Then
                Response.Write "<span class='success'>✓ PASS</span>"
            Else
                Response.Write "<span class='error'>✗ FAIL</span>"
            End If
            %>
        </div>

        <div class="test">
            <h3>Test 3: Filter with Case-Insensitive Compare</h3>
            <%
            Dim arr3, result3
            arr3 = Array("HELLO", "hello", "HeLLo", "goodbye")
            result3 = Filter(arr3, "hello", True, 1)
            Response.Write "Input: Array(""HELLO"", ""hello"", ""HeLLo"", ""goodbye"")<br>"
            Response.Write "Filter(arr3, ""hello"", True, 1) ' compare=1 = case-insensitive<br>"
            Response.Write "UBound: " & UBound(result3) & "<br>"
            Response.Write "Count: " & (UBound(result3) + 1) & "<br>"
            If UBound(result3) = 2 Then
                Response.Write "<span class='success'>✓ PASS: Found 3 matches</span>"
            Else
                Response.Write "<span class='error'>✗ FAIL</span>"
            End If
            %>
        </div>

        <div class="test">
            <h3>Test 4: Direct Indexed Access</h3>
            <%
            Dim arr4, directResult
            arr4 = Array("First", "Second", "Third")
            On Error Resume Next
            directResult = Filter(arr4, "Second")(0)
            If Err.Number <> 0 Then
                Response.Write "<span class='error'>✗ FAIL: Error " & Err.Number & " - " & Err.Description & "</span>"
            Else
                Response.Write "Filter(arr4, ""Second"")(0) = " & directResult & "<br>"
                If directResult = "Second" Then
                    Response.Write "<span class='success'>✓ PASS: Direct indexed access works</span>"
                Else
                    Response.Write "<span class='error'>✗ FAIL: Wrong value</span>"
                End If
            End If
            On Error Goto 0
            %>
        </div>

        <div class="test">
            <h3>Test 5: Empty Result</h3>
            <%
            Dim arr5, result5
            arr5 = Array("One", "Two", "Three")
            result5 = Filter(arr5, "NonExistent")
            Response.Write "Filter(arr5, ""NonExistent"")<br>"
            Response.Write "IsArray: " & IsArray(result5) & "<br>"
            Response.Write "UBound: " & UBound(result5) & "<br>"
            If IsArray(result5) And UBound(result5) = - 1 Then
                Response.Write "<span class='success'>✓ PASS: Empty array returned</span>"
            Else
                Response.Write "<span class='error'>✗ FAIL</span>"
            End If
            %>
        </div>
    </body>
</html>
