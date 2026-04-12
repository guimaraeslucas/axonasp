<%
@ Language = "VBScript"
%>
<!DOCTYPE html>
<html>
    <head>
        <title>Simple ReDim Preserve Test</title>
        <style>
            body {
                font-family: Tahoma;
                padding: 20px;
                background: #ece9d8;
            }
            .container {
                max-width: 600px;
                margin: 0 auto;
                background: #fff;
                padding: 20px;
                border: 1px solid #808080;
            }
            h1 {
                color: #003399;
                border-bottom: 2px solid #003399;
            }
            .pass {
                color: #388e3c;
                font-weight: bold;
            }
            .fail {
                color: #d32f2f;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Simple ReDim Preserve Test</h1>

            <h2>Test 1: Create 1D array and resize</h2>
            <%
            On Error Resume Next

            Dim arr1d()
            ReDim arr1d(4)
            arr1d(0) = "A"
            arr1d(1) = "B"

            Response.Write "Before ReDim: UBound = " & UBound(arr1d) & "<br>"
            Response.Write "arr1d(0) = " & arr1d(0) & ", arr1d(1) = " & arr1d(1) & "<br>"

            ReDim Preserve arr1d(9)

            If Err.Number <> 0 Then
                Response.Write "<span class='fail'>✗ FAIL: Error " & Err.Number & " - " & Err.Description & "</span><br>"
            Else
                Response.Write "After ReDim: UBound = " & UBound(arr1d) & "<br>"
                Response.Write "arr1d(0) = " & arr1d(0) & ", arr1d(1) = " & arr1d(1) & "<br>"
                If arr1d(0) = "A" And arr1d(1) = "B" And UBound(arr1d) = 9 Then
                    Response.Write "<span class='pass'>✓ PASS: 1D ReDim Preserve works</span><br>"
                Else
                    Response.Write "<span class='fail'>✗ FAIL: Data not preserved correctly</span><br>"
                End If
            End If

            On Error Goto 0
            %>

            <h2>Test 2: Create 2D array and resize last dim</h2>
            <%
            On Error Resume Next

            Dim arr2d()
            ReDim arr2d(2, 3)
            arr2d(0, 0) = "X"
            arr2d(1, 1) = "Y"

            Response.Write "Before ReDim: Bounds = (" & UBound(arr2d, 1) & ", " & UBound(arr2d, 2) & ")<br>"
            Response.Write "arr2d(0,0) = " & arr2d(0, 0) & ", arr2d(1,1) = " & arr2d(1, 1) & "<br>"

            ReDim Preserve arr2d(2, 5)

            If Err.Number <> 0 Then
                Response.Write "<span class='fail'>✗ FAIL: Error " & Err.Number & " - " & Err.Description & "</span><br>"
            Else
                Response.Write "After ReDim: Bounds = (" & UBound(arr2d, 1) & ", " & UBound(arr2d, 2) & ")<br>"
                Response.Write "arr2d(0,0) = " & arr2d(0, 0) & ", arr2d(1,1) = " & arr2d(1, 1) & "<br>"
                If arr2d(0, 0) = "X" And arr2d(1, 1) = "Y" And UBound(arr2d, 1) = 2 And UBound(arr2d, 2) = 5 Then
                    Response.Write "<span class='pass'>✓ PASS: 2D ReDim Preserve works</span><br>"
                Else
                    Response.Write "<span class='fail'>✗ FAIL: Data not preserved correctly</span><br>"
                End If
            End If

            On Error Goto 0
            %>

            <h2>Test 3: Try to change first dimension (should error)</h2>
            <%
            On Error Resume Next

            Dim arr2d_test()
            ReDim arr2d_test(2, 3)
            arr2d_test(0, 0) = "Z"

            ReDim Preserve arr2d_test(3, 3)

            If Err.Number = 5 Then
                Response.Write "<span class='pass'>✓ PASS: Correctly rejected first-dimension change (Error 5)</span><br>"
            Else
                Response.Write "<span class='fail'>✗ FAIL: Should have rejected, got Error " & Err.Number & "</span><br>"
            End If

            On Error Goto 0
            %>
        </div>
    </body>
</html>
