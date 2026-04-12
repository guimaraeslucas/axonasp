<%
@ Language = "VBScript"
%>
<!DOCTYPE html>
<html>
    <head>
        <title>Single 2D ReDim Test</title>
        <style>
            body {
                font-family: Tahoma;
                padding: 20px;
            }
            .test {
                border: 1px solid #ccc;
                padding: 15px;
                margin: 10px 0;
            }
            .pass {
                color: green;
                font-weight: bold;
            }
            .fail {
                color: red;
                font-weight: bold;
            }
        </style>
    </head>
    <body>
        <h1>Single 2D ReDim Preserve Test</h1>

        <%
        On Error Resume Next

        Dim arr2d()
        ReDim arr2d(2, 3)

        ' Fill with initial values
        arr2d(0, 0) = "A00"
        arr2d(0, 1) = "A01"
        arr2d(1, 0) = "B00"
        arr2d(1, 1) = "B01"
        arr2d(2, 0) = "C00"
        arr2d(2, 2) = "C02"

        ' Expand last dimension from 4 to 6
        ReDim Preserve arr2d(2, 5)

        ' Check for errors
        If Err.Number <> 0 Then
            Response.Write "<span class='fail'>✗ FAIL: Error " & Err.Number & " - " & Err.Description & "</span><br>"
        Else
            ' Verify data preservation
            Dim test1Pass
            test1Pass = (arr2d(0, 0) = "A00" And arr2d(0, 1) = "A01" And _
            arr2d(1, 0) = "B00" And arr2d(1, 1) = "B01" And _
            arr2d(2, 0) = "C00" And arr2d(2, 2) = "C02" And _
            UBound(arr2d, 1) = 2 And UBound(arr2d, 2) = 5)

            If test1Pass Then
                Response.Write "<span class='pass'>✓ PASS: 2D array expanded and data preserved</span><br>"
            Else
                Response.Write "<span class='fail'>✗ FAIL: Data mismatch</span><br>"
                Response.Write "arr2d(0,0)=" & arr2d(0, 0) & "<br>"
                Response.Write "arr2d(2,2)=" & arr2d(2, 2) & "<br>"
                Response.Write "UBound(arr2d,1)=" & UBound(arr2d,  1) & "<br>"
                Response.Write "UBound(arr2d,2)=" & UBound(arr2d, 2) & "<br>"
            End If
        End If

        On Error Goto 0
        %>
    </body>
</html>
