<%
@ Language = "VBScript"
%>
<!DOCTYPE html>
<html>
    <head>
        <title>UBound Test</title>
        <style>
            body {
                font-family: Tahoma;
                padding: 20px;
            }
            pre {
                background: #f0f0f0;
                padding: 10px;
            }
        </style>
    </head>
    <body>
        <h2>UBound Test</h2>

        <%
        On Error Resume Next

        ' Test 1: 1D array
        Dim arr1d()
        ReDim arr1d(3)
        Response.Write "1D Array Tests:<br>"
        Response.Write "UBound(arr1d) = " & UBound(arr1d) & "<br>"
        Response.Write "UBound(arr1d, 1) = " & UBound(arr1d, 1) & "<br>"

        ' Test 2: 2D array
        Response.Write "<br>2D Array Tests:<br>"
        Dim arr2d()
        ReDim arr2d(2, 3)
        Response.Write "UBound(arr2d) = " & UBound(arr2d) & "<br>"
        Response.Write "UBound(arr2d, 1) = " & UBound(arr2d, 1) & "<br>"
        Response.Write "UBound(arr2d, 2) = " & UBound(arr2d, 2) & "<br>"

        ' Test 3: 3D array
        Response.Write "<br>3D Array Tests:<br>"
        Dim arr3d()
        ReDim arr3d(1, 2, 3)
        Response.Write "UBound(arr3d, 1) = " & UBound(arr3d, 1) & "<br>"
        Response.Write "UBound(arr3d, 2) = " & UBound(arr3d, 2) & "<br>"
        Response.Write "UBound(arr3d, 3) = " & UBound(arr3d, 3) & "<br>"

        If Err.Number <> 0 Then
            Response.Write "<span style='color: red;'><br>Error " & Err.Number & ": " & Err.Description & "</span>"
        End If

        On Error Goto 0
        %>
    </body>
</html>
