<%@ Language="VBScript" %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Arrays & ReDim Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Arrays & ReDim Test</h1>
        <div class="intro">
            <p>Tests dynamic array creation with ReDim, ReDim Preserve, multi-dimensional arrays and VarType for arrays.</p>
        </div>
        <div class="box">

<%
' Test Array Redim and Preserve
Dim arr()
Dim i

Response.Write "<h3>Testing ReDim</h3>"

ReDim arr(2)
arr(0) = "A"
arr(1) = "B"
arr(2) = "C"

Response.Write "Size 2 (0-2): " & UBound(arr) & "<br>"
Response.Write "Item 0: " & arr(0) & "<br>"
Response.Write "Item 2: " & arr(2) & "<br>"

Response.Write "<h3>Testing ReDim Preserve</h3>"

ReDim Preserve arr(4)
arr(3) = "D"
arr(4) = "E"

Response.Write "Size 4 (0-4): " & UBound(arr) & "<br>"
Response.Write "Item 0 (Preserved): " & arr(0) & "<br>"
Response.Write "Item 2 (Preserved): " & arr(2) & "<br>"
Response.Write "Item 4 (New): " & arr(4) & "<br>"

Response.Write "<h3>Testing Dynamic Access</h3>"
Dim dyn
dyn = arr
Response.Write "From Dynamic Var (2): " & dyn(2) & "<br>"

%>

<%
Response.Write "<h3>Testing Multi-Dimensional ReDim Preserve</h3>"
Dim m
ReDim m(1,1)
m(0,0) = "X"
m(0,1) = "Y"
Response.Write "Before Preserve UBound dim1: " & UBound(m,1) & " UBound dim2: " & UBound(m,2) & "<br>"

ReDim Preserve m(2,2)
Response.Write "After Preserve UBound dim1: " & UBound(m,1) & " UBound dim2: " & UBound(m,2) & "<br>"
Response.Write "Preserved [0,0]: " & m(0,0) & "<br>"

Response.Write "<h3>VarType & vbArray test</h3>"
Response.Write "VarType(arr): " & VarType(arr) & " (expect 8204 for array)" & "<br>"
%>
        </div>
    </div>
</body>
</html>
