<%@ Language="VBScript" CodePage=65001 %>
<%@ Language="VBScript" CodePage=65001 %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - ByRef Parameter Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - ByRef Parameter Test</h1>
        <div class="intro">
            <p>Tests function parameters passed by reference using ByRef keyword.</p>
        </div>
        <% 
    ' Test with byref parameter
    private function StreamTest(path, byref size)
        size = 42
        StreamTest = "Content"
    end function

    Dim mySize
    mySize = 0
    Dim content
    content = StreamTest("file.txt", mySize)
    
    Response.Write("<p>Result: " & content & "</p>")
    Response.Write("<p>Size: " & mySize & "</p>")
%>
    </div>
</body>
</html>
