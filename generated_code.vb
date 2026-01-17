Response.Write("﻿" & vbCrLf)
' <%@ Language=VBScript %>
Response.Write("" & vbCrLf)
Response.Write("<!DOCTYPE html>" & vbCrLf)
Response.Write("<html lang=""en"">" & vbCrLf)
Response.Write("<head>" & vbCrLf)
Response.Write("    <meta charset=""UTF-8"">" & vbCrLf)
Response.Write("    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf)
Response.Write("    <title>G3pix AxonASP - Features & Operators Test</title>" & vbCrLf)
Response.Write("    <style>" & vbCrLf)
Response.Write("        * { margin: 0; padding: 0; box-sizing: border-box; }" & vbCrLf)
Response.Write("        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }" & vbCrLf)
Response.Write("        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }" & vbCrLf)
Response.Write("        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }" & vbCrLf)
Response.Write("        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }" & vbCrLf)
Response.Write("        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }" & vbCrLf)
Response.Write("        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }" & vbCrLf)
Response.Write("        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }" & vbCrLf)
Response.Write("        pre { background: #f4f4f4; padding: 10px; border-radius: 4px; overflow-x: auto; }" & vbCrLf)
Response.Write("        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }" & vbCrLf)
Response.Write("        form { margin: 15px 0; }" & vbCrLf)
Response.Write("        input, button { padding: 8px 12px; margin: 5px 5px 5px 0; border: 1px solid #ddd; border-radius: 4px; }" & vbCrLf)
Response.Write("        button { background: #667eea; color: #fff; cursor: pointer; }" & vbCrLf)
Response.Write("        button:hover { background: #764ba2; }" & vbCrLf)
Response.Write("    </style>" & vbCrLf)
Response.Write("</head>" & vbCrLf)
Response.Write("<body>" & vbCrLf)
Response.Write("    <div class=""container"">" & vbCrLf)
Response.Write("        <h1>G3pix AxonASP - Features & Operators Test</h1>" & vbCrLf)
Response.Write("        <div class=""intro"">" & vbCrLf)
Response.Write("            <p>Tests mathematical operators, subroutines with parameters, Request object and form operations.</p>" & vbCrLf)
Response.Write("        </div>" & vbCrLf)
Response.Write("" & vbCrLf)
Response.Write("    <h2>1. Math Operators</h2>" & vbCrLf)
Response.Write("    " & vbCrLf)

        Dim a, b
        a = 10
        b = 5
        Response.Write "a=" & a & ", b=" & b & "<br>"
        Response.Write "a * b = " & (a * b) & "<br>"
        Response.Write "a - b = " & (a - b) & "<br>"
        Response.Write "100 / 2 * 5 = " & (100 / 2 * 5) & " (Should be 250)<br>"
        Response.Write "10 - 5 - 2 = " & (10 - 5 - 2) & " (Should be 3)<br>"
    
Response.Write("" & vbCrLf)
Response.Write("" & vbCrLf)
Response.Write("    <h2>2. Subroutines with Params</h2>" & vbCrLf)
Response.Write("    " & vbCrLf)

        Sub Calc(x, y)
            Response.Write "Inside Calc: x=" & x & ", y=" & y & "<br>"
            Response.Write "Product: " & (x * y) & "<br>"
        End Sub

        Response.Write "Calling Calc(6, 7)...<br>"
        Call Calc(6, 7)
    
Response.Write("" & vbCrLf)
Response.Write("" & vbCrLf)
Response.Write("    <h2>3. Request Object</h2>" & vbCrLf)
Response.Write("    <p>Try appending ?name=Lucas to the URL.</p>" & vbCrLf)
Response.Write("    " & vbCrLf)

        Dim n
        n = Request.QueryString("name")
        If n = "" Then
            Response.Write "No name provided in QueryString.<br>"
        Else
            Response.Write "Hello, " & n & "! (from QueryString)<br>"
        End If
    
Response.Write("" & vbCrLf)
Response.Write("" & vbCrLf)
Response.Write("    <h3>Form Post</h3>" & vbCrLf)
Response.Write("    <form method=""POST"">" & vbCrLf)
Response.Write("        <input type=""text"" name=""city"" value=""Sao Paulo"">" & vbCrLf)
Response.Write("        <input type=""submit"" value=""Send"">" & vbCrLf)
Response.Write("    </form>" & vbCrLf)
Response.Write("    " & vbCrLf)

        Dim c
        c = Request.Form("city")
        If c <> "" Then
            Response.Write "You posted city: " & c & "<br>"
        End If
    
Response.Write("" & vbCrLf)
Response.Write("    </div>" & vbCrLf)
Response.Write("" & vbCrLf)
Response.Write("    </div>" & vbCrLf)
Response.Write("</body>" & vbCrLf)
Response.Write("</html>" & vbCrLf)
Response.Write("" & vbCrLf)
Response.Write("" & vbCrLf)
