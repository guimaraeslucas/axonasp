<%@ Language="VBScript" %>
<html>
<head><title>Debug Tests 1, 6, 7</title>
<style>
body { font-family: Arial; margin: 20px; }
.test { padding: 10px; margin: 10px 0; border: 1px solid #ccc; }
.pass { background: #e8f5e9; color: green; }
.fail { background: #ffebee; color: red; }
.info { background: #e0f7ff; color: blue; }
code { background: #f0f0f0; padding: 2px 5px; }
</style>
</head>
<body>

<h1>Debug: Tests 1, 6, 7</h1>

<%
    On Error Resume Next
    
    ' TEST 1 DEBUG
    Response.Write "<h2>TEST 1: New RegExp</h2>"
    Response.Write "<div class='test info'>"
    
    Err.Clear
    Set regex = New RegExp
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL: Error " & Err.Number & " - " & Err.Description & "</div>"
    else
        Response.Write "<div class='pass'>PASS: Object created</div>"
        Response.Write "Object type: " & TypeName(regex) & "<br>"
    end if
    
    Response.Write "</div>"
    
    ' TEST 6 DEBUG
    Response.Write "<h2>TEST 6: Execute() with Global=True</h2>"
    Response.Write "<div class='test info'>"
    
    Err.Clear
    Set regex = New RegExp
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL creating object: " & Err.Description & "</div>"
    else
        Response.Write "Object created<br>"
        
        Err.Clear
        regex.Pattern = "\d+"
        if Err.Number <> 0 then
            Response.Write "<div class='fail'>FAIL setting Pattern: " & Err.Description & "</div>"
        else
            Response.Write "Pattern set<br>"
        end if
        
        Err.Clear
        regex.Global = True
        if Err.Number <> 0 then
            Response.Write "<div class='fail'>FAIL setting Global: " & Err.Description & "</div>"
        else
            Response.Write "Global set<br>"
        end if
        
        Err.Clear
        testStr = "The code is 2024, user id 12345, pin 999"
        Set matches = regex.Execute(testStr)
        
        if Err.Number <> 0 then
            Response.Write "<div class='fail'>FAIL on Execute: " & Err.Description & "</div>"
        else
            Response.Write "Execute completed<br>"
            Response.Write "Matches is Nothing: " & (matches Is Nothing) & "<br>"
            
            if Not (matches Is Nothing) then
                Err.Clear
                cnt = matches.Count
                if Err.Number <> 0 then
                    Response.Write "<div class='fail'>FAIL accessing Count: " & Err.Description & "</div>"
                else
                    Response.Write "<div class='pass'>PASS: Count = " & cnt & "</div>"
                end if
            else
                Response.Write "<div class='fail'>FAIL: matches is Nothing</div>"
            end if
        end if
    end if
    
    Response.Write "</div>"
    
    ' TEST 7 DEBUG
    Response.Write "<h2>TEST 7: Execute() with Global=False</h2>"
    Response.Write "<div class='test info'>"
    
    Err.Clear
    Set regex = New RegExp
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL creating object: " & Err.Description & "</div>"
    else
        Response.Write "Object created<br>"
        
        Err.Clear
        regex.Pattern = "[a-z]+"
        regex.IgnoreCase = True
        regex.Global = False
        
        if Err.Number <> 0 then
            Response.Write "<div class='fail'>FAIL setting properties: " & Err.Description & "</div>"
        else
            Response.Write "Properties set<br>"
        end if
        
        Err.Clear
        testStr = "Hello World JavaScript 2024"
        Set matches = regex.Execute(testStr)
        
        if Err.Number <> 0 then
            Response.Write "<div class='fail'>FAIL on Execute: " & Err.Description & "</div>"
        else
            Response.Write "Execute completed<br>"
            Response.Write "Matches is Nothing: " & (matches Is Nothing) & "<br>"
            
            if Not (matches Is Nothing) then
                Err.Clear
                cnt = matches.Count
                if Err.Number <> 0 then
                    Response.Write "<div class='fail'>FAIL accessing Count: " & Err.Description & "</div>"
                else
                    Response.Write "<div class='pass'>PASS: Count = " & cnt & "</div>"
                    
                    if cnt > 0 then
                        Err.Clear
                        Set m = matches.Item(0)
                        if Err.Number <> 0 then
                            Response.Write "<div class='fail'>FAIL accessing Item(0): " & Err.Description & "</div>"
                        else
                            Response.Write "<div class='pass'>PASS: Item(0) = " & m.Value & "</div>"
                        end if
                    end if
                end if
            else
                Response.Write "<div class='fail'>FAIL: matches is Nothing</div>"
            end if
        end if
    end if
    
    Response.Write "</div>"
%>

</body>
</html>
