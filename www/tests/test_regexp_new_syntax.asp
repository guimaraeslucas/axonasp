<%@ Language="VBScript" %>
<html>
<head>
    <title>RegExp - New Syntax Test</title>
    <style>
        body { font-family: Arial; margin: 20px; }
        .pass { color: green; background: #e8f5e9; padding: 10px; margin: 5px 0; }
        .fail { color: red; background: #ffebee; padding: 10px; margin: 5px 0; }
        code { background: #f0f0f0; padding: 2px 5px; }
    </style>
</head>
<body>

<h1>RegExp - Testing "New RegExp" Syntax</h1>

<%
    On Error Resume Next
    
    Response.Write "<h2>Test 1: Create RegExp with 'New'</h2>"
    Set regex = New RegExp
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        Err.Clear
    else
        Response.Write "<div class='pass'>PASS: 'New RegExp' created object</div>"
    end if
    
    Response.Write "<h2>Test 2: Set Pattern via property assignment</h2>"
    regex.Pattern = "\d+"
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        Err.Clear
    else
        Response.Write "<div class='pass'>PASS: Pattern assigned via property</div>"
    end if
    
    Response.Write "<h2>Test 3: Set IgnoreCase and Global properties</h2>"
    regex.IgnoreCase = True
    regex.Global = True
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        Err.Clear
    else
        Response.Write "<div class='pass'>PASS: Boolean properties assigned</div>"
    end if
    
    Response.Write "<h2>Test 4: Test() method</h2>"
    result = regex.Test("abc123def")
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        Err.Clear
    else
        if result then
            Response.Write "<div class='pass'>PASS: Test() returned TRUE for '123' match</div>"
        else
            Response.Write "<div class='fail'>FAIL: Test() returned FALSE</div>"
        end if
    end if
    
    Response.Write "<h2>Test 5: Execute() method</h2>"
    Set matches = regex.Execute("abc123def456")
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL: " & Err.Description & "</div>"
        Err.Clear
    else
        if matches Is Nothing then
            Response.Write "<div class='fail'>FAIL: Execute() returned Nothing</div>"
        else
            cnt = matches.Count
            Response.Write "<div class='pass'>PASS: Execute() returned matches with Count=" & cnt & "</div>"
        end if
    end if
    
    Response.Write "<h2>Test 6: Access match properties</h2>"
    if Not (matches Is Nothing) And matches.Count > 0 then
        Set m = matches.Item(0)
        Response.Write "First match value: <code>" & m.Value & "</code><br>"
        Response.Write "Index: " & m.FirstIndex & "<br>"
        Response.Write "Length: " & m.Length & "<br>"
        Response.Write "<div class='pass'>PASS: Match properties accessible</div>"
    else
        Response.Write "<div class='fail'>FAIL: No matches to access</div>"
    end if
    
    Response.Write "<h2>Test 7: Replace() method</h2>"
    replaced = regex.Replace("abc123def456", "X")
    
    if Err.Number <> 0 then
        Response.Write "<div class='fail'>FAIL: " & Err.Description & "</div>"
    else
        Response.Write "Input: 'abc123def456'<br>"
        Response.Write "Result: <code>" & replaced & "</code><br>"
        if replaced = "abcXdefX" then
            Response.Write "<div class='pass'>PASS: Replace() worked correctly</div>"
        else
            Response.Write "<div class='fail'>FAIL: Expected 'abcXdefX', got '" & replaced & "'</div>"
        end if
    end if
%>

</body>
</html>
