<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head><title>Fix Verification</title></head>
<body>
    <h1>Fix Verification</h1>
    <p>Now(): [<%= Now() %>]</p>
    <p>Now: [<%= Now %>]</p>
    <p>RGB(255, 0, 0): [<%= RGB(255, 0, 0) %>] (Expected: 255)</p>
    <p>IsObject(Now): [<%= IsObject(Now) %>] (Expected: False)</p>
    <p>TypeName(123): [<%= TypeName(123) %>] (Expected: Integer)</p>
    <p>ScriptEngine: [<%= ScriptEngine %>] (Expected: VBScript)</p>
</body>
</html>
