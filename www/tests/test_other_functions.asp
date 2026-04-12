<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - Other Functions Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h2 { color: #555; margin-top: 30px; margin-bottom: 15px; }
        h3 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
        th { background: #f5f5f5; color: #333; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - Other VBScript Functions Test</h1>
        <div class="intro">
            <p>Tests ScriptEngine, TypeName, VarType, RGB, IsObject, CreateObject and Eval.</p>
        </div>

<% Response.Write("<h2>ScriptEngine Functions</h2>")
Response.Write("ScriptEngine: " & ScriptEngine() & "<br>")
Response.Write("ScriptEngineBuildVersion: " & ScriptEngineBuildVersion() & "<br>")
Response.Write("ScriptEngineMajorVersion: " & ScriptEngineMajorVersion() & "<br>")
Response.Write("ScriptEngineMinorVersion: " & ScriptEngineMinorVersion() & "<br>")

' TypeName and VarType
Response.Write("<h2>TypeName and VarType Functions</h2>")
Dim testVar
testVar = "Hello"
Response.Write("TypeName('Hello'): " & TypeName(testVar) & "<br>")
Response.Write("VarType('Hello'): " & VarType(testVar) & "<br>")
testVar = 42
Response.Write("TypeName(42): " & TypeName(testVar) & "<br>")
Response.Write("VarType(42): " & VarType(testVar) & "<br>")
testVar = True
Response.Write("TypeName(True): " & TypeName(testVar) & "<br>")
Response.Write("VarType(True): " & VarType(testVar) & "<br>")
testVar = Null
Response.Write("TypeName(Null): " & TypeName(testVar) & "<br>")
Response.Write("VarType(Null): " & VarType(testVar) & "<br>")

' RGB function
Response.Write("<h2>RGB Function</h2>")
Response.Write("RGB(255, 0, 0): " & RGB(255, 0, 0) & " (Red)<br>")
Response.Write("RGB(0, 255, 0): " & RGB(0, 255, 0) & " (Green)<br>")
Response.Write("RGB(0, 0, 255): " & RGB(0, 0, 255) & " (Blue)<br>")
Response.Write("RGB(128, 128, 128): " & RGB(128, 128, 128) & " (Gray)<br>")

' IsObject function
Response.Write("<h2>IsObject Function</h2>")
Response.Write("IsObject(Nothing): " & IsObject(Nothing) & "<br>")
Response.Write("IsObject('string'): " & IsObject("string") & "<br>")
Response.Write("IsObject(42): " & IsObject(42) & "<br>")

' CreateObject (stub implementation)
Response.Write("<h2>CreateObject Function (Stub)</h2>")
Dim regexObj
regexObj = CreateObject("VBScript.RegExp")
Response.Write("CreateObject('VBScript.RegExp'): " & TypeName(regexObj) & "<br>")

' Eval function
Response.Write("<h2>Eval Function</h2>")
Response.Write("Eval('2 + 3'): " & Eval("2 + 3") & "<br>")
Response.Write("Eval('10 * 5'): " & Eval("10 * 5") & "<br>")

Response.Write("<h3 style='margin-top:20px'>Testes de Outras Funções Completados!</h3>")
%>
    </div>
</body>
</html>

