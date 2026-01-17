<%@ Language=VBScript %>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>G3pix AxonASP - ASPLite Fixes Test</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: Tahoma, 'Segoe UI', Geneva, Verdana, sans-serif; padding: 30px; background: #f5f5f5; line-height: 1.6; }
        .container { max-width: 1000px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        h1 { color: #333; margin-bottom: 10px; border-bottom: 2px solid #667eea; padding-bottom: 10px; }
        h3, h4 { color: #666; margin-top: 15px; margin-bottom: 10px; }
        .intro { background: #e3f2fd; border-left: 4px solid #2196f3; padding: 15px; margin-bottom: 20px; border-radius: 4px; }
        .box { border-left: 4px solid #667eea; padding: 15px; margin-bottom: 15px; background: #f9f9f9; border-radius: 4px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>G3pix AxonASP - ASPLite Fixes Test</h1>
        <div class="intro">
            <p>Tests critical fixes for compatibility: AscW/ChrW, Dictionary enumeration, VarType and default methods in classes.</p>
        </div>
        <div class="box">

<%
Option Explicit

Response.Write "<h3>Testing ASPLite Fixes</h3>"

' 1. Test AscW and ChrW
Dim s, code
s = "A"
code = AscW(s)
Response.Write "AscW('A') = " & code & "<br>"
If code = 65 Then 
    Response.Write "<span style='color:green'>AscW OK</span><br>" 
Else 
    Response.Write "<span style='color:red'>AscW FAIL</span><br>"
End If

Dim c
c = ChrW(65)
Response.Write "ChrW(65) = " & c & "<br>"
If c = "A" Then 
    Response.Write "<span style='color:green'>ChrW OK</span><br>" 
Else 
    Response.Write "<span style='color:red'>ChrW FAIL</span><br>"
End If

' 2. Test Dictionary Enumeration (Keys)
Response.Write "<hr>"
Dim d, k, keysStr
Set d = Server.CreateObject("Scripting.Dictionary")
d.Add "Key1", "Value1"
d.Add "Key2", "Value2"

keysStr = ""
For Each k In d
    keysStr = keysStr & k & ";"
Next
Response.Write "Dictionary Keys: " & keysStr & "<br>"

If InStr(keysStr, "Key1") > 0 And InStr(keysStr, "Key2") > 0 Then
    Response.Write "<span style='color:green'>Dictionary Enumeration OK</span><br>"
Else
    Response.Write "<span style='color:red'>Dictionary Enumeration FAIL (Expected Key1;Key2;, got " & keysStr & ")</span><br>"
End If

' 3. Test VarType
Response.Write "<hr>"
Dim arr(2)
Dim obj
Set obj = Server.CreateObject("Scripting.Dictionary")

Dim vtArr, vtObj
vtArr = VarType(arr)
vtObj = VarType(obj)

Response.Write "VarType(Array) = " & vtArr & "<br>"
Response.Write "VarType(Object) = " & vtObj & "<br>"

If vtArr = 8204 Then 
    Response.Write "<span style='color:green'>VarType Array OK</span><br>" 
Else 
    Response.Write "<span style='color:red'>VarType Array FAIL (Expected 8204)</span><br>"
End If
If vtObj = 9 Then 
    Response.Write "<span style='color:green'>VarType Object OK</span><br>" 
Else 
    Response.Write "<span style='color:red'>VarType Object FAIL (Expected 9)</span><br>"
End If

' 4. Test Dim initialization (Empty/Nil)
Response.Write "<hr>"
Dim emptyVar
Response.Write "TypeName(emptyVar) = " & TypeName(emptyVar) & "<br>"
If IsEmpty(emptyVar) Then
    Response.Write "<span style='color:green'>Dim Initialization OK (IsEmpty is True)</span><br>"
Else
    Response.Write "<span style='color:red'>Dim Initialization FAIL (IsEmpty is False, TypeName=" & TypeName(emptyVar) & ")</span><br>"
End If

' 5. Test Default Method in Class
Response.Write "<hr>"
Class TestClass
    Public Default Function MyDefault(x)
        MyDefault = "Hello " & x
    End Function
End Class

Dim t
Set t = New TestClass
Dim res
res = t("World") ' Calling object as function
Response.Write "Default Method Result: " & res & "<br>"

If res = "Hello World" Then
    Response.Write "<span style='color:green'>Default Method OK</span><br>"
Else
    Response.Write "<span style='color:red'>Default Method FAIL (Expected 'Hello World', got '" & res & "')</span><br>"
End If
%>
        </div>
    </div>
</body>
</html>