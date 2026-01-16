<%
Response.Write "<h3>Testing VBScript.RegExp</h3>"

Dim re, match, matches
Set re = Server.CreateObject("VBScript.RegExp")

re.Pattern = "\b\w+\b"
re.Global = True
re.IgnoreCase = True

Dim inputStr
inputStr = "Hello World This is ASP"

Response.Write "Pattern: " & re.Pattern & "<br>"
Response.Write "Input: " & inputStr & "<br>"

' Test Method
Dim isMatch
isMatch = re.Test(inputStr)
Response.Write "Test result: " & isMatch & "<br>"

' Execute Method
Set matches = re.Execute(inputStr)
Response.Write "Match Count: " & matches.Count & "<br>"

Dim i
For i = 0 To matches.Count - 1
    Set match = matches.Item(i)
    Response.Write "Match " & i & ": " & match.Value & " (Index: " & match.FirstIndex & ", Length: " & match.Length & ")<br>"
Next

' Replace Method
re.Pattern = "ASP"
Dim replaced
replaced = re.Replace(inputStr, "GoLang")
Response.Write "Replace result: " & replaced & "<br>"

' Non-Global Replace
re.Global = False
re.Pattern = "\s"
replaced = re.Replace(inputStr, "-")
Response.Write "Non-Global Replace (first space): " & replaced & "<br>"

%>