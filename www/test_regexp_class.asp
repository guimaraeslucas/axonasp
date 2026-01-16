<%
Class HTMLStripper
    Public Function StripHTML(html)
        Dim objRegExp, result
        Set objRegExp = New RegExp
        
        objRegExp.Pattern = "<[^>]*>"
        objRegExp.Global = True
        objRegExp.IgnoreCase = True
        
        result = objRegExp.Replace(html, "")
        StripHTML = result
    End Function
End Class

Dim stripper
Set stripper = New HTMLStripper
Dim html
html = "<p>Hello <strong>World</strong>!</p>"

Dim cleaned
cleaned = stripper.StripHTML(html)

Response.Write "Original: " & html & "<br>"
Response.Write "Cleaned: " & cleaned & "<br>"
%>
