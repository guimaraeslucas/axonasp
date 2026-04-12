<%
Response.Write "<h3>Testing New RegExp Syntax in Class</h3>"

Class HTMLStripper
    Public Function StripHTML(strHTML)
        Dim objRegExp, strOutput
        Set objRegExp = New RegExp

        If IsObject(objRegExp) Then
            Response.Write "DEBUG: objRegExp IS Object<br>"
        Else
            Response.Write "DEBUG: objRegExp IS NOT Object (IsNull? " & IsNull(objRegExp) & " IsEmpty? " & IsEmpty(objRegExp) & ")<br>"
        End If

        objRegExp.IgnoreCase = True
        objRegExp.Global = True
        objRegExp.Pattern = "<(.|\n)+?>"

        strOutput = objRegExp.Replace(strHTML, "")
        StripHTML = strOutput
        
        Set objRegExp = Nothing
    End Function
End Class

Dim stripper, html
Set stripper = New HTMLStripper
html = "<p>Hello <b>World</b></p>"
Response.Write "Original: " & Server.HTMLEncode(html) & "<br>"
Response.Write "Stripped: " & stripper.StripHTML(html) & "<br>"
%>