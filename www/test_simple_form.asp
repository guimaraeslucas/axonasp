<%@ LANGUAGE="VBSCRIPT" %>
<%
    Response.Write("<h2>Simple Form Test</h2>")
    
    Dim val
    val = Request.Form("submitted")
    
    Response.Write("<p>Value: '" & val & "'</p>")
    Response.Write("<p>Length: " & Len(val) & "</p>")
    Response.Write("<p>Type: " & TypeName(val) & "</p>")
    
    Response.Write("<h3>Comparison tests:</h3>")
    Response.Write("<p>val = '1': " & (val = "1") & "</p>")
    Response.Write("<p>val = 1: " & (val = 1) & "</p>")
    Response.Write("<p>val <> '': " & (val <> "") & "</p>")
    Response.Write("<p>Len(val) > 0: " & (Len(val) > 0) & "</p>")
    
    If Len(val) > 0 Then
        Response.Write("<h3>SUCCESS! Form data found!</h3>")
        Response.Write("<ul>")
        Dim k
        For Each k In Request.Form
            Response.Write("<li>" & k & " = " & Request.Form(k) & "</li>")
        Next
        Response.Write("</ul>")
    End If
%>
