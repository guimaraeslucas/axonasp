<%@ LANGUAGE="VBSCRIPT" %>
<%
    Response.Write("<h2>Debug Form Data</h2>")
    
    Response.Write("<p>Method: " & Request.ServerVariables("REQUEST_METHOD") & "</p>")
    
    Response.Write("<h3>Form Collection:</h3>")
    Response.Write("<ul>")
    
    Dim key
    For Each key In Request.Form
        Response.Write("<li>" & key & " = " & Request.Form(key) & "</li>")
    Next
    
    Response.Write("</ul>")
    
    Response.Write("<h3>Direct Access Test:</h3>")
    Response.Write("<p>submitted = '" & Request.Form("submitted") & "'</p>")
    Response.Write("<p>username = '" & Request.Form("username") & "'</p>")
    
    Response.Write("<h3>Empty Check:</h3>")
    If Request.Form("submitted") = "" Then
        Response.Write("<p>submitted IS EMPTY</p>")
    Else
        Response.Write("<p>submitted IS NOT EMPTY: " & Request.Form("submitted") & "</p>")
    End If
%>
