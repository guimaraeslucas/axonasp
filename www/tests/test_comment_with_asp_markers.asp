<%
@ Language = VBScript
%>
<!DOCTYPE html>
<html>
    <head>
        <title>Test: Comment with ASP Markers</title>
    </head>
    <body>
        <h1>Test: VBScript Comment with ASP Markers</h1>

        <h2>Test 1: Comment with %> in the middle</h2>
        <%
        Dim test1
        test1 = "before"
        ' This is a comment with %> in the middle and more text after
        test1 = test1 & " after"
        Response.Write "<p>Test 1 result: " & test1 & "</p>"
        %>

        <h2>Test 2: Comment with multiple %> sequences</h2>
        <%
        Dim test2
        test2 = "start"
        ' Comment with multiple markers: %> text <%=var%> more text %>
        test2 = test2 & " end"
        Response.Write "<p>Test 2 result: " & test2 & "</p>"
        %>

        <h2>Test 3: Comment ending with %> (valid closure)</h2>
        <%
        Dim test3
        test3 = "valid"
        Response.Write "<p>Test 3 result: " & test3 & "</p>" ' comment ending With
        %>

        <h2>Test 4: Complex case from error report</h2>
        <%
        Dim iId, uploadPost
        iId = 12345
        Set uploadPost = Server.CreateObject("Scripting.Dictionary")
        uploadPost.Add "iId", iId

        If 1 = 1 Then 'reply
        %>parent.document.getElementById('uSS<%= encrypt(uploadPost.iId) %>').innerHTML='test';
        <%
            Response.Write "<p>This should work correctly</p>"
        End If
        %>

        <h2>Test 5: HTML with embedded ASP expressions</h2>
        <p>
            Parent: parent.document.form<%= test3 %>.field.value
            = 'test';
        </p>

        <p><strong>All tests completed successfully!</strong></p>
    </body>
</html>
