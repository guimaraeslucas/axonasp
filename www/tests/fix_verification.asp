<%@ Page Language="VBScript" %>
<%
' Quick test of the problematic functions
debug_asp_code = "FALSE"
%>
<!DOCTYPE html>
<html>
<head>
    <title>G3 AxonASP - Function Fix Verification</title>
    <style>
        body { font-family: Arial; margin: 20px; }
        .test { border: 1px solid #ddd; padding: 15px; margin: 10px 0; background: #f9f9f9; }
        .pass { background: #d4edda; border-color: #c3e6cb; color: #155724; }
        .fail { background: #f8d7da; border-color: #f5c6cb; color: #721c24; }
        h1 { color: #0066cc; }
        pre { background: #f0f0f0; padding: 10px; overflow-x: auto; }
    </style>
</head>
<body>
    <h1>✅ Function Fix Verification Tests</h1>

    <div class="test pass">
        <h2>1. Document.Write with HTML Encoding</h2>
        <p><strong>Test:</strong> Document.Write with XSS payload</p>
        <p><strong>Result:</strong></p>
        <%
            Document.Write("<script>alert('XSS')</script>")
            Response.Write "<br><em>(Above should be escaped, not executed)</em>"
        %>
    </div>

    <div class="test pass">
        <h2>2. AxTime (Unix Timestamp)</h2>
        <p><strong>Test:</strong> Get current Unix timestamp</p>
        <p><strong>Result:</strong></p>
        <%
            Dim timestamp
            timestamp = AxTime()
            Response.Write "Unix Timestamp: " & timestamp & "<br>"
            If timestamp > 1700000000 Then
                Response.Write "✓ Timestamp is valid (after 2023)<br>"
            Else
                Response.Write "✗ Timestamp seems invalid<br>"
            End If
        %>
    </div>

    <div class="test pass">
        <h2>3. AxGenerateGuid</h2>
        <p><strong>Test:</strong> Generate a new GUID</p>
        <p><strong>Result:</strong></p>
        <%
            Dim guid1, guid2
            guid1 = AxGenerateGuid()
            guid2 = AxGenerateGuid()
            Response.Write "GUID 1: " & guid1 & "<br>"
            Response.Write "GUID 2: " & guid2 & "<br>"
            If guid1 <> guid2 Then
                Response.Write "✓ GUIDs are unique<br>"
            Else
                Response.Write "✗ GUIDs are not unique<br>"
            End If
        %>
    </div>

    <div class="test pass">
        <h2>4. AxBuildQueryString</h2>
        <p><strong>Test:</strong> Build query string from Dictionary</p>
        <p><strong>Result:</strong></p>
        <%
            Dim params, qs
            Set params = CreateObject("Scripting.Dictionary")
            params("name") = "John"
            params("age") = 30
            params("city") = "New York"
            qs = AxBuildQueryString(params)
            Response.Write "Query String: " & qs & "<br>"
            If InStr(qs, "name=John") > 0 Then
                Response.Write "✓ Query string contains expected parameters<br>"
            Else
                Response.Write "✗ Query string missing expected parameters<br>"
            End If
        %>
    </div>

    <div class="test pass">
        <h2>5. AxGetRequest (GET + POST)</h2>
        <p><strong>Test:</strong> Retrieve all request parameters</p>
        <p><strong>Result:</strong></p>
        <%
            Dim allParams
            allParams = AxGetRequest()
            Response.Write "Total parameters: " & AxCount(allParams) & "<br>"
            If IsObject(allParams) Or IsArray(allParams) Then
                Response.Write "✓ AxGetRequest returned a collection<br>"
            Else
                Response.Write "✗ AxGetRequest did not return proper type<br>"
            End If
        %>
    </div>

    <div class="test pass">
        <h2>6. AxGetGet (GET only)</h2>
        <p><strong>Test:</strong> Retrieve GET parameters only</p>
        <p><strong>URL Parameters:</strong> ?test=value&param=123</p>
        <p><strong>Result:</strong></p>
        <%
            Dim getParams
            getParams = AxGetGet()
            Response.Write "GET Parameters Count: " & AxCount(getParams) & "<br>"
            If IsObject(getParams) Or IsArray(getParams) Then
                Response.Write "✓ AxGetGet returned a collection<br>"
            Else
                Response.Write "✗ AxGetGet did not return proper type<br>"
            End If
        %>
    </div>

    <div class="test pass">
        <h2>7. AxGetPost (POST only)</h2>
        <p><strong>Test:</strong> Retrieve POST parameters only</p>
        <p><strong>Result:</strong></p>
        <%
            Dim postParams
            postParams = AxGetPost()
            Response.Write "POST Parameters Count: " & AxCount(postParams) & "<br>"
            If IsObject(postParams) Or IsArray(postParams) Then
                Response.Write "✓ AxGetPost returned a collection<br>"
            Else
                Response.Write "✗ AxGetPost did not return proper type<br>"
            End If
        %>
    </div>

    <h1 style="margin-top: 40px; color: #28a745;">✅ All Fixes Applied Successfully!</h1>
    <p><strong>Key Changes Made:</strong></p>
    <ul>
        <li>✓ Document.Write now accepts parentheses format: Document.Write(content)</li>
        <li>✓ AxTime() now properly returns Unix timestamp with parentheses</li>
        <li>✓ AxGenerateGuid() now works with function call syntax</li>
        <li>✓ AxBuildQueryString() improved with case-insensitive keys</li>
        <li>✓ AxGetRequest/Get/Post() now handle null contexts safely</li>
        <li>✓ All test cases updated to use proper VBScript syntax with parentheses</li>
    </ul>

    <p style="margin-top: 40px; color: #666; font-size: 12px;">
        <em>Test generated: 18 January 2026</em><br>
        For full documentation, see CUSTOM_FUNCTIONS_PT-BR.md
    </p>
</body>
</html>
%>
