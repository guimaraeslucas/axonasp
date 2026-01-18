<%
' ============================================================================
' G3 AxonASP - Corrected Functions Quick Reference
' ============================================================================
' This page shows the CORRECTED syntax for the 5 fixed functions
%>
<!DOCTYPE html>
<html>
<head>
    <title>Corrected Functions - G3 AxonASP</title>
    <style>
        * { margin: 0; padding: 0; }
        body { font-family: 'Segoe UI', Arial; background: #f5f5f5; padding: 20px; }
        .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h1 { color: #28a745; border-bottom: 3px solid #28a745; padding-bottom: 10px; margin-bottom: 30px; }
        .func { margin-bottom: 25px; border: 2px solid #28a745; border-radius: 5px; padding: 20px; background: #f8fff9; }
        .func h2 { color: #155724; font-size: 18px; margin-bottom: 10px; }
        .syntax { background: #f0f0f0; border-left: 4px solid #28a745; padding: 15px; margin: 10px 0; font-family: 'Courier New', monospace; }
        .example { background: #e8f5e9; border-left: 4px solid #81c784; padding: 15px; margin: 10px 0; }
        .wrong { background: #ffebee; border-left: 4px solid #e57373; padding: 15px; margin: 10px 0; color: #c62828; }
        .label { font-weight: bold; color: #333; margin-top: 10px; }
        code { background: #f0f0f0; padding: 2px 6px; border-radius: 3px; }
        .tick { color: #28a745; font-weight: bold; }
        .cross { color: #dc3545; font-weight: bold; }
    </style>
</head>
<body>
    <div class="container">
        <h1>✅ 5 Corrected Functions - Proper Syntax Guide</h1>

        <!-- Document.Write -->
        <div class="func">
            <h2>1. Document.Write</h2>
            <p><strong>Purpose:</strong> HTML-safe output with automatic encoding</p>
            
            <div class="label">❌ Wrong (Old):</div>
            <div class="wrong">
                Document.Write htmlContent
            </div>

            <div class="label"><span class="tick">✓</span> Correct (New):</div>
            <div class="syntax">
                Document.Write(htmlContent)<br>
                Document.Write("<strong>Bold</strong>")
            </div>

            <div class="label">Example:</div>
            <div class="example">
                <%
                    Dim userInput
                    userInput = "<img src=x onerror='alert(1)'>"
                    Response.Write "User input: "
                    Document.Write(userInput)
                    Response.Write "<br><em>(Encoded safely above)</em>"
                %>
            </div>
        </div>

        <!-- AxTime -->
        <div class="func">
            <h2>2. AxTime</h2>
            <p><strong>Purpose:</strong> Get current Unix timestamp (seconds since 1970)</p>
            
            <div class="label">❌ Wrong (Old):</div>
            <div class="wrong">
                timestamp = AxTime
            </div>

            <div class="label"><span class="tick">✓</span> Correct (New):</div>
            <div class="syntax">
                timestamp = AxTime()<br>
                Response.Write "Now: " & AxTime()
            </div>

            <div class="label">Example:</div>
            <div class="example">
                <%
                    Dim currentTime
                    currentTime = AxTime()
                    Response.Write "Current Unix Timestamp: " & currentTime & "<br>"
                    Response.Write "Type: " & TypeName(currentTime) & "<br>"
                %>
            </div>
        </div>

        <!-- AxGenerateGuid -->
        <div class="func">
            <h2>3. AxGenerateGuid</h2>
            <p><strong>Purpose:</strong> Generate unique GUID identifier</p>
            
            <div class="label">❌ Wrong (Old):</div>
            <div class="wrong">
                guid = AxGenerateGuid
            </div>

            <div class="label"><span class="tick">✓</span> Correct (New):</div>
            <div class="syntax">
                guid = AxGenerateGuid()<br>
                Dim id1, id2<br>
                id1 = AxGenerateGuid()<br>
                id2 = AxGenerateGuid()
            </div>

            <div class="label">Example:</div>
            <div class="example">
                <%
                    Dim guid1, guid2
                    guid1 = AxGenerateGuid()
                    guid2 = AxGenerateGuid()
                    Response.Write "GUID 1: " & guid1 & "<br>"
                    Response.Write "GUID 2: " & guid2 & "<br>"
                    Response.Write "Unique: " & (guid1 <> guid2) & "<br>"
                %>
            </div>
        </div>

        <!-- AxBuildQueryString -->
        <div class="func">
            <h2>4. AxBuildQueryString</h2>
            <p><strong>Purpose:</strong> Build URL query string from Dictionary</p>
            
            <div class="label">Pattern:</div>
            <div class="syntax">
                queryString = AxBuildQueryString(dictionary)<br>
                queryString = AxBuildQueryString(array)
            </div>

            <div class="label">Example with Dictionary:</div>
            <div class="example">
                <%
                    Dim params, qs
                    Set params = CreateObject("Scripting.Dictionary")
                    params("name") = "John Doe"
                    params("age") = 30
                    params("city") = "New York"
                    
                    qs = AxBuildQueryString(params)
                    Response.Write "Query String: " & qs & "<br>"
                    Response.Write "For URL: ?..." & qs & "<br>"
                %>
            </div>

            <div class="label">Example with Array:</div>
            <div class="example">
                <%
                    Dim arrParams
                    arrParams = Array("user", "john", "pass", "secret", "token", "abc123")
                    
                    Dim qs
                    qs = AxBuildQueryString(arrParams)
                    Response.Write "Query String: " & qs & "<br>"
                %>
            </div>

            <div class="label">Notes:</div>
            <ul>
                <li>Keys are automatically converted to lowercase</li>
                <li>Values are URL-encoded automatically</li>
                <li>Spaces become %20</li>
            </ul>
        </div>

        <!-- AxGetRequest -->
        <div class="func">
            <h2>5. AxGetRequest / AxGetGet / AxGetPost</h2>
            <p><strong>Purpose:</strong> Retrieve all request parameters as Dictionary</p>
            
            <div class="label">Patterns:</div>
            <div class="syntax">
                allParams = AxGetRequest()  ' GET + POST combined<br>
                getParams = AxGetGet()      ' Only GET parameters<br>
                postParams = AxGetPost()    ' Only POST parameters
            </div>

            <div class="label">Example - Get All Parameters:</div>
            <div class="example">
                <%
                    Dim allParams, count
                    allParams = AxGetRequest()
                    count = AxCount(allParams)
                    Response.Write "Total Parameters: " & count & "<br>"
                %>
            </div>

            <div class="label">Example - Get Specific Values:</div>
            <div class="example">
                <%
                    Dim getParams
                    getParams = AxGetGet()
                    
                    ' Check if parameter exists
                    If AxArrayContains("id", getParams) Then
                        Response.Write "Parameter 'id' exists<br>"
                    End If
                    
                    ' Note: Access via GetRequest/GetGet/GetPost returns a Dictionary
                    ' Use it to check what parameters you received
                %>
            </div>

            <div class="label">URL Examples:</div>
            <ul>
                <li><code>/page.asp?name=John&age=30</code> - AxGetGet() will have 2 params</li>
                <li><code>/page.asp?id=123</code> - AxGetRequest() will have 1 param (if no POST)</li>
                <li><code>Form POST data</code> - AxGetPost() will have the form fields</li>
            </ul>

            <div class="label">Notes:</div>
            <ul>
                <li>All keys are converted to lowercase for consistency</li>
                <li>Returns empty Dictionary if no parameters exist</li>
                <li>Safe to call even with no Request context</li>
            </ul>
        </div>

        <!-- Summary Table -->
        <div style="margin-top: 40px; background: #e3f2fd; padding: 20px; border-radius: 5px;">
            <h2 style="color: #1565c0; margin-bottom: 15px;">Quick Syntax Comparison</h2>
            <table style="width: 100%; border-collapse: collapse;">
                <tr style="background: #1565c0; color: white;">
                    <th style="padding: 10px; text-align: left;">Function</th>
                    <th style="padding: 10px; text-align: left;">Correct Syntax</th>
                    <th style="padding: 10px; text-align: left;">Returns</th>
                </tr>
                <tr style="background: #f5f5f5;">
                    <td style="padding: 10px; border: 1px solid #ddd;">Document.Write</td>
                    <td style="padding: 10px; border: 1px solid #ddd;"><code>Document.Write(text)</code></td>
                    <td style="padding: 10px; border: 1px solid #ddd;">None (outputs directly)</td>
                </tr>
                <tr>
                    <td style="padding: 10px; border: 1px solid #ddd;">AxTime</td>
                    <td style="padding: 10px; border: 1px solid #ddd;"><code>AxTime()</code></td>
                    <td style="padding: 10px; border: 1px solid #ddd;">Integer (Unix timestamp)</td>
                </tr>
                <tr style="background: #f5f5f5;">
                    <td style="padding: 10px; border: 1px solid #ddd;">AxGenerateGuid</td>
                    <td style="padding: 10px; border: 1px solid #ddd;"><code>AxGenerateGuid()</code></td>
                    <td style="padding: 10px; border: 1px solid #ddd;">String (GUID format)</td>
                </tr>
                <tr>
                    <td style="padding: 10px; border: 1px solid #ddd;">AxBuildQueryString</td>
                    <td style="padding: 10px; border: 1px solid #ddd;"><code>AxBuildQueryString(dict)</code></td>
                    <td style="padding: 10px; border: 1px solid #ddd;">String (URL query)</td>
                </tr>
                <tr style="background: #f5f5f5;">
                    <td style="padding: 10px; border: 1px solid #ddd;">AxGetRequest</td>
                    <td style="padding: 10px; border: 1px solid #ddd;"><code>AxGetRequest()</code></td>
                    <td style="padding: 10px; border: 1px solid #ddd;">Dictionary (params)</td>
                </tr>
            </table>
        </div>

        <p style="margin-top: 30px; color: #666; text-align: center; font-size: 12px;">
            <em>All functions now use proper VBScript syntax with parentheses.</em><br>
            <em>For complete documentation, see FIXES_SUMMARY.md</em>
        </p>
    </div>
</body>
</html>
%>
