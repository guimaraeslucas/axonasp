<%
@ CodePage = 65001
%>
<%
' Test Global.asa VM Support
' This page tests Session_OnStart, Session_OnEnd, Application_OnStart, Application_OnEnd
' and other Global.asa features when running under AxonVM

Response.ContentType = "text/html; charset=utf-8"
%>
<!DOCTYPE html>
<html>
    <head>
        <title>Test Global.asa VM Support</title>
        <style>
            body {
                font-family: Tahoma, sans-serif;
                background: #ece9d8;
                margin: 20px;
            }
            .section {
                border: 1px solid #808080;
                padding: 15px;
                margin: 15px 0;
                background: #fff;
            }
            .header {
                background: linear-gradient(to right, #003399, #3366cc);
                color: white;
                padding: 10px;
                margin: -15px -15px 15px -15px;
            }
            .pass {
                color: green;
                font-weight: bold;
            }
            .fail {
                color: red;
                font-weight: bold;
            }
            .info {
                color: #335ea8;
                font-weight: bold;
            }
            table {
                width: 100%;
                border-collapse: collapse;
            }
            td,
            th {
                border: 1px solid #999;
                padding: 8px;
                text-align: left;
            }
            th {
                background: #ddd;
            }
        </style>
    </head>
    <body>
        <div class="section">
            <div class="header">AxonVM Global.asa Support Test</div>

            <p>
                This test verifies the Global.asa functionality when running
                under AxonVM.
            </p>

            <h3>Session State</h3>
            <table>
                <tr>
                    <td><strong>Session ID:</strong></td>
                    <td><%= Session.SessionID %></td>
                </tr>
                <tr>
                    <td><strong>Session Timeout:</strong></td>
                    <td>
                        <%= Session.Timeout %>
                        minutes
                    </td>
                </tr>
                <tr>
                    <td><strong>Session LCID:</strong></td>
                    <td><%= Session.LCID %></td>
                </tr>
                <tr>
                    <td><strong>Session CodePage:</strong></td>
                    <td><%= Session.CodePage %></td>
                </tr>
            </table>

            <h3>Testing Session_OnStart</h3>
            <p>
                The Session_OnStart event should set a variable
                'SessionStartCount' if it was executed.
            </p>
            <table>
                <tr>
                    <td><strong>SessionStartCount:</strong></td>
                    <td>
                        <%
                        If Session.Contents("SessionStartCount") <> "" Then
                        %>
                        <span class="pass"
                            >✓ SET to
                            <%= Session.Contents("SessionStartCount") %></span
                        >
                        <%
                        Else
                        %>
                        <span class="fail"
                            >✗ NOT SET - Session_OnStart may not have been
                            executed</span
                        >
                        <%
                        End If
                        %>
                    </td>
                </tr>
            </table>

            <h3>Testing Session Variables</h3>
            <p>You can set session variables here to test persistence:</p>
            <table>
                <tr>
                    <td><strong>Current Session Variables:</strong></td>
                    <td>
                        <%
                        Dim objKeys, objValues, i
                        objKeys = Session.Contents.Keys()

                        If IsArray(objKeys) Then
                            If UBound(objKeys) >  = 0 Then
                        %>
                        <div
                            style="
                                background: #f5f5f5;
                                padding: 10px;
                                border: 1px solid #ccc;
                            ">
                            <%
                                    For i = 0 To UBound(objKeys)
                                        Dim varName, varValue
                                        varName = objKeys(i)
                                        varValue = Session.Contents(varName)
                            %>
                            <div>
                                <strong
                                    ><%= varName %>:</strong
                                >
                                <%= varValue %>
                            </div>
                            <%
                                    Next
                            %>
                        </div>
                        <%
                            Else
                        %>
                        <em
                            >(No session variables set - Session.Contents is
                            empty)</em
                        >
                        <%
                            End If
                        Else
                        %>
                        <em>Error: Contents.Keys() did not return an array</em>
                        <%
                        End If
                        %>
                    </td>
                </tr>
            </table>

            <h3>Testing Session.Abandon()</h3>
            <p>Click the button below to test Session.Abandon():</p>
            <form method="post" style="display: inline">
                <input type="hidden" name="action" value="abandon" />
                <input
                    type="submit"
                    value="Test Session.Abandon()"
                    style="padding: 5px 15px" />
            </form>

            <%
            If Request.Form("action") = "abandon" Then
                Dim oldSessionID
                oldSessionID = Session.SessionID
            %>
            <div
                style="
                    background: #e8f5e9;
                    border: 1px solid #4caf50;
                    padding: 10px;
                    margin-top: 10px;
                ">
                <strong>Session.Abandon() was called</strong><br />
                <em>New session ID will be generated on next request</em><br />
                <strong>Old Session ID:</strong>
                <%= oldSessionID %>
            </div>
            <%
                Session.Abandon()
            End If
            %>

            <h3>Testing Session.Contents.RemoveAll()</h3>
            <p>Click the button below to clear all session variables:</p>
            <form method="post" style="display: inline">
                <input type="hidden" name="action" value="removeall" />
                <input
                    type="submit"
                    value="Session.Contents.RemoveAll()"
                    style="padding: 5px 15px" />
            </form>

            <%
            If Request.Form("action") = "removeall" Then
            %>
            <div
                style="
                    background: #fff3cd;
                    border: 1px solid #ffc107;
                    padding: 10px;
                    margin-top: 10px;
                ">
                <strong>Session.Contents.RemoveAll() was called</strong><br />
                <em>All session variables have been cleared</em>
            </div>
            <%
                Session.Contents.Removeall()
            End If
            %>

            <h3>Application Object Test</h3>
            <table>
                <tr>
                    <td><strong>Application('GlobalCounter'):</strong></td>
                    <td>
                        <%
                        If Application("GlobalCounter") <> "" Then
                        %>
                        <span class="pass"
                            >✓ Value:
                            <%= Application("GlobalCounter") %></span
                        >
                        <%
                        Else
                        %>
                        <span class="fail">✗ NOT SET</span>
                        <%
                        End If
                        %>
                    </td>
                </tr>
            </table>

            <h3>Server Information</h3>
            <table>
                <tr>
                    <td><strong>Server Name:</strong></td>
                    <td><%= Request.ServerVariables("SERVER_NAME") %></td>
                </tr>
                <tr>
                    <td><strong>Server Port:</strong></td>
                    <td><%= Request.ServerVariables("SERVER_PORT") %></td>
                </tr>
                <tr>
                    <td><strong>Script Path:</strong></td>
                    <td><%= Request.ServerVariables("SCRIPT_NAME") %></td>
                </tr>
            </table>
        </div>
    </body>
</html>
