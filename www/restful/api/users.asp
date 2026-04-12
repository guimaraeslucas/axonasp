<%
@Language = "VBSCRIPT" CodePage = "65001"
%>
<%
Option Explicit

Dim json
Dim payload
Dim methodName
Dim userId
Dim Output

Response.ContentType = "application/json"
Response.Charset = "utf-8"

Set json = Server.CreateObject("G3JSON")
Set payload = json.NewObject()

methodName = UCase(Request.ServerVariables("REQUEST_METHOD"))

If methodName = "CLI" Then
    methodName = "GET"
End If

If methodName = "GET" Then
    userId = Trim(Request.QueryString("id"))

    If userId = "" Then
        Output = "{""success"":true,""resource"":""users"",""count"":2,""data"": [{""id"":1,""name"":""Alice""},{""id"":2,""name"":""Bob""}]}"
        Response.Write Output
    Else
        If userId = "1" Then
            Output = "{""success"":true,""resource"":""users"",""data"": {""id"":1,""name"":""Alice""}}"
            Response.Write Output
        Else
            If userId = "2" Then
                Output = "{""success"":true,""resource"":""users"",""data"": {""id"":2,""name"":""Bob""}}"
                Response.Write Output
            Else
                Response.Status = "404 Not Found"
                payload("success") = False
                payload("resource") = "users"
                payload("message") = "User not found"
                payload("id") = userId
                Response.Write json.Stringify(payload)
            End If
        End If
    End If
Else
    If methodName = "POST" Then
        Dim nameParam

        nameParam = Trim(Request.Form("name"))

        If nameParam = "" Then
            Response.Status = "400 Bad Request"
            payload("success") = False
            payload("resource") = "users"
            payload("message") = "name field is required"
            Response.Write json.Stringify(payload)
        Else
            payload("success") = True
            payload("resource") = "users"
            payload("message") = "User created"
            payload("id") = 3
            payload("name") = nameParam
            Response.Status = "201 Created"
            Response.Write json.Stringify(payload)
        End If
    Else
        Response.Status = "405 Method Not Allowed"
        payload("success") = False
        payload("resource") = "users"
        payload("message") = "Method not allowed"
        payload("method") = methodName
        Response.Write json.Stringify(payload)
    End If
End If

Set payload = Nothing
Set json = Nothing
%>
