<%
Option Explicit
Response.LCID = 1046 ' Brazilian LCID (use your locale code here).
' Could also be the LCID property of the page declaration or Session.LCID to set it to the entire session.
Response.buffer = True
%>
<!--#include file="tmp_jsonObject_oern.class.asp" -->
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <title>ASPJSON</title>

        <style type="text/css">
            body {
                font-family: Arial, Helvetica, sans-serif;
            }

            pre {
                border: solid 1px #cccccc;
                background-color: #eee;
                padding: 5px;
                text-indent: 0;
                width: 90%;
                white-space: pre-wrap;
                word-wrap: break-word;
            }
        </style>
    </head>
    <body>
        <h1>JSON Object and Array Tests</h1>
        <%
        Server.ScriptTimeout = 10
        Dim jsonObj, jsonString, jsonArr, outputObj
        Dim testLoad, testAdd, testRemove, testValue, testChange, testArrayPush, testLoadRecordset
        Dim testLoadArray, testChangeDefaultPropertyName, testGetItemAt

        testLoad = True
        testLoadArray = True
        testAdd = True
        testRemove = True
        testValue = True
        testChange = True

        testArrayPush = True

        testLoadRecordset = True

        testChangeDefaultPropertyName = True

        Set jsonObj = New JSONobject
        Set jsonArr = New jsonArray

        jsonObj.debug = False

        If testLoad Then
            jsonString = "{ ""strings"" : ""valorTexto"", ""numbers"": 123.456, ""bools"": true, ""arrays"": [1, ""2"", 3.4, [5, -6, [7, 8, [[[""9"", ""10""]]]]]], ""emptyArray"": [], ""emptyObject"": {}, ""objects"": { ""prop1"": ""outroTexto"", ""prop2"": [ { ""id"": 1, ""name"": ""item1"" }, { ""id"": 2, ""name"": ""item2"", ""teste"": { ""maisum"": [1, 2, 3] } } ] }, ""multiline"": ""Texto com\r\nMais de\r\numa linha e escapado com \\."" }"

            If testLoadArray Then jsonString = "[" & jsonString & "]"

            Dim start
            start = Timer()
            Set outputObj = jsonObj.parse(jsonString)

            If testLoadArray And Left(jsonString, 1) <> "[" Then jsonString = "[" & jsonString & "]"
        %>
        <h3>Parse Input</h3>
        <pre>        <%= jsonString %></pre>
        <%
            Response.flush

            Dim start
            start = Timer()
            Set outputObj = jsonObj.parse(jsonString)
            If testLoadArray Then Set jsonArr = outputObj

            Response.Write "Load time: " & (Timer() - start) & " s<br>"
        End If

        If testAdd Then
            Dim arr, multArr, nestedObject
            arr = Array(1, "teste", 234.56, "mais teste", "234", Now)

            ReDim multArr(2, 3)
            multArr(0, 0) = "0,0"
            multArr(0, 1) = "0,1"
            multArr(0, 2) = "0,2"
            multArr(0, 3) = "0,3"

            multArr(1, 0) = "1,0"
            multArr(1, 1) = "1,1"
            multArr(1, 2) = "1,2"
            multArr(1, 3) = "1,3"

            multArr(2, 0) = "2,0"
            multArr(2, 1) = "2,1"
            multArr(2, 2) = "2,2"
            multArr(2, 3) = "2,3"

            jsonObj.Add "nome", "Jozé"
            jsonObj.Add "ficticio", True
            jsonObj.Add "idade", 25
            jsonObj.Add "saldo", -52
            jsonObj.Add "bio", "Nascido em São Paulo\Brasil" & vbcrlf & "Sem filhos" & vbcrlf & vbtab & "Jogador de WoW"
            jsonObj.Add "data", Now
            jsonObj.Add "lista", arr
            jsonObj.Add "lista2", multArr

            Set nestedObject = New JSONobject
            nestedObject.Add "sub1", "value of sub1"
            nestedObject.Add "sub2", "value of ""sub2"""

            jsonObj.Add "nested", nestedObject
        End If

        If testRemove Then
            jsonObj.Remove "numbers"
            jsonObj.Remove "aNonExistantPropertyName" ' this sould run silently, even To non existant properties
        End If

        If testChangeDefaultPropertyName Then
            jsonObj.defaultPropertyName = "CustomName"
            jsonArr.defaultPropertyName = "CustomArrName"
        End If

        If testValue Then
        %>
        <h3>Get Values</h3>
        <%
            Response.Write "nome: " & jsonObj.value("nome") & "<br>"
            Response.Write "idade: " & jsonObj("idade") & "<br>" ' short syntax
            Response.Write "non existant property:" & jsonObj("aNonExistantPropertyName") & "(" & TypeName(jsonObj("aNonExistantPropertyName")) & ")<br>"

            If IsObject(jsonObj(jsonObj.defaultPropertyName)) Then
                Response.Write "default property name (" & jsonObj.defaultPropertyName & "): <pre>" & jsonObj(jsonObj.defaultPropertyName).Serialize() & "</pre>"
            Else
                Response.Write "default property name (" & jsonObj.defaultPropertyName & "):" & jsonObj(jsonObj.defaultPropertyName) & "<br>"
            End If
        End If

        If testChange Then
        %>
        <h3>Change Values</h3>
        <%

            Response.Write "nome before: " & jsonObj.value("nome") & "<br>"

            jsonObj.change "nome", "Mario"

            Response.Write "nome after: " & jsonObj.value("nome") & "<br>"

            jsonObj.change "nonExisting", -1

            Response.Write "Non existing property is created with: " & jsonObj.value("nonExisting") & "<br>"
        End If

        If testArrayPush Then
            Dim newJson
            Set newJson = New JSONobject
            newJson.Add "newJson", "property"
            newJson.Add "version", newJson.version

            jsonArr.Push newJson
            jsonArr.Push 1
            jsonArr.Push "strings too"
        End If

        If testLoadRecordset Then
        %>
        <h3>Load a Recordset</h3>
        <!--
		   METADATA
		   TYPE="TypeLib"
		   NAME="Microsoft ActiveX Data Objects 2.5 Library"
		   UUID="{00000205-0000-0010-8000-00AA006D2EA4}"
		   VERSION="2.5"
		-->
        <%
            Dim rs
            Set rs = CreateObject("ADODB.Recordset")

            ' prepera an in memory recordset
            ' could be, and mostly, loaded from a database
            rs.CursorType = adOpenKeyset
            rs.CursorLocation = adUseClient
            rs.LockType = adLockOptimistic

            rs.Fields.Append "ID", adInteger, , adFldKeyColumn
            rs.Fields.Append "Nome", adVarChar, 50, adFldMayBeNull
            rs.Fields.Append "Valor", adDecimal, 14, adFldMayBeNull
            rs.Fields("Valor").NumericScale = 2

            rs.Open

            rs.AddNew
            rs("ID") = 1
            rs("Nome") = "Nome 1"
            rs("Valor") = 10.99
            rs.Update

            rs.AddNew
            rs("ID") = 2
            rs("Nome") = "Nome 2"
            rs("Valor") = 29.90
            rs.Update

            rs.MoveFirst
            jsonObj.LoadFirstRecord rs
            ' or
            rs.MoveFirst
            jsonArr.LoadRecordSet rs

            rs.Close

            Set rs = Nothing
        End If

        If testLoad Then
            start = Timer()
        %>
        <h3>Parse Output</h3>
        <pre>        <%= outputObj.Write %></pre>
        <%
            Response.Write Timer() - start
            Response.Write " s<br>"
            Response.flush
        End If
        %>

        <h3>
            JSON Object Output
<%
If testLoad Then
%>
            
            (Same as parse output:
<%
    If TypeName(jsonObj) = TypeName(outputObj) Then
%>yes
<%
    Else
%>
no
<%
    End If
%>
)
<%
End If
%>

        </h3>
        <%
        jsonString = jsonObj.Serialize()
        %>
        <pre><%= Left(jsonString, 2000) %>
<%
If Len(jsonString) > 2000 Then
%>
        ... (too long, truncated)
<%
End If
%>
        </pre>
        <%
        Response.flush
        %>

        <h3>
            Array Output
<%
If testLoad Then
%>
            
            (Same as parse output:
<%
    If TypeName(jsonArr) = TypeName(outputObj) Then
%>yes
<%
    Else
%>
no
<%
    End If
%>
)
<%
End If
%>

        </h3>
        <%
        jsonString = jsonArr.Serialize()
        %>
        <pre><%= Left(jsonString, 2000) %>
<%
If Len(jsonString) > 2000 Then
%>
        ... (too long, truncated)
<%
End If
%>
        </pre>
        <%
        Response.flush
        %>

        <h3>Array Loop</h3>
        <pre>
<%
Dim i, Items, Item

' more readable loop
i = 0
Response.Write "For Each Loop (readability):<br>==============<br>"
start = Timer()
For Each Item In jsonArr.Items
    Response.Write "Index "
    Response.Write i
    Response.Write ": "

    If IsObject(Item) And TypeName(Item) = "JSONobject" Then
        Item.Write()
    Else
        Response.Write Item
    End If

    Response.Write "<br>"
    i = i + 1
    If i Mod 100 = 0 Then Response.flush
Next
Response.Write Timer() - start
Response.Write " s<br>"

Response.Write "<br><br>For Loop (speed):<br>=========<br>"
start = Timer()

' faster but less readable
For i = 0 To jsonArr.length - 1
    Response.Write "Index "
    Response.Write i
    Response.Write ": "

    If IsObject(jsonArr(i)) Then
        Set Item = jsonArr(i)

        If TypeName(Item) = "JSONobject" Then
            Item.Write()
        Else
            Response.Write Item
        End If
    Else
        Item = jsonArr(i)
        Response.Write Item
    End If

    Response.Write "<br>"
    If i Mod 100 = 0 Then Response.flush
Next
Response.Write Timer() - start
Response.Write " s<br>"

Set newJson = Nothing
Set outputObj = Nothing
Set jsonObj = Nothing
Set jsonArr = Nothing
%>
        </pre>

        <h3>JSON Script Output</h3>

        <%
        Dim realOutput
        Dim expectedOutput

        Dim javascriptCode
        Dim javascriptkey

        Dim jsonScr

        javascriptCode = "function() { alert('test'); }"
        javascriptKey = "script"

        expectedOutput = "{""" & javascriptKey & """:" & javascriptCode & "}"

        Set jsonScr = New JSONscript
        jsonScr.value = javascriptCode

        Set jsonObj = New JSONobject
        jsonObj.Add javascriptKey, jsonScr

        realOutput = jsonObj.Serialize()

        %>
        <h4>
            Output
<%
If (realOutput = expectedOutput) Then
%>
            
            (correct)
<%
Else
%>
            
            (INCORRECT!)
<%
End If
%>
            
        </h4>
        <pre><%= realOutput %></pre>
    </body>
</html>

