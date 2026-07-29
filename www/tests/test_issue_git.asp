<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Option Explicit%>
<%
' Step 1: Create a Dictionary and store it into Session
Dim dict
Set dict = Server.CreateObject("Scripting.Dictionary")
dict.Add "title", "Hello World"
dict.Add "count", 42

Response.Write "=== Step 1: Store object into Session ===<br>"
Response.Write "TypeName: " & TypeName(dict) & "<br>"
Response.Write "dict.Count: " & dict.Count & "<br>"
Response.Write "dict('title'): " & dict("title") & "<br><br>"

Set Session("myDict") = dict

' Step 2: Retrieve from Session and inspect what we got back
Dim retrieved
Set retrieved = Session("myDict")

Response.Write "=== Step 2: Retrieve from Session ===<br>"
Response.Write "TypeName: " & TypeName(retrieved) & "<br>"
Response.Write "IsObject: " & IsObject(retrieved) & "<br>"
Response.Write "Count: " & retrieved.Count & "<br><br>"

' Step 3: Try to use it as an object — this triggers the error
Response.Write "=== Step 3: Try to use as object ===<br>"
Response.Write "Attempting retrieved.Count...<br>"

Dim c
c = retrieved.Count   ' <-- This line will fail

Response.Write "Count = " & c & "<br>"
%>