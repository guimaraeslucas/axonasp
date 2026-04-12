<%
Option Explicit
dim jsonObj, jsonString, outputObj
set jsonObj = new JSONobject
jsonObj.debug = false

Response.Write "Before parse: TypeName(jsonObj) = " & TypeName(jsonObj) & "<br>"

jsonString = "[{" &"""test""" & ": " & """value""" & "}]"
set outputObj = jsonObj.parse(jsonString)

Response.Write "After parse: TypeName(jsonObj) = " & TypeName(jsonObj) & "<br>"
Response.Write "jsonObj is still JSONobject: " & (TypeName(jsonObj) = "JSONobject") & "<br>"

Response.Write "Creating JSONpair<br>"
dim pair
set pair = new JSONpair
Response.Write "JSONpair created: " & TypeName(pair) & "<br>"

Response.Write "Setting pair.name<br>"
pair.name = "test"
Response.Write "pair.name = " & pair.name & "<br>"
%>
<!--#include file="json-teste/jsonObject.class.asp" -->
