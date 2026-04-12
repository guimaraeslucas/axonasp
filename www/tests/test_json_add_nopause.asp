<%
Option Explicit
dim jsonObj
set jsonObj = new JSONobject
jsonObj.debug = false

On Error Resume Next
jsonObj.add "nome", "Jozé"

if Err.Number <> 0 then
    Response.Write "ERROR: " & Err.Description & "<br>"
else
    Response.Write "Add succeeded<br>"
end if
%>
<!--#include file="json-teste/jsonObject.class.asp" -->
