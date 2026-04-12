<%
Option Explicit
dim jsonObj
set jsonObj = new JSONobject
jsonObj.debug = false

Response.Write "About to call add<br>"

On Error Resume Next
jsonObj.add "nome", "test"
 
if Err.Number <> 0 then
    Response.Write "ERROR: " & Err.Description & "<br>"
else
    Response.Write "Add succeeded<br>"
end if
%>
<!--#include file="json-teste/jsonObject.class.asp" -->
