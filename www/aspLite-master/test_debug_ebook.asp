<!-- #include file="aspLite/aspLite.asp"-->
<%
Response.Write "Start Debug FormMessages<br>"

dim fm : set fm = aspl.formmessages
Response.Write "FormMessages Count: " & fm.Count & "<br>"

Response.Write "Start Debug Fields Inc<br>"
dim asplEvent : asplEvent = aspl.getRequest("asplEvent")

if not aspl.isEmpty(asplEvent) then
    Dim path : path = "ebook/" & asplEvent & ".inc"
    Dim content : content = aspl.loadText(path)
    Dim code : code = aspl.removeCRB(content)
    
    Response.Write "Executing Global...<br>"
    On Error Resume Next
    ExecuteGlobal code
    If Err.Number <> 0 Then
        Response.Write "ExecuteGlobal Error: " & Err.Description & " (" & Err.Number & ")<br>"
    End If
    On Error Goto 0
    Response.Write "After ExecuteGlobal<br>"
else
    Response.Write "asplEvent is empty<br>"
end if
Response.Write "End Debug Fields Inc<br>"
%>
