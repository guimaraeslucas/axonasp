<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<%if convertBool(Session(cId & "isAUTHENTICATEDSecondAdmin")) then response.end %>
<!doctype html>
<html>
<head>
<%

dim mode
mode=""

select case lcase(GetFileExtension(request("path")))

case "js"
mode="text/javascript"

case "html","htm","txt"
mode="text/html"

case "css"
mode="text/css"

end select

'security 1
if mode="" then response.end 'not allowed file extension 

session("bHasSetUF")=false
    
dim fso,filecontent,fileobj,filename,filepath
set fso=server.createobject("scripting.filesystemobject")

'security
if not fso.fileexists(server.mappath(request("path"))) then
response.write "file not found"
set fso=nothing
response.end 
end if

'security 3

if left(request("path"),len(C_VIRT_DIR & Application("QS_CMS_userfiles")))<>C_VIRT_DIR & Application("QS_CMS_userfiles") then
response.write "file not allowed"
set fso=nothing
response.end 
end if

set fileobj=fso.getFile(server.mappath(request("path")))
filename=fileobj.name
filepath=fileobj.parentfolder
set fileobj=nothing

'processing
'if request.form("btnaction")="Back" then
'dim pp
'pp=request("path")
'pp=replace(pp,filename,"",1,-1,1)
'response.redirect("assetmanager/assetmanagerIF.asp?QS_inpFolder=" & server.mappath(pp))
'end if

dim file
set file=fso.openTextFile(server.mappath(request("path")),1)

on error resume next
filecontent=file.readAll()
set file=nothing
on error goto 0

'copy file
if request.querystring("copy")="1" then

newFileName=left(filename,len(filename)-len(getFileExtension(filename))-1)

dim counter
counter=""
while fso.fileexists(filepath & "\" & newFileName & " - copy" & counter & "." & getFileExtension(filename))
if counter="" then counter=0
counter=cint(counter)+1
wend

dim newFileName
if cstr(counter)="" then
newFileName=newFileName & " - copy"
else
newFileName=newFileName & " - copy" & counter
end if

dim copyfile
set copyfile=fso.createTextFile(filepath & "\" & newFileName & "." & getFileExtension(filename),2)
copyfile.write filecontent
copyfile.close
set copyfile=nothing
response.redirect("bs_editcode.asp?copied=1&path="  & server.urlencode(request("path")))

end if

dim fbmessage
'process save button
if request.form("btnAction")="Save" then
    filecontent=request.form("code")
    set file=fso.openTextFile(server.mappath(request("path")),2)
    file.write(filecontent)
    file.close
    set file=nothing
    fbmessage="<div style=""background-color:Yellow;color:Green"">&nbsp;file is saved&nbsp;</div>"
end if

if request.form("lineWrapping")<>"" then
filecontent=request.form("code")
end if

set file=nothing
set fso=nothing

if request.querystring("copied")="1" then
fbmessage="<div style=""background-color:Yellow;color:Green"">&nbsp;file is copied&nbsp;</div>"
end if

Function GetFileExtension(sFileName)
Dim Pos
Pos = instrrev(sFileName, ".")
If Pos>0 Then GetFileExtension = Mid(sFileName, Pos+1)
End Function
   
%>
<meta charset="utf-8"/>
<link rel="stylesheet" href="codemirror/lib/codemirror.css">
<link rel="stylesheet" href="codemirror/addon/display/fullscreen.css">
<script src="codemirror/lib/codemirror.js"></script>
<script src="codemirror/mode/xml/xml.js"></script>
<script src="codemirror/mode/javascript/javascript.js"></script>
<script src="codemirror/mode/css/css.js"></script>
<script src="codemirror/mode/htmlembedded/htmlembedded.js"></script>
<script src="codemirror/addon/display/fullscreen.js"></script>
</head>
<body>
<article>
    <div style="background-color: #808080;color: #fff ">&nbsp;<%=server.htmlencode(request("path"))%></div>
<form method="post" action="bs_editcode.asp" name=codemirror>
    <input type="hidden" name="path" value="<%=server.htmlencode(request("path"))%>" />
	<input type="hidden" name="lineWrapping" value="<%=sanitize(request("lineWrapping"))%>" />
<textarea id="code" name="code" rows="5"><%=filecontent%></textarea>
    <table>
        <tr>
            <td><input type="submit" name="btnAction" value ="Save" /></td>
            <td><input type="submit" onclick="window.close();return false;" name="btnAction" value="Close" /></td>
			<td><input type="checkbox" <%if convertBool(request.form("lineWrapping")) then response.write " checked=""checked"" "%> onclick="document.codemirror.lineWrapping.value=this.checked;document.codemirror.submit();" /> wrap text</td>
            <td><a onclick="return confirm('Are you sure to copy this file?')" href="bs_editcode.asp?copy=1&amp;path=<%=server.urlencode(request("path"))%>">copy file</a></td>
			<td><%=fbmessage%></td>
        </tr>
    </table> 
</form>
<script type="text/javascript">

var editor = CodeMirror.fromTextArea(document.getElementById("code"), {
    lineNumbers: true,
    theme: "default",  
    mode: "<%=mode%>",<%
	if convertBool(request("lineWrapping")) then
	%>
	lineWrapping: true,<%
	end if
	%>
    extraKeys: {
        "F11": function (cm) {
            cm.setOption("fullScreen", !cm.getOption("fullScreen"));
        },
        "Esc": function (cm) {
            if (cm.getOption("fullScreen")) cm.setOption("fullScreen", false);
        }
    }
});
editor.setSize("100%", "520px");

</script>

  </article>
</body>
</html>