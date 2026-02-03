<%
if not session("authenticated") then Response.End 
%>
<!--#include file="code.asp"-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" dir="ltr" lang="en-US" xml:lang="en">
<head><title>PrivateMailBot - Upload</title></head>
<body bgcolor="#C0D7DD" onload="javascript:window.focus();">
<h4>Upload file:</h4>
<%
select case Request.QueryString ("message")
	case "nofile"
		Response.Write "<p><font color='Red'><b>Please select a file!</b></font></p>"
	case "filetobig"
		Response.Write "<p><font color='Red'><b>The file exceeds 10 MB!</b></font></p>"
	case "filetype"
		Response.Write "<p><font color='Red'><b>Only .mdb, .xls .txt and .htm files please!</b></font></p>"
	case "ok"
		Response.Write "<p><font color='Green'><b>Upload completed!</b></font></p>"
end select
%>
<form method="post" enctype="multipart/form-data" action="uploadProcess.asp"> 
<table align="center" cellpadding="5">
	
	<tr>
		<td class="label">Select a file:</td>
		<td align="left"><input type="file" name="newfile" /><br /><font style="font-size:smaller">.mdb, .xls, .txt and .htm files - max 10 MB</font></td>		
	</tr>	
	<tr>
		<td>&nbsp;</td>
		<td><input type="submit" value="Upload" onclick="javascript:getElementById('pleasewait').style.display='block';" /></td>	
	</tr>
	<tr>	
		<td>&nbsp;</td>
		<td><div style="margin:2px;color:Red;display:none" id="pleasewait"><b>Please wait...</b></div></td>
	</tr>
</table>
</form>
<%
dim pb, file
set pb=new PrivateBot

if Request.querystring("del") <>"" then
	pb.deleteFile(Request.querystring("del"))
end if

if pb.getFolder.files.count>0 then
%>
	<h4>Files:</h4>
	<ul>
	<%
	for each file in pb.getFolder.files
	%>
	<li><%=file.name%> (<%=round((file.size/1024),2)%> KB) - <a onclick="javascript:if(confirm('Are you sure to delete this file?')){location.assign(this.href);}else{return false;}" href="upload.asp?del=<%=server.URLEncode (file.name)%>">del</a></li>
	<%
	next
	%>
	</ul>
<%
end if
%>
</body>
</html>