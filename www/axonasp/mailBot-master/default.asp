<!--#include file="start.asp"-->
<!--#include file="top.asp"-->
<!--#include file="code.asp"-->
<!--#include file="freeASPUpload.asp"-->
<%
if request("btnAction")="" then
	Response.Redirect ("default.asp?btnAction=About")
end if

if not session("authenticated") and request("btnAction")<>"About" then

	%>
	<h2 class="art-PostHeaderIcon-wrapper">
	    <img src="images/PostHeaderIcon.png" width="12" height="12" alt="PostHeaderIcon" />
	    <span class="art-PostHeader">Login</span>
	</h2>
	<div class="art-PostContent">
	<%	
	dim pw
	if Request.Form ("btnAction")="Login" then
		if Request.Form ("password")=adminpassword then
			session("authenticated")=true
			Response.Redirect "default.asp?btnAction=QuickSearch"
		else
			%>
			<p><font color="Red">Wrong password...</font></p>
			<%
		end if
	end if
	
	if adminpassword="admin" then
		pw=adminpassword
		%>
		<p><font color="Red">WARNING: The default password "<strong>admin</strong>" is still active. Make sure to change the password in the file "config.asp"</font></p>
		<%
	else
		pw=Request.Form ("password")
	end if
	%>
	<form action="default.asp" method="post" name="mainform">
		<input type="hidden" value="Login" name="btnAction" />
		<table>
			<tr>
				<td>Password:</td>
				<td><input type="password" name="password" value="<%=sanitize(pw)%>" /></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td><input type="submit" value="Login" name="dummy" /></td>
			</tr>
		</table>
	
		<script type="text/javascript">document.mainform.password.focus()</script>					
				
	</form>
<%
else
	dim pb
	set pb=new privatebot

	select case request("btnAction")

		case "logoff"
			
			session("authenticated")=false
			Response.Redirect "default.asp"
			
		case "","QuickSearch"
			%>
			<h2 class="art-PostHeaderIcon-wrapper">
			    <img src="images/PostHeaderIcon.png" width="12" height="12" alt="PostHeaderIcon" />
			    <span class="art-PostHeader">QuickSearch</span>
			</h2>
			<div class="art-PostContent">
			<p>Copy/Paste any text (or html) in the field below, next click "QuickSearch". This routine searches for email addresses in your text.</p>
			<form action="default.asp" method="post" name="mainform">
			<textarea cols="90" rows="10" name="freetext"><%=sanitize(Request.Form ("freetext"))%></textarea><br />
			<input type="submit" value="QuickSearch" name="btnAction" />
			<input type="reset" value="Clear" onclick="javascript:document.mainform.freetext.value='';return false;" />
			</form>
			<script type="text/javascript">document.mainform.freetext.focus()</script>
			<%

			if Request.Form ("btnAction")="QuickSearch" then
				pb.quickSearch(Request.Form ("freetext"))
			end if
		
		case "FileSearch"
			%>
			<h2 class="art-PostHeaderIcon-wrapper">
			    <img src="images/PostHeaderIcon.png" width="12" height="12" alt="PostHeaderIcon" />
			    <span class="art-PostHeader">FileSearch</span>
			</h2>
			<div class="art-PostContent">

			<form action="default.asp" method="post" name="mainform">
			<p>This routine searches for email addresses in all the files in the "files" folder. You can search Access (.mdb), Excel (.xls), text (.txt) and html (.htm) files.<br />You can <b>exclude</b> email addresses (containing any of the characters you specify) or <b>only include</b> specific email addresses. You can use a comma-separated list of values as well. </p>
			<table>
				<tr>
					<td>Exclude:</td>
					<td><input type="text" size="130" name="exclude" value="<%=sanitize(Request.Form ("exclude"))%>" /><br />eg: no-reply@,postmaster@,anytext,anydomain</td>
				</tr>
				<tr>
					<td>Only include:</td>
					<td><input type="text" size="130" name="include" value="<%=sanitize(Request.Form ("include"))%>" /><br />eg: @hotmail.com,@gmail.com,anytext,anydomain</td>
				</tr>
				<tr>
					<td>&nbsp;</td>
					<td><input type="submit" value="FileSearch" name="btnAction" onclick="javascript:getElementById('pleasewait').style.display='block';" />&nbsp;&nbsp;&nbsp;<a onclick="window.open(this.href,'','resizable=yes,location=no,menubar=no,scrollbars=yes,status=no,toolbar=no,fullscreen=no,dependent=no,width=450,height=450,status'); return false" href="upload.asp">Manage files</a></td>
				</tr>
				<tr>
					<td valign="top">Select file(s):</td>
					<td>
						<%		
						dim checked
						if Request.Form ("btnAction")="" then
							checked=" checked='checked' "			
						end if	
						
						dim file
						for each file in pb.getFolder.files
						
							select case lcase(GetFileExtension(file))
								case "mdb","xls","htm","txt"
									%>
									<input type="checkbox" <%=checked%> <%if instr(Request.Form("selectedfiles"),file.name)<>0 then Response.Write " checked='checked' "%> name="selectedfiles" value="<%=sanitize(file.name)%>" /><%=file.name%> (<%=round((file.size/1024),2)%> KB)<br />
									<%
							end select
						next
						%>
					</td>
				</tr>	
				<tr>
					<td>&nbsp;</td>
					<td><div style="color:Red;display:none" id="pleasewait"><b>Please wait...</b></div></td>
				</tr>		
			</table>
			</form>
			<script type="text/javascript">document.mainform.exclude.focus()</script>
			<%
			if Request.Form ("btnAction")="FileSearch" then
				pb.selectedfiles=Request.Form("selectedfiles")
				pb.filter=Request.Form ("exclude")
				pb.only=Request.Form ("include")			
				pb.getEmailList
			end if
		
		case "About"
			%>
			<h2 class="art-PostHeaderIcon-wrapper">
			    <img src="images/PostHeaderIcon.png" width="12" height="12" alt="PostHeaderIcon" />
			    <span class="art-PostHeader">About</span>
			</h2>
			<div class="art-PostContent">
			<!--#include file="readme.md"-->
			<%
	end select

	set pb=nothing

end if
%>
<!--#include file="bottom.asp"-->
