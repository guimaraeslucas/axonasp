<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<!-- #include file="includes/header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"-->
<%dim postback
postback=convertBool(request.form("postback"))

if postback then

	customer.sArrowUP=request.form("sArrowUP")

			
	if customer.save() then
		response.redirect("bs_arrowUp.asp?fbMessage=fb_saveOK")
	end if
end if

%>
<p align="center"></p>

<form method="post" action="bs_arrowup.asp" name="mainform">
<input type="hidden" name="postback" value="1" />
<table align="center" cellpadding="2">
	<tr>
		<td>Select arrow:</td>
		<td>
		<div style="text-align:center;float:left;width:80px">
			No arrow<br />
			<input  type="radio" <%if isLeeg(customer.sArrowUP) then response.write " checked=""checked"" "%> value="" name="sArrowUP" />
		</div>
		<br><br><br>
		
		<div class="cleared"></div>
		
		<%
		dim fso
		set fso=server.createobject("scripting.filesystemobject")
		dim folder
		set folder=fso.getFolder(server.mappath(C_DIRECTORY_QUICKERSITE & "/fixedImages/arrows"))
		dim file
		for each file in folder.files
		%>
		<div style="text-align:center;float:left;width:80px">
			<img style="margin:0 auto" src="<%=C_DIRECTORY_QUICKERSITE%>/fixedimages/arrows/<%=file.name%>" /><br />
			<input  type="radio" <%if file.name=customer.sArrowUP then response.write " checked=""checked"" "%> value="<%=file.name%>" name="sArrowUP" />
		</div>
		<%
		next
		%>		
		</td>
	</tr>
	<tr>
		<td class=QSlabel>&nbsp;</td>
		<td><input class="art-button" type=submit value="<%=l("save")%>" name="btnaction" /></td>
	</tr>
</table>
</form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
