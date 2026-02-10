<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCustom404%>
<!-- #include file="includes/header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"-->
<%dim postback
postback=convertBool(request.form("postback"))

if postback then

	customer.bCustom404=convertBool(request.form("bCustom404"))
	customer.sCustom404Title=left(request.form("sCustom404Title"),250)
	customer.sCustom404Body=request.form("sCustom404Body")
	customer.i404TemplateID=convertGetal(decrypt(request.form("i404TemplateID")))
			
	if customer.save() then
		response.redirect("bs_404.asp?fbMessage=fb_saveOK")
	end if
end if

if isLeeg(customer.sCustom404Body) and isLeeg(customer.sCustom404Title) then
customer.sCustom404Title="File Not Found"
customer.sCustom404Body="The file ""[404FILENAME]"" you searched for cannot be found."
end if


%>
<p align="center"></p>

<form method="post" action="bs_404.asp" name="mainform">
<input type="hidden" name="postback" value="1" />
<table align="center" cellpadding="2">
<tr><td class=QSlabel valign=top>Enable custom 404 page?</td><td><input name="bCustom404" type="checkbox" value="1" <%if convertBool(customer.bCustom404) then response.write "checked=""checked"""%> />
<span style="float:right"><input type="submit" value="Save" class="art-button" /></span></td>
</tr>

<tr><td class=QSlabel>[PAGETITLE]:</td><td><input type="text" name="sCustom404Title" size="70" maxlength="255" value="<%=sanitize(customer.sCustom404Title)%>" /></td></tr>
<tr><td class=QSlabel>Message:<br />[PAGEBODY]<br /></td><td><%=dumpFCKInstance(customer.sCustom404Body,"siteBuilderMailSource","sCustom404Body")%></textarea><br />tip: you can use <input type=text size=14 value="[404FILENAME]" onclick="this.select();" /> to refer to the missing file</td></tr>
<tr><td class=QSlabel>Template:</td><td><select name=i404TemplateID><option value=0>Please select</option><%=customer.showSelectedtemplate("option", customer.i404TemplateID)%></select></td></tr>


<tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit value="<%=l("save")%>" name="btnaction" /></td></tr></table>
</form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
