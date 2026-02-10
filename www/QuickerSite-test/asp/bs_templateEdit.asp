<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bTemplates%><%application(QS_textDirection)=QS_ltr%><!-- #include file="includes/header.asp"--><%application(QS_textDirection)=""%><%dim template
set template=new cls_template
select case Request ("btnaction")
case l("save")
checkCSRF()
template.getRequestValues()
if template.save then 
message.Add("fb_saveOK")
end if
case l("delete")
checkCSRF()
template.remove
Response.Redirect ("bs_templateList.asp")
case l("preview")
checkCSRF()
template.getRequestValues()
if template.save then 
session("iTemplateID")=template.iId
Response.Redirect (C_DIRECTORY_QUICKERSITE & "/default.asp?iId=" & encrypt(getHomePage.iId))
end if
case else
if convertGetal(template.iId)=0 then
template.init()
end if
end select
dim bTemplatesCanBeRemoved
bTemplatesCanBeRemoved=true
if customer.templates.count<=1 or customer.defaultTemplate=template.iId then
bTemplatesCanBeRemoved=false
end if%><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_template)%><form action="bs_templateEdit.asp" method="post" name="mainform"><input type="hidden" name="itemplateId" value="<%=encrypt(template.iID)%>" /><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td colspan=2 class=header><%=l("general")%>:</td></tr><tr><td class=QSlabel><%=l("name")%>:*</td><td><input type=text size=40 maxlength=255 name="sName" value="<%=sanitize(template.sName)%>" /> <strong><a href="bs_variables.asp" class="QSPP">Common variables</a></strong></td></tr><tr><td class="QSlabel">HTML:*</td><td><textarea style="word-wrap:normal; overflow-x:auto; overflow-y:auto" name="sValue" rows="60" cols="110"><%=quotrep(template.sValue)%></textarea></td></tr><!--
<tr><td class="QSlabel">HTML Mobile:</td><td><textarea style="word-wrap:normal; overflow-x:auto; overflow-y:auto" name="sMobileValue" rows="13" cols="130"><%'=quotrep(template.sMobileValue)%></textarea></td></tr><tr><td class="QSlabel">HTML WAP:</td><td><textarea style="word-wrap:normal; overflow-x:auto; overflow-y:auto" name="sWAPValue" rows="13" cols="130"><%'=quotrep(template.sWAPValue)%></textarea></td></tr>--><tr><td class="QSlabel"></td><td><input type="checkbox" <%=convertChecked(template.bCompress)%> name="bCompress" value="<%=true%>" /> Use compressed version of template on site</td></tr><tr><td class="QSlabel">&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class="QSlabel">&nbsp;</td><td><input class="art-button" type="submit" onclick="javascript:document.mainform.target='_self';" value="<%=l("save")%>" name="btnaction" /><%if isNumeriek(template.iID) then
if bTemplatesCanBeRemoved then%><input class="art-button" type=submit name=btnaction onclick="javascript:document.mainform.target='_self';return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><%end if%><input class="art-button" type=submit name=btnaction onclick="javascript:document.mainform.target='_new'" value="<% =l("preview")%>" /><%end if%></td></tr></table></form><!-- #include file="bs_templateBack.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
