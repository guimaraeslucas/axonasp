<!-- #include file="begin.asp"-->


<%bNOPopup=true%><!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bpopup%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_popup)%><%dim popup,popupModeList
set popupModeList=new cls_popupModeList
set popup=new cls_popup
dim postBack
postBack=convertBool(Request.Form("postBack"))
if request.form("btnaction")=l("delete") then
popup.remove
Response.Redirect ("bs_popupList.asp")
end if
if postBack then
popup.getRequestValues()
checkCSRF()
if popup.save then 
Response.Redirect ("bs_popupList.asp")
end if
end if%><!-- #include file="bs_popupBack.asp"--><form action="bs_popupEdit.asp" method="post" name="mainform"><input type=hidden name=ipopupId value="<%=encrypt(popup.iID)%>" /><input type=hidden name=postBack value="<%=true%>" /><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td colspan="2" class="header"><%=l("general")%>:</td></tr><tr><td class="QSlabel">Name:*</td><td colspan="2"><input type=text size=80 maxlength=50 name="sName" value="<%=quotRep(popup.sName)%>" /></td></tr><tr><td colspan="2">You have to specify either an url, either write a text to show in the popup.</td></tr><tr><td class="QSlabel">URL:</td><td colspan="2"><input type=text size=80 maxlength=255 name="sUrl" value="<%=quotRep(popup.sUrl)%>" /></td></tr><tr><td class="QSlabel">Text:</td><td colspan="2"><%createFCKInstance popup.sValue,"siteBuilder","sValue"%>	</td></tr><tr><td class="QSlabel">Enabled</td><td><input type="checkbox" value="<%=true%>" name="bEnabled" <%if convertBool(popup.bEnabled) then Response.Write "checked"%> /></td></tr><tr><td class="QSlabel">Mode</td><td><select name="iMode"><%=popupModeList.showSelected("option",popup.iMode)%></select></td></tr><tr><td class=QSlabel><%=l("online")%>&nbsp;<%=l("from")%>:</td><td><input type="text" id="dOnlineFrom" name="dOnlineFrom" value="<%=convertEuroDate(popup.dOnlineFrom)%>" /><%=JQDatePicker("dOnlineFrom")%><i><%=l("untill")%>:</i> 
<input type="text" id="dOnlineUntil" name="dOnlineUntil" value="<%=convertEuroDate(popup.dOnlineUntil)%>" /><%=JQDatePicker("dOnlineUntil")%></td></tr><tr><td class="QSlabel">Width:*</td><td><select name="iWidth"><%=numberList(10,1000,1,popup.iWidth)%></select>&nbsp;px</td></tr><tr><td class="QSlabel">Height:*</td><td><select name="iHeight"><%=numberList(10,1000,1,popup.iHeight)%></select>&nbsp;px</td></tr><tr><td class="QSlabel">Auto-close after</td><td><select name="iAutoclose"><%=numberList(0,120,1,popup.iAutoclose)%></select>&nbsp;seconds (0: no auto-close)</td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name="btnaction" value="<% =l("save")%>" /> 
<%if isNumeriek(popup.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><%end if%></td></tr></table></form><!-- #include file="bs_popupBack.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
