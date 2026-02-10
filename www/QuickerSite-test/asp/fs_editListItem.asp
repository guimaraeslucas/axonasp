<!-- #include file="begin.asp"-->


<%pagetoemail=true
editPage=true
dim EPSuccess, EPSaveSuccess
EPSuccess=false
EPSaveSuccess=false
dim accessToEP
accessToEP=false
dim getLperP, iLPID
set getLperP=logon.contact.getLper
if not isLeeg(Request("iListPageID")) then
iLPID	= convertGetal(decrypt(Request("iListPageID")))
selectedPage.iListPageID	= iLPID
else
iLPID=convertGetal(selectedPage.iListPageID)
end if
if not getLperP.exists(iLPID) then Response.Redirect (C_DIRECTORY_QUICKERSITE & "/asp/noaccess.htm")
'save
if Request.Form("btnaction")=save_listitem then
selectedPage.sTitle	= convertStr(Request.Form ("sTitle"))
selectedPage.dPage	= convertDateFromPicker(Request.Form ("dPage"))
selectedPage.dOnlineFrom	= convertDateFromPicker(Request.Form ("dOnlineFrom"))
selectedPage.dOnlineUntill	= convertDateFromPicker(Request.Form ("dOnlineUntill"))
selectedPage.sLPExternalURL	= convertStr(Request.Form ("sLPExternalURL"))
selectedPage.bLPExternalOINW	= convertBool(Request.Form ("bLPExternalOINW"))
selectedPage.sValue	= convertStr(Request.Form ("sValue"))
selectedPage.sUserFriendlyURL	= convertStr(Request.Form ("sUserFriendlyURL"))
selectedPage.iListPageID	= convertGetal(decrypt(Request.Form("iListPageID")))
selectedPage.bLossePagina	= false
selectedPage.bOnline	= false
if Request.Form ("actionBtn")=l("delete") then
selectedPage.bDeleted=true
selectedPage.parentPage.removeRang(selectedPage)
selectedPage.iRang=0
selectedPage.bOnline=false
selectedPage.bHomepage=false
selectedPage.deleteListItems()
end if
if selectedPage.save then
Response.Redirect ("fs_editListItems.asp?iID="& EnCrypt(selectedPage.iListPageID))
end if
end if%><!-- #include file="includes/commonheader.asp"-->
<body style="color:#000;background-color:#FFF;background-image:none" onload="javascript:window.focus();">
<form action="fs_editListItem.asp" method="post" name=mainform><%=QS_secCodeHidden%>
<input type="hidden" value="<% =save_listitem%>" name=btnaction />
<INPUT type="hidden" value="<% =EnCrypt(selectedPage.iId) %>" name=iId />
<INPUT type="hidden" value="<% =EnCrypt(selectedPage.iListPageID) %>" name=iListPageID />
<table align=center  style="background-color:#FFF"><tr><td colspan=2><%=l("expllistitem")%></td></tr><tr><td class=QSlabel><%=l("list")%>:</td><td><b><%=selectedPage.listPage.sTitle%></b></td></tr><tr><td class=QSlabel><%=l("title")%>&nbsp;<%=l("item")%>:*</td><td><input required type=text maxlength=255 size=45 name=sTitle value="<%=quotRep(selectedPage.sTitle)%>" /></td></tr><tr><td class=QSlabel><%=l("date")%>:</td><td><input type="text" id="dPage" name="dPage" value="<%=convertEuroDate(selectedPage.dPage)%>" /><%=JQDatePicker("dPage")%></td></tr><tr><td class=QSlabel><%=l("online")%>&nbsp;<%=l("from")%>:</td><td><input type="text" id="dOnlineFrom" name="dOnlineFrom" value="<%=convertEuroDate(selectedPage.dOnlineFrom)%>" /><%=JQDatePicker("dOnlineFrom")%><i><%=l("untill")%>:</i> 
<input type="text" id="dOnlineUntill" name="dOnlineUntill" value="<%=convertEuroDate(selectedPage.dOnlineUntill)%>" /><%=JQDatePicker("dOnlineUntill")%></td></tr>
<%
if convertBool(customer.bUserFriendlyURL) then
%>
<tr><td class="QSlabel"><%=l("userfriendlyurl")%>:</td><td><%=customer.sVDUrl%>/<input type="text" value="<%=quotrep(selectedPage.sUserFriendlyURL)%>" size="40" maxlength="49" name="sUserFriendlyURL" /></td></tr>
<%
end if
%>

<%if ((trim(convertStr(selectedPage.sLPExternalURL))<>"http://" or convertGetal(selectedPage.iId)=0) and isLeeg(selectedPage.iFeedID) and isLeeg(selectedPage.sValue)) or (not isLeeg(selectedPage.sValue) and selectedPage.sLPExternalURL<>"http://" and not isLeeg(selectedPage.sLPExternalURL)) then%><tr><td class=QSlabel><%=l("externalurl")%>:</td><td><input type=text maxlength=255 size=45 name="sLPExternalURL" value="<%=quotRep(selectedPage.sLPExternalURL)%>" /></td></tr><tr><td class=QSlabel valign=top><%=l("openinnewwindow")%></td><td><input name=bLPExternalOINW type=checkbox value=1 <%if selectedPage.bLPExternalOINW then Response.Write "checked"%> /><%if selectedPage.sLPExternalURL="http://" then%><br />-> <%=l("externalurlLP")%><%end if%></td></tr><%end if
if not isLeeg(selectedPage.sValue) or convertGetal(selectedPage.iId)=0 or not isLeeg(selectedPage.iFeedID) or isLeeg(selectedPage.sLPExternalURL)  then%><tr><td align=center colspan=2><%createFCKInstance selectedPage.sValue,"siteBuilder","sValue"%></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=dummy value="<% =l("save")%>" />&nbsp;<input class="art-button" type=reset  value="<% =l("reset")%>" id=reset1 name=reset1 />&nbsp;<%if convertGetal(selectedPage.iId)<>0 then%><input class="art-button" type=submit name=actionBtn value="<% =l("delete")%>" onclick="javascript:return confirm('<%=l("areyousure")%>');" /><%end if%></td></tr></table></form><%=message.showAlert()%><%=cPopup.dumpJS()%> 
</body></html><%cleanUPASP%>
