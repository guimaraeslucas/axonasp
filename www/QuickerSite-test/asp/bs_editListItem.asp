<!-- #include file="begin.asp"-->
<!-- #include file="bs_security.asp"-->
<!-- #include file="bs_process.asp"-->
<!-- #include file="includes/header.asp"-->
<!-- #include file="bs_initBack.asp"-->
<!-- #include file="bs_header.asp"-->
<%=getBOHeader("")%><!-- #include file="bs_processPage.asp"-->
<form action="bs_editListItem.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" value="<% =save_listitem%>" name=btnaction /><INPUT type="hidden" value="<% =EnCrypt(page.iId) %>" name=iId />
<INPUT type="hidden" value="<% =EnCrypt(page.iListPageID) %>" name=iListPageID /><INPUT type="hidden" value="<%=false%>" name=bLossePagina />
<table align=center style="height:80%" width=640 cellpadding="2"><tr><td colspan=2><%=l("expllistitem")%></td></tr>
<tr><td class=QSlabel><%=l("list")%>:</td><td><b><%=page.listPage.sTitle%></b></td></tr>
<tr><td valign=middle class=QSlabel><%=l("title")%>&nbsp;<%=l("item")%>:*</td><td valign=middle><input type=text maxlength=255 size=45 name=sTitle value="<%=quotRep(page.sTitle)%>" />
<!-- #include file="bs_preview.asp"--></td></tr>
<tr><td class=QSlabel><%=l("date")%>:</td><td><input type="text" id="dPage" name="dPage" value="<%=convertEuroDate(page.dPage)%>" /><%=JQDatePicker("dPage")%></td></tr>
<tr><td class=QSlabel><%=l("online")%>&nbsp;<%=l("from")%>:</td><td><input type="text" id="dOnlineFrom" name="dOnlineFrom" value="<%=convertEuroDate(page.dOnlineFrom)%>" /><%=JQDatePicker("dOnlineFrom")%><i><%=l("untill")%>:</i> 
<input type="text" id="dOnlineUntill" name="dOnlineUntill" value="<%=convertEuroDate(page.dOnlineUntill)%>" /><%=JQDatePicker("dOnlineUntill")%></td></tr>
<%if ((trim(convertStr(page.sLPExternalURL))<>"http://" or convertGetal(page.iId)=0) and isLeeg(page.iFeedID) and isLeeg(page.sValue)) or (not isLeeg(page.sValue) and page.sLPExternalURL<>"http://" and not isLeeg(page.sLPExternalURL)) then%>
<tr><td class=QSlabel><%=l("externalurl")%>:</td><td><input type=text maxlength=255 size=45 name="sLPExternalURL" value="<%=quotRep(page.sLPExternalURL)%>"></td></tr>
<tr><td class=QSlabel valign=top><%=l("openinnewwindow")%></td><td><input name=bLPExternalOINW type=checkbox value=1 <%if page.bLPExternalOINW then Response.Write "checked"%> /><%if page.sLPExternalURL="http://" then%><br />-> <%=l("externalurlLP")%><%end if%></td></tr><%end if
if not isLeeg(page.sValue) or convertGetal(page.iId)=0 or not isLeeg(page.iFeedID) or isLeeg(page.sLPExternalURL)  then%>
<!-- #include file="bs_feed.asp"-->

<tr><td align=center colspan=2><%createFCKInstance page.sValue,"siteBuilder","sValue"%></td></tr>
<%if convertBool(customer.bUserFriendlyURL) and secondAdmin.bPageUFL then%>
<tr><td class="QSlabel"><%=l("userfriendlyurl")%>:</td><td><%=customer.sVDUrl%>/<input type="text" value="<%=quotrep(page.sUserFriendlyURL)%>" size="40" maxlength="49" name="sUserFriendlyURL" /></td></tr>
<%end if%>


<%end if%>

<tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr>
<tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name=dummy value="<% =l("save")%>" />&nbsp;<input  class="art-button" type=reset  value="<% =l("reset")%>" id=reset1 name=reset1 />
<!-- #include file="bs_deleteButtonList.asp"--></td></tr></table></form>
<!-- #include file="bs_convertToItemWithContent.asp"-->
<table align=center><tr><td align=center>-> <b><a href="bs_listPage.asp?iID=<%=encrypt(page.iListPageID)%>"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
