<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader("")%><!-- #include file="bs_processPage.asp"--><form action="bs_editItem.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" value="<% =l("save")%>" name=btnaction /><INPUT type="hidden" value="<% =EnCrypt(page.iId) %>" name=iId /><INPUT type="hidden" value="<% =EnCrypt(page.iParentID) %>" name=iParentID /><INPUT type="hidden" value="<% =page.bLossePagina%>" name=bLossePagina /><table cellpadding=2 align=center style="height:80%" width=640><tr><td colspan=2 class=header><%=l("itemwithcontent")%>:</td></tr><!-- #include file="bs_common.asp"--><!-- #include file="bs_sortorder.asp"--><!-- #include file="bs_onlineOrNot.asp"--><%if page.parentPage.bOnline then
if not page.bHomepage then
if secondAdmin.bPageSetHomepage then%><tr><td class=QSlabel><%=l("homepage")%></td><td><input type="checkbox" style="BORDER:0px" name=bHomePage value="1" /></td></tr><%end if
else%><tr style="height:1px"><td style="height:1px"><INPUT type="hidden" value="1" name=bOnline /><INPUT type="hidden" value="1" name=bHomepage /></td></tr><%end if
end if%>


<tr><td align=center colspan=2>

<%if secondAdmin.bPageBody then%><%createFCKInstance page.sValue,"siteBuilder","sValue"%><%end if%>


</td></tr><tr><td style='height:7px'></td></tr><!-- #include file="bs_url.asp"--><!-- #include file="bs_feed.asp"--><!-- #include file="bs_catalog.asp"--><!-- #include file="bs_form.asp"--><!-- #include file="bs_application.asp"--><!-- #include file="bs_theme.asp"--><!-- #include file="bs_redirectTO.asp"--><!-- #include file="bs_template.asp"--><!-- #include file="bs_metaTags.asp"--><%if secondAdmin.bPageUrlRSS then%><tr><td class=QSlabel>url RSS:</td><td><input type="text" name="sRSSLink" maxlength="250" value="<%=quotrep(page.sRSSLink)%>" size="60" /></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type="submit" name="dummy" value="<% =l("save")%>" />&nbsp;<input class="art-button" type="reset"  value="<% =l("reset")%>" /><!-- #include file="bs_deleteButton.asp"--></td></tr></table></form><!-- #include file="bs_editPageBlocksInc.asp"--><!-- #include file="bs_addnewitem.asp"--><!-- #include file="bs_convertToContainerItem.asp"--><!-- #include file="bs_convertToListpage.asp"--><!-- #include file="bs_convertToExternalURL.asp"--><!-- #include file="bs_convertToFreePage.asp"--><!-- #include file="bs_back.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
