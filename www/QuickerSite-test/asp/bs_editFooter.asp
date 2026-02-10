<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bSetupPageElements%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_pageelements)%><%if request("btnaction")=l("save") then
checkCSRF()
customer.sFooter=convertStr(Request.Form ("sFooter"))
if customer.save then message.Add ("fb_saveOK")
end if%><form action="bs_editFooter.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" value="<% =l("save")%>" name=btnaction /><table align=center style="height:80%" width=640 cellpadding="2"><tr><td><b><%=l("footer")%>:</b></td></tr><tr><td><%createFCKInstance customer.sFooter,"siteBuilderFooter","sFooter"%></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
