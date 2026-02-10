<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bSetupPageElements%><%if request("btnaction")=l("save") then
checkCSRF()
customer.sLeftBanner	= removeEmptyP(convertStr(Request.Form ("sLeftBanner")))
customer.sBannerMenu	= removeEmptyP(convertStr(Request.Form ("sBannerMenu")))
customer.sRightBanner	= removeEmptyP(convertStr(Request.Form ("sRightBanner")))
customer.sHighlights	= convertStr(Request.Form ("sHighlights"))
customer.sContactInfo	= convertStr(Request.Form ("sContactInfo"))
if secondAdmin.bApplicationpath then customer.bannerApplication	= convertStr(Request.Form ("bannerApplication"))
if customer.save then message.Add ("fb_saveOK")
end if%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Setup)%><%=getBOSetupMenu(btn_pageelements)%><form action="bs_editBannerMenu.asp" method="post" name=mainform><%=QS_secCodeHidden%><input type="hidden" value="<% =l("save")%>" name=btnaction /><table align=center style="height:500px" width=640 cellpadding="2"><%if convertGetal(customer.defaultTemplate)=0 then%><tr><td class=header><%=l("leftbanner")%>:</td><td class=header><%=l("bannerunderleftmenu")%>:</td><td class=header><%=l("rightbanner")%>:</td></tr><tr><td><%createFCKInstance customer.sLeftBanner,"siteBuilderBanner","sLeftBanner"%></td><td><%createFCKInstance customer.sBannerMenu,"siteBuilderBanner","sBannerMenu"%></td><td><%createFCKInstance customer.sRightBanner,"siteBuilderBanner","sRightBanner"%></td></tr><%if customer.bApplication and secondAdmin.bApplicationpath then%><tr><td class=QSlabel style="text-align:right"><%=l("applicationpath")%>:</td><td><input type=text size=33 maxlength=255 value="<%=quotRep(customer.bannerApplication)%>" name="bannerApplication" /></td><td>&nbsp;</td></tr><%end if
else%><tr><td valign=top class=header><input type="text" size="40" maxlength="50" value="<%=sanitize(customer.sContactInfo)%>" name="sContactInfo" /></td><td valign=top  class=header><input type="text" size="40" maxlength="50" value="<%=sanitize(customer.sHighlights)%>" name="sHighlights" /></td></tr><tr><td valign=top ><%createFCKInstance customer.sBannerMenu,"siteBuilderBanner","sBannerMenu"%></td><td valign=top ><%createFCKInstance customer.sRightBanner,"siteBuilderBanner","sRightBanner"%></td></tr><%end if%></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
