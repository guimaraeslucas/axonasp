<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bTheme%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_theme)%><%dim copyTheme
set copyTheme=new cls_theme
copyTheme.copy()
if convertGetal(copyTheme.iId)<>0 then Response.Redirect ("bs_themeEdit.asp?ithemeID="&encrypt(copyTheme.iId))
dim themes
set themes=customer.themes%><p align=center><a href="bs_themeEdit.asp"><b><%=l("newtheme")%></b></a></p><%if themes.count>0 then%><table align=center cellpadding=3 cellspacing=0><%dim themeKey, theme
for each themeKey in themes%><tr><td style="border-top:1px solid #DDD"><a href="bs_themeEdit.asp?iThemeID=<%=encrypt(themeKey)%>"><%=themes(themeKey).sName%></a></td><td style="border-top:1px solid #DDD"><%=getIcon(l("copyItem"),"copyItem","bs_themesList.asp?"&QS_secCodeURL&"&amp;ithemeID="&encrypt(themekey),"javascript:return confirm('"& l("areyousuretocopy") &"');","copy"&themekey)%></td><td style="border-top:1px solid #DDD"><%if not isLeeg(themes(themekey).sCode) then%><input size="20" type="text" onclick="javascript:this.select();" value="[QS_theme:<%=sanitize(themes(themekey).sCode)%>]" id=text1 name=text1 /><%end if%>&nbsp;</td><td style="border-top:1px solid #DDD"><i>iID: <%=themeKey%></i></td></tr><%next%></table><%else%><p align=center><%=l("notheme")%></p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
