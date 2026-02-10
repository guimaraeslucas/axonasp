<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bGuestbook%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_gb)%><%dim copyGB
set copyGB=new cls_guestbook
copyGB.copy()
if convertGetal(copyGB.iId)<>0 then Response.Redirect ("bs_gbEdit.asp?iGBID=" & encrypt(copyGB.iId))
dim guestbooks
set guestbooks=customer.guestbooks%><p align=center><%=getArtLink("bs_gbEdit.asp",l("newguestbook"),"","","")%></p><%if guestbooks.count>0 then%><table align=center cellpadding=3 cellspacing=0><%dim gbKey, poll
for each gbKey in guestbooks%><tr><td style="border-top:1px solid #DDD"><a href="bs_gbEdit.asp?iGBID=<%=encrypt(gbKey)%>"><%=guestbooks(gbKey).sName%></a></td><td style="border-top:1px solid #DDD"><%=getIcon(l("copyItem"),"copyItem","bs_gbList.asp?"&QS_secCodeURL&"&amp;iGBID="&encrypt(gbKey),"javascript:return confirm('"& l("areyousuretocopy") &"');","copy"&gbKey)%></td><td style="border-top:1px solid #DDD"><%=getIcon("Manage Entries","edit","bs_gbEditItems.asp?"&QS_secCodeURL&"&amp;iGBID="&encrypt(gbKey),"","edit"&gbKey)%></td><td style="border-top:1px solid #DDD"><%=getIcon("Excel","excel","#","javascript:window.open('bs_gbExcel.asp?iGBID="& encrypt(gbKey) &"');","excel"&gbKey)%></td><td style="border-top:1px solid #DDD"><%if not isLeeg(guestbooks(gbKey).sCode) then%><input size="20" type="text" onclick="javascript:this.select();" value="[QS_GUESTBOOK:<%=sanitize(guestbooks(gbKey).sCode)%>]" /><%end if%>&nbsp;</td><td style="border-top:1px solid #DDD"><i>iID: <%=gbKey%></i></td></tr><%next%></table><%else%><p align=center><%=l("noguestbooks")%></p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
