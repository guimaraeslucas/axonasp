<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Newsletter)%><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterList.asp"><%=l("back")%></a></b> <-</td></tr></table><%dim rs,newsletterMailing
set newsletterMailing=new cls_newsletterMailing
if request.querystring("delete")="1" then
newsletterMailing.remove()
end if
set newsletterMailing=nothing
set rs=db.execute("select iId from tblNewsletterMailing where iCustomerID=" & cId & " order by dSentDate desc")
if not rs.eof then%><table align=center cellpadding="10" cellspacing="0"><tr><td><b>Date sent</b></td><td><b>Newsletter</b></td><td><b>Email list</b></td><td align="center"><b>Total</b></td><td align="center"><b>Read</b></td><td align="center"><b>Not read</b></td></tr><%else%><table align=center cellpadding="10" cellspacing="0"><tr><td>No mailings..</td></tr><%end if
while not rs.eof
set newsletterMailing=new cls_newsletterMailing
newsletterMailing.pick(rs(0))
dim percRead,percNotRead
if newsletterMailing.nmbrReceivers(0)<>0 then
percRead=round((newsletterMailing.nmbrReceivers(1)/newsletterMailing.nmbrReceivers(0))*100,0)
percNotRead=round((newsletterMailing.nmbrReceivers(2)/newsletterMailing.nmbrReceivers(0))*100,0)
end if%><tr><td valign="top" style="border-top:1px solid #DDD"><a name="<%=encrypt(newsletterMailing.iId)%>"></a><%=convertEuroDateTime(newsletterMailing.dSentDate)%></td><td valign="top" style="border-top:1px solid #DDD"><%=newsletterMailing.newsletter.sName%></td><td valign="top" style="border-top:1px solid #DDD"><%=newsletterMailing.category.sName%></td><td valign="top" style="border-top:1px solid #DDD" align="center"><%=newsletterMailing.nmbrReceivers(0)%> <b><a href="bs_NewsletterMailingHistory.asp?t=0&amp;showRead=<%=encrypt(newsletterMailing.iId)%>#<%=encrypt(newsletterMailing.iId)%>">?</a></b><%if convertGetal(decrypt(request.querystring("showRead")))=newsletterMailing.iId and request.querystring("t")="0" then  newsletterMailing.showRead(request.querystring("t"))%></td><td valign="top" style="border-top:1px solid #DDD" align="center"><%=newsletterMailing.nmbrReceivers(1)%> (<%=convertGetal(percRead)%>%) <b><a href="bs_NewsletterMailingHistory.asp?t=1&amp;showRead=<%=encrypt(newsletterMailing.iId)%>#<%=encrypt(newsletterMailing.iId)%>">?</a></b><%if convertGetal(decrypt(request.querystring("showRead")))=newsletterMailing.iId and request.querystring("t")="1" then  newsletterMailing.showRead(request.querystring("t"))%></td><td valign="top" style="border-top:1px solid #DDD" align="center"><%=newsletterMailing.nmbrReceivers(2)%> (<%=convertGetal(percNotRead)%>%) <b><a href="bs_NewsletterMailingHistory.asp?t=2&amp;showRead=<%=encrypt(newsletterMailing.iId)%>#<%=encrypt(newsletterMailing.iId)%>">?</a></b><%if convertGetal(decrypt(request.querystring("showRead")))=newsletterMailing.iId and request.querystring("t")="2" then  newsletterMailing.showRead(request.querystring("t"))%></td><td valign="top" style="border-top:1px solid #DDD" align="center"><a onclick="javascript:return confirm('Are you sure to PERMANENTLY remove this mailing?');" href="bs_newsletterMailingHistory.asp?delete=1&amp;iNewsletterMailingID=<%=encrypt(newsletterMailing.iId)%>"><%=l("delete")%></a></td></tr><%set newsletterMailing=nothing
rs.movenext
wend
set rs=nothing%></table><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterList.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
