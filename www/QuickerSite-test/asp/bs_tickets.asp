<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetContacts%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_contacts)%><%if Request.Form ("pageAction")="remove" then
checkCSRF()
dim iTicketID, t
for each iTicketID in Request.Form ("iTicketID")
set t=new cls_ticket
t.pick(decrypt(iTicketID))
t.remove
set t=nothing
next
end if
if Request.QueryString ("pageAction")="resend" then
checkCSRF()
dim resendTicket
set resendTicket=new cls_ticket
resendTicket.sendTicket()
set resendTicket=nothing
Response.Redirect ("bs_tickets.asp?fbMessage=fb_activationlinkresend")
end if
dim tickets
set tickets=customer.tickets%><p align=center><b><%=l("pendingactivationlinks")%></b></p><p align=center><a href="bs_contacthome.asp"><b><%=l("back")%></b></a></p><%if tickets.count>0 then%><form method="post" action="bs_tickets.asp" name="mainform"><%=QS_secCodeHidden%><input type="hidden" name="pageAction" value="remove" /><table align="center" class="sortable" id="main" cellpadding="2"><tr><th><%=l("createdon")%></th><th><%=l("email")%></th><th>Log</th><th><%=l("delete")%></th><th><%=l("resend")%></th></tr><%dim ticketKey, ticket
for each ticketKey in tickets%><tr><td valign=top><%=tickets(ticketKey).dCreatedTS%></td><td valign=top><%=clickEmail(tickets(ticketKey).sEmail)%></td><td valign=top width=330><%=splitby(tickets(ticketKey).sVisitorDetails,35)%></td><td valign=top><input type=checkbox name=iTicketID value="<%=encrypt(ticketKey)%>" /></td><td valign=top><a onclick="javascript:return confirm('<%=l("areyousure")%>');" href="bs_tickets.asp?<%=QS_secCodeURL%>&amp;pageAction=resend&amp;iTicketID=<%=encrypt(ticketKey)%>"><%=l("resend")%></a></td></tr><%next%></table></form><p align=center><a href="#" onclick="javascript:if (confirm('<%=l("areyousure")%>')) {document.mainform.submit();}"><%=l("removeselecteditems")%></a></p><%else%><p align=center><%=l("noresults")%></p><%end if%><p align=center><a href="bs_contacthome.asp"><b><%=l("back")%></b></a></p><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
