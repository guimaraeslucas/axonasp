<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetMail%><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_sentmessages)%><%dim theMail
set theMail=new cls_mail
if convertGetal(decrypt(Request.QueryString ("iMail_ID")))<>0 then
checkCSRF()
theMail.delete()
end if
set theMail=nothing
dim mails, mail
set mails=customer.mails%><%if mails.count>0 then%><p align=center><%=l("sentmessages")%>:</p><table align=center id=mails class="sortable" cellpadding="2"><tr><th><%=l("subject")%>:</th><th><%=l("sent")%>&nbsp;<%=l("on")%>:</th><th>N&deg; <%=l("contacts")%></th><th>&nbsp;</th></tr><%for each mail in mails%><tr><td><a href="#" onclick="javascript: openPopUpWindow('mailRead','bs_mailDetail.asp?iMail_id=<%=encrypt(mail)%>',670,500)"><%=mails(mail).sSubject%></a></td><td align=center><%=convertEuroDate(mails(mail).dDateSent)%></td><td align=center><a href="#" onclick="javascript: openPopUpWindow('mailReceivers','bs_mailReceivers.asp?iMail_id=<%=encrypt(mail)%>',390,700)"><%=mails(mail).iNumberRec%></a></td><td align=center><a href="#" onclick="javascript:if(confirm('<%=l("areyousure")%>')){location.assign('bs_mailHistory.asp?<%=QS_secCodeURL%>&amp;iMail_ID=<%=encrypt(mail)%>')}"><img border=0 src='<%=C_DIRECTORY_QUICKERSITE%>/fixedImages/dustbin.gif' alt="<%=l("delete")%>" /></a></td></tr><%next
set mails=nothing%></table><%else%><p align=center><%=l("nomessagessent")%></p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
