<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetMail%><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_sentmessages)%><p align=center><%=request.querystring("counter")%>&nbsp;<%=l("messagessent")%></p><p align=center>-> <a href="bs_mailHistory.asp"><%=l("sentmessages")%></a> <-</p><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
