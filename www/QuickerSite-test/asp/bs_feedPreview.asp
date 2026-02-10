<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bFeed%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_feed)%><table align=center><tr><td align=center>-> <b><a href="bs_feedList.asp"><%=l("back")%></a></b> <-</td></tr></table><br /><br /><table width=600 align=center><tr><td><%dim feed
set feed=new cls_feed
Response.write feed.build()
set feed=nothing%></td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
