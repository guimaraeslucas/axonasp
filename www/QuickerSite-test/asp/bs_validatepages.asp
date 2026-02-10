<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><!-- #include file="bs_process.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Home)%><%dim pv,p,pobj
set pv=customer.pagesTobeValidated%><table align="center" cellpadding="2"><tr><td><%if pv.count>0 then%><ul><%for each p in pv
set pobj=new cls_page
pobj.pick(p)%><li><a href="bs_validatepage.asp?iID=<%=encrypt(p)%>"><%=pobj.sTitle%></a> <small>(updated by <%=pobj.updater.sNickName%> on <%=convertEuroDateTime(pobj.dUpdatedOn)%>)</small></li><%set pobj=nothing
next%></ul><%else %>No pages to validate.
<%end if%></td></tr></table><table align=center><tr><td align=center>-> <b><a href="bs_default.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
