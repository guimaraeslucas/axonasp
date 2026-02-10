<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bPoll%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_poll)%><%dim copyPoll
set copyPoll=new cls_poll
copyPoll.copy()
if convertGetal(copyPoll.iId)<>0 then Response.Redirect ("bs_pollEdit.asp?iPollID=" & encrypt(copyPoll.iId))
dim polls
set polls=customer.polls%><p align=center><%=getArtLink("bs_pollEdit.asp",l("newpoll"),"","","")%></p><%if polls.count>0 then%><table align=center cellpadding=3 cellspacing=0><%dim pollKey, poll
for each pollKey in polls%><tr><td style="border-top:1px solid #DDD"><a href="bs_pollEdit.asp?iPollID=<%=encrypt(pollkey)%>"><%=polls(pollkey).sQuestion%></a></td><td style="border-top:1px solid #DDD"><%=getIcon(l("copyItem"),"copyItem","bs_pollList.asp?"&QS_secCodeURL&"&amp;ipollID="&encrypt(pollkey),"javascript:return confirm('"& l("areyousuretocopy") &"');","copy"&pollkey)%></td><td style="border-top:1px solid #DDD"><%if not isLeeg(polls(pollkey).sCode) then%><input size="20" type="text" onclick="javascript:this.select();" value="[QS_POLL:<%=sanitize(polls(pollkey).sCode)%>]" /><%end if%>&nbsp;</td><td style="border-top:1px solid #DDD"><i>iID: <%=pollkey%></i></td></tr><%next%></table><%else%><p align=center><%=l("nopoll")%></p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
