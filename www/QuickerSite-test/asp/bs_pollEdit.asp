<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bPoll%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><%=getBOHeader(btn_poll)%><%dim poll
set poll=new cls_poll
poll.bCanVote=false
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
poll.getRequestValues()
end if
select case Request.Form ("btnaction")
case l("save")
checkCSRF()
poll.getRequestValues()
if poll.save then 
Response.Redirect ("bs_pollList.asp")
end if
case l("delete")
checkCSRF()
poll.remove
Response.Redirect ("bs_pollList.asp")
case l("reset")
checkCSRF()
poll.reset
Response.Redirect ("bs_pollList.asp")
end select%><!-- #include file="bs_pollBack.asp"--><form action="bs_pollEdit.asp" method="post" name="mainform"><input type=hidden name=iPollId value="<%=encrypt(poll.iID)%>" /><input type="hidden" value="<%=true%>" name=postback /><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td colspan="2" class="header"><%=l("general")%>:</td></tr><tr><td class="QSlabel">Poll:*</td><td colspan="2"><input type=text size=50 maxlength=255 name="sQuestion" value="<%=quotRep(poll.sQuestion)%>" /></td></tr><tr><td class="QSlabel"><%=l("showpolltitle")%></td><td><input type="checkbox" value="<%=true%>" name="bShowTitle" <%if poll.bShowTitle then Response.Write "checked"%> /></td></tr><tr><td class=QSlabel><%=l("votingdeadline")%></td><td><input type="text" id="dVoteDeadline" name="dVoteDeadline" value="<%=convertEuroDate(poll.dVoteDeadline)%>" /><%=JQDatePicker("dVoteDeadline")%></td></tr><tr><td class=QSlabel><%=l("online")%>&nbsp;<%=l("from")%>:</td><td><input type="text" id="dVoteFrom" name="dVoteFrom" value="<%=convertEuroDate(poll.dVoteFrom)%>" /><%=JQDatePicker("dVoteFrom")%><i><%=l("untill")%>:</i> 
<input type="text" id="dVoteUntil" name="dVoteUntil" value="<%=convertEuroDate(poll.dVoteUntil)%>" /><%=JQDatePicker("dVoteUntil")%></td></tr><tr><td class="QSlabel"><%=l("code")%>:*</td><td>[QS_POLL:<input type="text" size="10" maxlength="45" name="sCode" value="<%=quotRep(poll.sCode)%>" />]</td></tr><tr><td class="QSlabel">Label "View Results":*</td><td colspan="2"><input type=text size=50 maxlength=255 name="label_viewresults" value="<%=quotRep(poll.label_viewresults)%>" /></td></tr><tr><td class="QSlabel">Label "Vote Now!":*</td><td colspan="2"><input type=text size=50 maxlength=255 name="label_votenow" value="<%=quotRep(poll.label_votenow)%>" /></td></tr><tr><td class="QSlabel">Label "Number of votes: [XX]":*</td><td colspan="2"><input type=text size=50 maxlength=255 name="label_numberofvotes" value="<%=quotRep(poll.label_numberofvotes)%>" /></td></tr><tr><td colspan="3"><hr /></td></tr><%dim ip
for ip=1 to 12%><tr><td class="QSlabel">Option <b><%=ip%></b>:</td><td><input type="text" size="40" maxlength="45" name="sA<%=ip%>" value="<%=quotRep(poll.getQuestion(ip))%>" /> <input type="text" id="sA<%=ip%>c" name="sA<%=ip%>c" value="<%=quotrep(poll.getQuestionColor(ip))%>" /></td></tr><%next%><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name="btnaction" value="<% =l("save")%>" /><%if isNumeriek(poll.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><br /><br /><%end if%></td></tr></table><%if convertGetal(poll.iId)<>0 then%><table align="center" cellpadding="2"><tr><td valign="top"><%=poll.showresults%></td></tr><tr><td><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("reset")%>" /></td></tr></table><%end if%></form><%for ip=1 to 12
response.write JQColorPicker("sA" & ip & "c")
next%><%Response.Flush 
if convertGetal(poll.iId)<>0 then
dim fsearch
set fsearch=new cls_fullSearch
fsearch.pattern="\[QS_POLL:+("&poll.sCode&")+[\]]"
dim result
result=fsearch.search()
if not isLeeg(result) then
Response.Write "<p align=center>"&l("wherepollused")&"</p>"
Response.Write result
end if
set fsearch=nothing
end if%><!-- #include file="bs_pollBack.asp"--><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
