
<%set cal=new cls_calendar
cal.pick(request("iCalID"))
dim iLegend
if postback then
if convertBool(request.form("iLegend")) then
iLegend=" checked "
else
iLegend=""
end if
else
iLegend=" checked "
end if
dim nm
if request.form("months")<>"" and request.form("time")="0" then
nm=request.form("months")
else
nm="12"
end if
dim legend
if not isLeeg(iLegend) then
legend="legend=true&amp;"
else
legend="legend=false&amp;"
end if
dim timeParam
if not isLeeg(request.form("time")) then
timeParam="time=" & request.form("time")
end if
dim embedJSCode,embedIFcode,embedHTMLcode
embedJSCode="<script type=""text/javascript"" src=""" & C_DIRECTORY_QUICKERSITE & "/asp/bs_ac_view.asp?"&timeParam&"&amp;"&legend&"months="&nm&"&amp;mode=js&amp;iCalID=" & cal.iId & """></script>"
embedIFcode="<iframe frameborder=""0"" style=""width:100%;height:850px"" src=""" & customer.sQSUrl & "/asp/bs_ac_view.asp?"&timeParam&"&amp;months="&nm&"&amp;"&legend&"mode=if&amp;iCalID=" & cal.iId & """ />"
embedHTMLcode=C_DIRECTORY_QUICKERSITE & "/asp/bs_ac_view.asp?"&timeParam&"&amp;mode=html&amp;"&legend&"months="&nm&"&amp;iCalID=" & cal.iId%><p>Add calender <strong><%=quotrep(cal.sName)%></strong> to my website.</p><form action="bs_ac.asp" method="post" name="embedForm"><input type="hidden" value="<%=true%>" name="postback" /><input type="hidden" value="embedcode" name="calaction" /><input type="hidden" value="<%=cal.iId%>" name="iCalID" /><table><tr><td><select name="time" onchange="javascript:document.embedForm.submit();"><option value="">Select what exactly you want to display</option><option value="0" <%if request.form("time")="0" then response.write "selected=""selected"""%>>Show X-number of next months</option><option value="1" <%if request.form("time")="1" then response.write "selected=""selected"""%>>Current Calendar Year</option><option value="2" <%if request.form("time")="2" then response.write "selected=""selected"""%>>Next Calendar Year</option></select> <input onclick="javascript:document.embedForm.submit();" type="checkbox" name="iLegend" <%=iLegend%> value="1" /> Include Legend</td></tr><%if request.form("time")="0" then%><tr><td>Select the number of months to display:
<select name="months" onchange="javascript:document.embedForm.submit();"><option value=""></option><option value="1" <%if request.form("months")="1" then response.write "selected=""selected"""%>>1</option><option value="2" <%if request.form("months")="2" then response.write "selected=""selected"""%>>2</option><option value="3" <%if request.form("months")="3" then response.write "selected=""selected"""%>>3</option><option value="4" <%if request.form("months")="4" then response.write "selected=""selected"""%>>4</option><option value="5" <%if request.form("months")="5" then response.write "selected=""selected"""%>>5</option><option value="6" <%if request.form("months")="6" then response.write "selected=""selected"""%>>6</option><option value="7" <%if request.form("months")="7" then response.write "selected=""selected"""%>>7</option><option value="8" <%if request.form("months")="8" then response.write "selected=""selected"""%>>8</option><option value="9" <%if request.form("months")="9" then response.write "selected=""selected"""%>>9</option><option value="10" <%if request.form("months")="10" then response.write "selected=""selected"""%>>10</option><option value="11" <%if request.form("months")="11" then response.write "selected=""selected"""%>>11</option><option value="12" <%if request.form("months")="12" then response.write "selected=""selected"""%>>12</option><option value="18" <%if request.form("months")="18" then response.write "selected=""selected"""%>>18</option><option value="24" <%if request.form("months")="24" then response.write "selected=""selected"""%>>24</option></select></td></tr><%end if%></table></form><%if not isLeeg(request.form("time")) then
if request.form("time")<>"0" or (request.form("time")="0" and not isLeeg(request.form("months"))) then%><hr /><p><b>Copy/Paste any of the codes below in your current website</b></p><hr /><table><tr><td><b><u>Recommended use:</u></b> Copy/Paste the <b>JavaScript</b> below exactly where you want the calendar to be displayed in your website (<u>do not put it in the header</u>)</td></tr><tr><td><textarea onclick="javascript:this.select();" cols="95" rows="4"><%=quotrep(embedJSCode)%></textarea></td></tr></table><hr /><table><tr><td>Copy/Paste the code below to <u><b>link</b></u> (<a href="XXX"></a>) to the calendar from anywhere in your website</td></tr><tr><td><textarea onclick="javascript:this.select();" cols="95" rows="4"><%=quotrep(embedHTMLcode)%></textarea></td></tr></table><hr /><table><tr><td>Copy/Paste the code below to <u><b>iframe</b></u> the calendar (for including in other websites)</td></tr><tr><td><textarea onclick="javascript:this.select();" cols="95" rows="4"><%=quotrep(embedIFcode)%></textarea></td></tr></table><ul><li>TIP: You can set the number of months to be displayed in the <b>months-parameter</b> (default 12, max. 36)</li></ul><%end if
end if
set cal=nothing
set cv=nothing%>
