
<%set cal=new cls_calendar
cal.pick(request("iCalID"))
dim booking
set booking=new cls_calendarbooking
booking.pick(request("iBookingID"))
if request.querystring("seeNotes")<>"" then
response.clear
response.write "<div style=""padding:20px;background-color:#FFF;color:#000""><h4>" & booking.sName & "</h4>" & booking.sNotes & "</div><div style=""text-align:center;padding:5px;background-color:#555;color:#FFF""><a href=""#"" style=""color:#EEE"" onclick=""parent.location.assign('bs_ac.asp?iBookingID=" & booking.iId & "&amp;calAction=booking&iCalid=" & cal.iId &"')"">Modify</a></div>"
response.end
end if
if postback then
select case request.form("btAct")
case "Save"
booking.getRequestValues()
if booking.save then
response.redirect("bs_ac.asp?iCalID="&cal.iId&"&calAction=bookings")
end if
case "Delete"
booking.delete
response.redirect("bs_ac.asp?iCalID="&cal.iId&"&calAction=bookings")
end select
end if
set statuslist=new cls_statusList%><p>Reservation for <strong><%=quotrep(cal.sName)%></strong></p><form action="bs_ac.asp" method="post" name="bookingForm"><input type="hidden" value="<%=true%>" name="postback" /><input type="hidden" value="booking" name="calaction" /><input type="hidden" value="<%=booking.iId%>" name="iBookingID" /><input type="hidden" value="<%=cal.iId%>" name="iCalID" /><table border="0"><tr><td valign="top" style="width:600px%"><table cellpadding="2" border="0" cellpadding="4" cellspacing="0"><%if not isLeeg(booking.iId) then%><tr><td align="right">Unique ID:</td><td><%=booking.iId%></td></tr><%end if%><tr><td align="right">Status:</td><td><select name="iStatus"><%=statuslist.showSelected("option",booking.iStatus)%></select></td></tr><tr><td align="right">Start date:*</td><td><input type="text" id="dStartDate" name="dStartDate" size="13" value="<%=convertEuroDate(booking.dStartDate)%>" /><input type="checkbox" name="bSOnlyAfternoon" value="1" <%if booking.bSOnlyAfternoon then response.write "checked"%> /> Afternoon only
</td></tr><tr><td align="right">End date:*</td><td><input type="text" id="dEndDate" name="dEndDate" size="13" value="<%=convertEuroDate(booking.dEndDate)%>" /><input type="checkbox" name="bEOnlyMorning" value="1" <%if convertBool(booking.bEOnlyMorning) then response.write "checked"%> /> Morning only
<%=booking.customErrorMessage%></td></tr><tr><td>&nbsp;</td><td><%=JQDatePickerFT("dStartDate","dEndDate")%></td></tr><tr><td align="right">Name:*</td><td><input type="text" size="35" maxlength="50" name="sName" value="<%=quotrep(booking.sName)%>" /></td></tr><tr><td align="right">Email:</td><td><input type="email" size="35" maxlength="50" name="sEmail" value="<%=quotrep(booking.sEmail)%>" /></td></tr><tr><td align="right">Phone:</td><td><input type="text" size="35" maxlength="50" name="sPhone" value="<%=quotrep(booking.sPhone)%>" /></td></tr><tr><td align="right">Notes:</td><td><%createFCKInstance booking.sNotes,"siteBuilderMailSource","sNotes"%></td></tr><tr><td>&nbsp;</td><td><small>(*) Mandatory fields</small></td></tr><tr><td>&nbsp;</td><td><input class="art-button" type="submit" value="Save" name="btAct" /> <%if convertGetal(booking.iId)<>0 then %><input class="art-button" type="submit" value="Delete" onclick="javascript:return confirm('Are you sure to delete this reservation?');" name="btAct" /><%end if%></td></tr></table><ul><li><a href="bs_ac.asp?iCalID=<%=cal.iId%>&amp;calAction=bookings">Back to list of reservations</a></li></ul></td><td style="padding-left:20px;width:300px" align="center" valign="top"><script type="text/javascript" src="<%=C_DIRECTORY_QUICKERSITE%>/asp/bs_ac_view.asp?time=0&amp;legend=true&amp;months=12&amp;mode=js&amp;iCalID=<%=cal.iId%>"></script></td></tr></table></form><%set cal=nothing%>
