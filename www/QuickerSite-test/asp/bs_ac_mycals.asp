
<%set  calObj=new cls_calendar
set mycals= calObj.mycals
if mycals.count>0 then
response.write "<ul>"
end if
for each cal in mycals
response.write "<li><strong>" & quotrep(mycals(cal).sName) & "</strong>"
response.write "<ul>"
response.write "<li><a href=""bs_ac.asp?iCalID=" & cal & "&amp;calAction=editCal"">Modify / Delete</a></li>"
response.write "<li><a href=""bs_ac.asp?iCalID=" & cal & "&amp;calAction=bookings"">Reservations</a> - <a href=""bs_ac.asp?iCalID=" & cal & "&amp;calAction=booking"">Add Reservation</a></li>"
response.write "<li><a target=""_blank"" class=""QSPP"" href=""" & C_DIRECTORY_QUICKERSITE & "/asp/bs_ac_view.asp?time=1&amp;mode=html&amp;iCalID=" & cal & """>Preview Calendar</a></li>"
response.write "<li><a href=""bs_ac.asp?iCalID=" & cal & "&amp;calAction=embedCode"">Add this Calendar to my website</a></li>"
response.write "</ul>"
response.write "</li>"
next
if mycals.count>0 then
response.write "</ul>"
end if
if mycals.count=0 then
response.write "<p align=""center"">No calendars yet.</p>"
end if
set mycals=nothing
set  calObj=nothing%>
