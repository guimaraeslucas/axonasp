
<%class calendarview
public iCalendarID
private ss,days
sub class_initialize
set ss=new cls_statuslist
set days=server.createobject("scripting.dictionary")
end sub
sub class_terminate
set ss=nothing
set days=nothing
end sub
private function getDays
if days.count=0 then
'get days
if not isLeeg(iCalendarID) then
dim rs,dValue,sValue
set rs=db.execute("select iId from tblCalendarBooking where iCalendarID="& iCalendarID & " order by dEndDate asc")
dim booking,Id,dc,dd
while not rs.eof
dc=0
set booking=new cls_calendarbooking
booking.pick(rs(0))
dd=DateDiff("d",booking.dStartDate,booking.dEnddate)
for ID=booking.dStartDate to booking.dEnddate
dc=dc+1
'switch over days
if days.exists(convertStr(booking.dStartDate)) then
if dc=1 then
days(convertStr(booking.dStartDate))=days(convertStr(booking.dStartDate)) & booking.iStatus & "OA"
end if
if dc=dd+1 then
days(convertStr(booking.dStartDate))=days(convertStr(booking.dStartDate)) & booking.iStatus & "OM"
end if
else
if dc=1 then
if booking.bSOnlyAfternoon then
days(convertStr(booking.dStartDate))=booking.iStatus & "OA"
else
days(convertStr(booking.dStartDate))=booking.iStatus
end if
else
days(convertStr(booking.dStartDate))=booking.iStatus
end if
if dc=dd+1 then
if booking.bEOnlyMorning then
days(convertStr(booking.dStartDate))=booking.iStatus & "OM"
end if
end if
end if
booking.dStartDate=dateAdd("d",1,booking.dStartDate)
next
set booking=nothing
rs.movenext()
wend
set rs=nothing
end if
end if
set getDays=days
end function
public function viewMonth(maand,jaar)
if isLeeg(maand) then maand=month(date())
if isLeeg(jaar) then jaar=year(date())
dim copyDays
set copyDays=getDays()
dim v
v="<table border=""0""  style=""font-size:10pt;border-style:none;background-color:#FFFFFF"" cellpadding=""0"" cellspacing=""0"">"
v=v&"<tr>"
v=v&"<td colspan=""7"" style=""vertical-align:middle;border-style:none;height:20px;background-color:#BBB;color:#333;text-align:center"" align=""center"">" & QSMonthName(maand) & "&nbsp;" & jaar &"</td>"
v=v&"</tr>"
v=v&"<tr>"
v=v&"[R3]"
v=v&"</tr>"
v=v&"<tr>"
dim i,firstday,d,addDays,vdays
firstday=weekday(dateserial(jaar,maand,1))
sdate=dateserial(jaar,maand,1)
for i=1 to firstday-1
vdays="<td align=""center"" style=""font-size:9pt;vertical-align:middle;border-style:none;margin:0px;padding:0px;width:20px;height:20px;background-color:#EEE;color:#BBB;text-align:center"">" & day(dateAdd("d",sdate,-i)) & "</td>" & vdays
next
v=v&vdays
dim goBacks
goBacks=0
for i=1 to 31
d=dateSerial(jaar,maand,i)
if maand=month(d) then
v=v& "<td valign=""middle"" align=""center"" style=""[R1]"
if copyDays.exists(convertStr(d)) then
v=v & ss.getCss(copyDays(convertStr(d)))
end if
v=v& """>"
v=v & i 
v=v& "</td>"
if convertGetal(i + convertGetal(firstday-1)) MOD 7 = 0 then
v=v&"</tr><tr>"
addDays=7
goBacks=goBacks+1
else
addDays=addDays-1
end if
else
exit for 
end if
next
for i=1 to addDays
v=v& "<td align=""center"" style=""[R2]"">" & i & "</td>"
next
dim j
for i=goBacks to 4
v=v& "<tr>"
for j=addDays+1 to addDays+7
v=v& "<td align=""center"" style=""[R2]"">" & j & "</td>"
next
v=v& "</tr>"
next
v=v&"</tr></table>"
v=replace(v,"</td><tr>","</td></tr><tr>",1,-1,1)
v=replace(v,"<tr></tr>","",1,-1,1)
v=replace(v,"</tr></tr>","</tr>",1,-1,1)
viewMonth=v
end function
end class
function QSMonthName(monthNr)

dim arrMonths
arrMonths=split(l("months"),"/")

select case monthNr
case 1
QSMonthName=arrMonths(0)
case 2
QSMonthName=arrMonths(1)
case 3
QSMonthName=arrMonths(2)
case 4
QSMonthName=arrMonths(3)
case 5
QSMonthName=arrMonths(4)
case 6
QSMonthName=arrMonths(5)
case 7
QSMonthName=arrMonths(6)
case 8
QSMonthName=arrMonths(7)
case 9
QSMonthName=arrMonths(8)
case 10
QSMonthName=arrMonths(9)
case 11
QSMonthName=arrMonths(10)
case 12
QSMonthName=arrMonths(11)
end select
end function%>
