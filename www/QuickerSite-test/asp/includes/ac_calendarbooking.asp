
<%class cls_calendarbooking
public iId,iStatus,dStartDate,dEndDate,bSOnlyAfternoon,bEOnlyMorning,sUniqueKey,sName,sEmail,sPhone,sNotes
public iCalendarID,dCreatedTS,dUpdatedTS,customErrorMessage
sub class_initialize
end sub
public function getRequestValues()
iStatus	= request.form("iStatus")
dStartDate	= convertDateFromPicker(request.form("dStartDate"))
dEndDate	= convertDateFromPicker(request.form("dEndDate"))
bSOnlyAfternoon	= convertBool(request.form("bSOnlyAfternoon"))
bEOnlyMorning	= convertBool(request.form("bEOnlyMorning"))
sName	= request.form("sName")
sEmail	= request.form("sEmail")
sPhone	= request.form("sPhone")
sNotes	= request.form("sNotes")
iCalendarID	= request.form("iCalID")
end function
public function pick(id)
if isNumeriek(id) then
dim sql
sql="select * from tblCalendarBooking where iId=" & left(id,9)
set rs=db.execute(sql)
if not rs.eof then
iId	= rs("iId")
iStatus	= rs("iStatus")
dStartDate	= rs("dStartDate")
dEndDate	= rs("dEndDate")
bSOnlyAfternoon	= rs("bSOnlyAfternoon")
bEOnlyMorning	= rs("bEOnlyMorning")
sUniqueKey	= rs("sUniqueKey")
sName	= rs("sName")
sEmail	= rs("sEmail")
sPhone	= rs("sPhone")
sNotes	= rs("sNotes")
iCalendarID	= rs("iCalendarID")
dCreatedTS	= rs("dCreatedTS")
dUpdatedTS	= rs("dUpdatedTS")
end if
set rs=nothing 
end if
end function
public function check()
check=true
if isLeeg(sName) then 
message.addError("err_mandatory")
check=false
end if
if isLeeg(dStartDate) then 
message.addError("err_mandatory")
check=false
end if
if isLeeg(dEndDate) then 
message.addError("err_mandatory")
check=false
end if
if dEndDate<dStartDate then 
customErrorMessage="<p><font color=""Red""><strong>End date cannot be before start date!</strong></font></p>"
check=false
end if
if check then
'check on double bookings
dim bs,b
set bs=calObj.bookings
for each b in bs
if dStartDate<bs(b).dEndDate and dEnddate>bs(b).dEndDate and b<>iId then
check=false
customErrorMessage="y"
end if
if dStartDate>bs(b).dStartDate and dStartDate<bs(b).dEndDate and b<>iId then
check=false
customErrorMessage="y"
end if
if dEndDate>bs(b).dStartDate and dEndDate<bs(b).dEndDate and b<>iId then
check=false
customErrorMessage="y"
end if
if dEndDate=bs(b).dStartDate and b<>iId then
if not (convertBool(bEOnlyMorning) and convertBool(bs(b).bSOnlyAfternoon)) then
check=false
customErrorMessage="y"
end if
end if
if dStartDate=bs(b).dEnddate and b<>iId then
if not (convertBool(bSOnlyAfternoon) and convertBool(bs(b).bEOnlyMorning)) then
check=false
customErrorMessage="y"
end if
end if
if not check then exit for
next 
if customErrorMessage="y" then
customErrorMessage="<p><font color=""Red""><strong>Double bookings are not allowed!</strong></font></p>"
end if
end if
end function
public function calObj
set calObj=new cls_calendar
if not isLeeg(iCalendarID) then
calObj.pick(iCalendarID)
end if
end function
public function save()
save=check
if not save then exit function
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblCalendarBooking where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblCalendarBooking where iId="& iId
end if
rs("iStatus")	= iStatus
rs("dStartDate")	= dStartDate
rs("dEndDate")	= dEndDate
rs("bSOnlyAfternoon")	= bSOnlyAfternoon
rs("bEOnlyMorning")	= bEOnlyMorning
rs("sUniqueKey")	= sUniqueKey
rs("sName")	= sName
rs("sEmail")	= sEmail
rs("sPhone")	= sPhone
rs("sNotes")	= sNotes
rs("iCalendarID")	= iCalendarID
rs("dUpdatedTS")	= dUpdatedTS
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function delete
set rs=db.execute("delete from tblCalendarBooking where iId="& convertGetal(iId))
set rs=nothing
end function
end class%>
