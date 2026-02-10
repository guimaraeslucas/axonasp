<!-- #include file="begin.asp"-->


<!-- #include file="includes/ac_calendar.asp"--><!-- #include file="includes/ac_calendarview.asp"--><!-- #include file="includes/ac_calendarbooking.asp"--><!-- #include file="includes/ac_statuslist.asp"--><%dim calObj,mycals,cal,calendar,cc,sCode,cv,statuslist,monthlist,monthYear,setMonth,setYear,arrD,rs,sql,i
dim action,postback,color,sDate,daysArr
daysArr=split(l("days"),"/")

action=request("calAction")
postback=convertBool(request.form("postBack"))
response.clear()
set cal=new cls_calendar
cal.secure=false
cal.pick(request("iCalID"))
if convertGetal(cal.iId)=0 then
response.write "Calendar not available"
set cal=nothing
response.end
end if
set cv=new calendarview
cv.iCalendarID=cal.iId
set statuslist=new cls_statusList
dim view
view="<div id=""calendarDIV"&cal.iId&""">"
dim mm
if convertGetal(request.querystring("months"))<>0 then
mm=convertGetal(request.querystring("months"))
else
mm=12
end if
if convertGetal(mm)>36 then 
mm=36
end if
select case request.querystring("time")
case "0" ' x-number of months
sdate=dateSerial(year(date()),month(date()),day(date()))
if not isLeeg(request.form("ajax")) then
select case request.form("ajax")
case "prev"
session(cal.iId&"c0")=convertGetal(session(cal.iId&"c0"))-1
case "next"
session(cal.iId&"c0")=convertGetal(session(cal.iId&"c0"))+1
end select
sdate=dateAdd("m",convertGetal(session(cal.iId&"c0"))*mm,sdate)
else
session(cal.iId&"c0")=0
end if
for i=1 to mm
view=view& "<div style=""float:left;margin:3px"">" & cv.viewMonth(month(sdate),year(sdate)) & "</div>"
sdate=dateAdd("m",1,sdate)
next
case "1" ' current calendar year
sdate=dateSerial(year(date()),1,1)
if not isLeeg(request.form("ajax")) then
select case request.form("ajax")
case "prev"
session(cal.iId&"c1")=convertGetal(session(cal.iId&"c1"))-1
case "next"
session(cal.iId&"c1")=convertGetal(session(cal.iId&"c1"))+1
end select
sdate=dateAdd("m",convertGetal(session(cal.iId&"c1"))*mm,sdate)
else
session(cal.iId&"c1")=0
end if
for i=1 to mm
view=view& "<div style=""float:left;margin:3px"">" & cv.viewMonth(month(sdate),year(sdate)) & "</div>"
sdate=dateAdd("m",1,sdate)
next
case "2" ' next calendar year
sdate=dateSerial(year(date())+1,1,1)
if not isLeeg(request.form("ajax")) then
select case request.form("ajax")
case "prev"
session(cal.iId&"c2")=convertGetal(session(cal.iId&"c2"))-1
case "next"
session(cal.iId&"c2")=convertGetal(session(cal.iId&"c2"))+1
end select
sdate=dateAdd("m",convertGetal(session(cal.iId&"c2"))*mm,sdate)
else
session(cal.iId&"c2")=0
end if
for i=1 to mm
view=view& "<div style=""float:left;margin:3px"">" & cv.viewMonth(month(sdate),year(sdate)) & "</div>"
sdate=dateAdd("m",1,sdate)
next
end select 
dim showNav
showNav=true
if showNav then
view=view & "<div style=""clear:both""> </div>"
view=view & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" style=""font-size:9pt;margin:0px;padding:0px;border-style:none;width:100%"">"
view=view & "<tr><td style=""font-size:9pt;width:100%;border-style:none;text-align:center"" align=""center"">"
view=view & "<table cellpadding=""0"" cellspacing=""0"" border=""0"" style=""font-size:10pt;margin:0px;padding:0px;border-style:none;width:95%""><tr>"
view=view & "<td align=""right"" style=""text-align:right;border-style:none;padding-right:10px""><a href=""#"" onclick=""javascript:get"&cal.iId&"('prev');return false;""><strong>&lt; " & l("prevshort") & "</strong></a></td>"
view=view & "<td style=""border-style:none;padding-left:10px;text-align:left"" align=""left""><a href=""#""  onclick=""javascript:get"&cal.iId&"('next');return false;""><strong>" & l("nextshort") & " &gt;</strong></a></td>"
view=view & "</tr></table></td></tr></table>"
end if
'legend
if request.querystring("legend")<>"false" then
view=view & "<div style=""clear:both""> </div>"
view=view & "<table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""font-size:10pt;margin:0px;padding:0px;border-style:none;width:100%;text-align:center"">"
view=view & "<tr><td align=""center"" style=""font-size:9pt;margin:0px;padding:0px;border-style:none;text-align:center"">"
view=view & "<table border=""0"" style=""width:100%;text-align:center;border-style:none;margin:5px;padding:0px"" cellpadding=""0"" cellspacing=""0"">"
view=view & "<tr>"
view=view & "<td style=""margin:0px;padding:0px;border-style:none;width:20px;height:20px;background-image:url("& C_DIRECTORY_QUICKERSITE & "/fixedImages/ACimages/10.gif)""></td><td style=""text-align:left;margin:0px;padding:0px;border-style:none;"">&nbsp;" & l("Available") & "&nbsp;&nbsp;&nbsp;</td>"
view=view & "<td style=""margin:0px;padding:0px;border-style:none;width:20px;height:20px;background-image:url("& C_DIRECTORY_QUICKERSITE & "/fixedImages/ACimages/15.gif)""></td><td style=""text-align:left;margin:0px;padding:0px;border-style:none;"">&nbsp;" & l("Unavailable") & "&nbsp;&nbsp;&nbsp;</td>"
view=view & "<td style=""margin:0px;padding:0px;border-style:none;width:20px;height:20px;background-image:url("& C_DIRECTORY_QUICKERSITE & "/fixedImages/ACimages/17.gif)""></td><td style=""text-align:left;margin:0px;padding:0px;border-style:none;"">&nbsp;" & l("Pending") & "&nbsp;&nbsp;&nbsp;</td>"
view=view & "<td style=""margin:0px;padding:0px;border-style:none;width:20px;height:20px;background-image:url("& C_DIRECTORY_QUICKERSITE & "/fixedImages/ACimages/20.gif)""></td><td style=""text-align:left;margin:0px;padding:0px;border-style:none;"">&nbsp;" & l("Booked")  & "&nbsp;&nbsp;&nbsp;</td>"
view=view & "</tr>"
view=view & "</table>"
view=view & "</td></tr></table>"
end if
view=view&"</div>" 'end calendarDIV
view=replace(view,"[R1]","vertical-align:middle;margin:0px;padding:0px;border-style:none;color:#000;background-color:#DDD;color:#222;width:20px;height:20px;text-align:center;",1,-1,1)
view=replace(view,"[R2]","vertical-align:middle;margin:0px;padding:0px;border-style:none;width:20px;height:20px;color:#BBB;background-color:#EEE;text-align:center;",1,-1,1)
view=replace(view,"[R3]","<td align=""center"" style=""vertical-align:middle;margin:0px;padding:0px;border-style:none;text-align:center""><small>" & daysArr(6) & "</small></td><td style=""vertical-align:middle;margin:0px;padding:0px;border-style:none;text-align:center"" align=""center""><small>" & daysArr(0) & "</small></td><td style=""vertical-align:middle;margin:0px;padding:0px;border-style:none;text-align:center"" align=""center""><small>" & daysArr(1) & "</small></td><td style=""vertical-align:middle;margin:0px;padding:0px;border-style:none;text-align:center"" align=""center""><small>" & daysArr(2) & "</small></td><td style=""vertical-align:middle;margin:0px;padding:0px;border-style:none;text-align:center"" align=""center""><small>" & daysArr(3) & "</small></td><td style=""vertical-align:middle;margin:0px;padding:0px;border-style:none;text-align:center"" align=""center""><small>" & daysArr(4) & "</small></td><td style=""vertical-align:middle;margin:0px;padding:0px;border-style:none;text-align:center"" align=""center""><small>" & daysArr(5) & "</small></td>",1,-1,1)
view=view & "<!--" & PrintTimer(startTimer) & "-->"
view="<table align=""center"" style=""margin: 0px; padding: 0px; border-style: none; width: 100%;""><tbody>	<tr><td style=""margin: 0px; padding: 0px; border-style: none;"">" & view & "</td></tr></tbody></table>"
select case request.querystring("mode")
case "if","html"
response.write "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN" & """"
response.write """" & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">"
response.write "<html xmlns=""http://www.w3.org/1999/xhtml"" lang=""en"" xml:lang=""en"">"
response.write "<head><style=""td,th {font-size:9pt !important}""></style><title>" & quotrep(cal.sName) & "</title>" & getAjaxJS(false)
response.write "</head><body style=""text-align:center;background-color:"
if not isleeg(cal.sBGcolor) then
response.write cal.sBGcolor
else
response.write "#FFFFFF"
end if
response.write ";font-family:"
if not isleeg(cal.sFontFamily) then
response.write cal.sFontFamily
else
response.write "Verdana"
end if
response.write ";font-size:9pt;""><table border=""0"" cellpadding=""0"" cellspacing=""0"" style=""width:100%""><tr><td style=""border-style:none;text-align:center"" align=""center""><table border=""0"" cellpadding=""0"" cellspacing=""0""><tr><td style=""border-style:none"">" & view & "</td></tr></table></td></tr></table></body></html>"
case "js"
if isLeeg(request.form("ajax")) then
response.write "document.write ('" & replace(view,"'","\'",1,-1,1) & "');" & vbcrlf
response.write getAjaxJS(true)
else
response.write view
end if
end select
function getAjaxJS(removeScript)
on error resume next
if showNav then
dim fso
set fso=server.createObject("scripting.filesystemobject")
getAjaxJS=fso.opentextfile(server.mappath("ajax.txt"),1).readAll()
getAjaxJS=replace(getAjaxJS,"[URL]",C_DIRECTORY_QUICKERSITE & "/asp/bs_ac_view.asp?legend=" & request.querystring("legend") & "\u0026months=" & request.querystring("months") & "\u0026time=" & request.querystring("time") & "\u0026mode=" & request.querystring("mode") & "\u0026iCalID="& cal.iID,1,-1,1)
getAjaxJS=replace(getAjaxJS,"calendarDIV","calendarDIV" & cal.iId,1,-1,1)
getAjaxJS=replace(getAjaxJS,"[CALID]",cal.iId,1,-1,1)
if removeScript then
getAjaxJS=replace(getAjaxJS,"<script type=""text/javascript"">","",1,-1,1)
getAjaxJS=replace(getAjaxJS,"</script>","",1,-1,1)
end if
set fso=nothing
end if
on error goto 0
end function
set cal=nothing
set cv=nothing
response.end %>
