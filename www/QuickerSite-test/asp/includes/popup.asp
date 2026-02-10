
<%class cls_popup
Public iId,sName,sUrl,dOnlineFrom,dOnlineUntil,iHeight,iWidth,iCustomerID,bEnabled,iTemplateID,sValue,iMode,iShows,sViewmode,bUseAsSystemMessage,messageMode,bSticky,bUseAsPopup,iAutoclose
Private Sub Class_Initialize
On Error Resume Next
bUseAsSystemMessage=false
bUseAsPopup=false
bSticky=false
iAutoclose=0
bEnabled=false
iHeight=300
iWidth=400
sUrl="http://"
iMode=0
iShows=0
sViewmode="1"
sValue="<table cellspacing=""0"" cellpadding=""20"" style=""width: 100%;""><tbody><tr><td align=""center""><p style=""text-align: center;"">text goes here...</p></td></tr></tbody></table>"
pick(decrypt(request("iPopupID")))
err.clear()
On Error Goto 0
end sub
Public Property Get sGetViewmode
if isLeeg(sViewmode) then
sGetViewmode="1"
else
sGetViewmode=sViewmode
end if
end property
Public property get sGetUrl
if not isLeeg(sUrl) and sUrl<>"http://" then
sGetUrl=sUrl
else
sGetUrl=customer.sQSUrl & "/default.asp?pageAction=showPopup&iPopupID=" & encrypt(iId)
end if
end property
Public property get bOnline
bOnline=isBetween(dOnlineFrom,date,dOnlineUntil)
if bOnline then
session("popupLoadCount"& iId)=convertGetal(session("popupLoadCount"& iId))+1
select case convertGetal(iMode)
case 0
' do nothing
case 1
if isLeeg(request.cookies("popupLoad"& iId)) then
Response.Cookies("popupLoad"& iId) = "yes"
Response.Cookies("popupLoad"& iId).expires=dateAdd("d",365,date())
else
bOnline=false
end if
case 2
if not convertGetal(session("popupLoadCount"& iId))=1 then
bOnline=false
end if
case 3
if not convertGetal(session("popupLoadCount"& iId))=2 then
bOnline=false
end if
case 4
if not convertGetal(session("popupLoadCount"& iId))=3 then
bOnline=false
end if
case 5
if not convertGetal(session("popupLoadCount"& iId))=4 then
bOnline=false
end if
case 6
if not convertGetal(session("popupLoadCount"& iId))=5 then
bOnline=false
end if
case 7
if not convertGetal(session("popupLoadCount"& iId))=6 then
bOnline=false
end if
case 8
if not convertGetal(session("popupLoadCount"& iId))=7 then
bOnline=false
end if
case 9
if not convertGetal(session("popupLoadCount"& iId))=8 then
bOnline=false
end if
case 10
if not convertGetal(session("popupLoadCount"& iId))=9 then
bOnline=false
end if
end select
end if
if bOnline then
session("PopupLoaded")="yes"
end if
'response.write bOnline
'response.end 
end property
Public Function getRequestValues()
sName	= left(trim(Request.Form ("sName")),50)
sUrl	= trim(Request.Form ("sUrl"))
dOnlineFrom	= convertDateFromPicker(Request.Form ("dOnlineFrom"))
dOnlineUntil	= convertDateFromPicker(Request.Form ("dOnlineUntil"))
iHeight	= convertGetal(Request.Form ("iHeight"))
iWidth	= convertGetal(Request.Form ("iWidth"))
bEnabled	= convertBool(request.form("bEnabled"))
iTemplateID	= convertGetal(Request.Form ("iTemplateID"))
iAutoclose	= convertGetal(request.form("iAutoclose"))
sValue	= removeEmptyP(request.form("sValue"))
iMode	= convertGetal(request.form("iMode"))
sViewmode	= request.form("sViewmode")
end Function
Public Function Pick(id)
dim sql, rs
dim aArr(2)
if isNumeriek(id) then
sql = "select * from tblPopup where iCustomerID="&cid&" and iId=" & id
set rs = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
sUrl	= rs("sUrl")
dOnlineFrom	= rs("dOnlineFrom")
dOnlineUntil	= rs("dOnlineUntil")
iHeight	= rs("iHeight")
iWidth	= rs("iWidth")
bEnabled	= rs("bEnabled")
iTemplateID	= rs("iTemplateID")
sValue	= rs("sValue")
iMode	= rs("iMode")
iShows	= rs("iShows")
sViewmode	= rs("sViewmode")
iAutoclose	= rs("iAutoclose")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sUrl) and isLeeg(sValue) then
check=false
message.AddError("err_mandatory")
end if
End Function
Public Function Save()
if check() then
Save=true
else
Save=false
exit function
end if
'set db=nothing
'set db=new cls_database
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblPopup where 1=2"
rs.AddNew
else
rs.Open "select * from tblPopup where iId="& iId
end if
rs("sName")	= sName
rs("sUrl")	= sUrl
rs("dOnlineFrom")	= dOnlineFrom
rs("dOnlineUntil")	= dOnlineUntil
rs("iHeight")	= iHeight
rs("iWidth")	= iWidth
rs("bEnabled")	= bEnabled
rs("iTemplateID")	= iTemplateID
rs("iCustomerID")	= cId
rs("sValue")	= sValue
rs("iMode")	= iMode
rs("iShows")	= iShows
rs("sViewmode")	= sViewmode
rs("iAutoclose")	= iAutoclose
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblPopup where iId=" & iId)
set rs=nothing
end if
end function
public function copy()
if isNumeriek(iId) then
iId=null
save()
end if
end function
public function dumpJS
if bNOPopup then exit function
if convertGetal(iId)<> 0 then
dumpJS="<script type=""text/javascript"">" & vbcrlf
dumpJS=dumpJS&"$(window).load(function(){" & vbcrlf
dumpJS=dumpJS&"var sGetUrl = " & """" & sGetUrl & """" & ";" & vbcrlf

dumpJS=dumpJS&" if ($(this).width() > 0) { " & vbcrlf

dumpJS=dumpJS&"$.colorbox({close: """ & quotrep(l("close")) & """, open:true, innerWidth:" & iWidth & ", title:""" & quotrep(sName) & """, innerHeight:" & iHeight & ", iframe:true, href:sGetUrl}); " & vbcrlf
if convertGetal(iAutoclose)<>0 then
dumpJS=dumpJS&"var t=setTimeout($.colorbox.close," & convertGetal(iAutoclose)*1000 & "); " & vbcrlf
end if
dumpJS=dumpJS&"}})"& vbcrlf


dumpJS=dumpJS&"</script>" & vbcrlf
'iAutoclose
elseif bUseAsSystemMessage then
dumpJS="<script type=""text/javascript"">" & vbcrlf
dumpJS=dumpJS&"$(window).load(function(){" & vbcrlf
select case messageMode
case "fb"
dumpJS=dumpJS&"$.colorbox({close: """ & quotrep(l("close")) & """, fixed: true, open:true, width:""400"", inline:true, href:""#fbMessage""}); " & vbcrlf
case "err"
dumpJS=dumpJS&"$.colorbox({close: """ & quotrep(l("close")) & """, fixed: true, open:true, width:""400"", inline:true, href:""#errMessage""}); " & vbcrlf
end select
if not bSticky then
dumpJS=dumpJS&"var t=setTimeout($.colorbox.close,1100); " & vbcrlf
end if
dumpJS=dumpJS&"});" & vbcrlf & "</script>" & vbcrlf
elseif bUseAsPopup then
dumpJS="<script type=""text/javascript"">" & vbcrlf
dumpJS=dumpJS&"$(document).ready(function(){" & vbcrlf
dumpJS=dumpJS&"$("".cbPopUP"").colorbox({close: """ & quotrep(l("close")) & """, innerWidth:" & iWidth & ", title:""" & quotrep(sName) & """, innerHeight:" & iHeight & ", iframe:true, onClosed:function(){ location.reload(true);}}); " & vbcrlf
dumpJS=dumpJS&"});</script>" & vbcrlf
end if
end function
public function close()
headerDictionary.add generatePassword(),"<script type=""text/javascript"">parent.$.fn.colorbox.close();</script>"
end function
end class%>
