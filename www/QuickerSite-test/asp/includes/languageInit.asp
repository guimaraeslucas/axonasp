
<%dim p_initLanguage2
set p_initLanguage2=nothing
function initLanguage2(byref code)
if p_initLanguage2 is nothing then
dim useLanguage
uselanguage=getCurrentLanguage
set p_initLanguage2=server.CreateObject ("scripting.dictionary")
dim sql, getrows, recordCount, rs, sValue, sCode, k
sql="select tblLabel.sCode, tblLabelValue.sValue from tblLabelValue INNER JOIN tblLabel on tblLabel.iId=tblLabelValue.iLabelId"
sql=sql&" where tblLabelValue.iLanguageID=" & useLanguage
set rs=db.getDynamicRSLabels
rs.open sql
if not rs.eof then
getrows	= rs.getRows()
recordCount	= rs.recordCount
end if
if not isNull(getrows) then
for k=0 to recordCount-1
sValue	= getrows(1,k)
sValue	= show(sValue)
sCode	= getrows(0,k)
'Response.Write sCode &"<br />"
'Response.Flush 
'if not p_initLanguage2.exists then
p_initLanguage2.Add sCode,sValue
'end if
Application(QS_CMS_lb_& useLanguage & sCode)=rebrand(sValue)
next
end if
set rs=nothing
end if
initLanguage2=p_initLanguage2(code)
end function

function l(code)
	
	dim uselanguage
	uselanguage=getCurrentLanguage
	
	if isLeeg(Application(QS_CMS_lb_& useLanguage & code)) then
		l=initLanguage2(lcase(code))
	else
		l=Application(QS_CMS_lb_& useLanguage & code)
	end if
	
end function

function rebrand (sTextToRebrand)

	if not isLeeg(sTextToRebrand) then
		rebrand=replace(sTextToRebrand,"http://www.quickersite.com/r/ufl",MYQS_urlUFLs,1,-1,1)
		rebrand=replace(sTextToRebrand,"http://www.quickersite.com/r/?sCode=TEMPLATES",MYQS_urlTemplates,1,-1,1)
		rebrand=replace(rebrand,"QuickerSite",MYQS_name,1,-1,1)
	end if
	
end function

function getCurrentLanguage()

	if not isleeg(session("QS_LANG")) then
		getCurrentLanguage=session("QS_LANG")
	elseif convertGetal(customer.language)<>0 then
		getCurrentLanguage=customer.language
	else
		getCurrentLanguage=1
	end if
	
end function%>
