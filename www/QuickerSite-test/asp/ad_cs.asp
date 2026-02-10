<!-- #include file="begin.asp"-->


<%'QuickerSite - (C) 2006-09 by P. Cooreman%><%dim customerObj
set customerObj=new cls_customer
customerObj.sName	= "New site"
customerObj.sURL	= "http://localhost"
customerObj.dOnlineFrom	= date()
customerObj.bApplication	= false
customerObj.bScanReferer	= false
customerObj.sAlternateDomains	= ""
customerObj.bUserFriendlyURL	= false
customerObj.bAllowStorageOutsideWWW = false
customerObj.bEnableMainRSS	= false
customerObj.adminPassword	= sha256(QS_defaultPW)
customerObj.webmasterEmail	= "your@email.com"
if customerObj.save() then
end if
set customerObj=nothing%>
