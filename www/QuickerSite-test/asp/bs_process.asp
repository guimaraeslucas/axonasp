
<%dim page
set page=new cls_page
page.pick(decrypt(request("iId")))
select case request("btnaction")
case "ConvertToFP"
checkCSRF()
if page.convertToFP then
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
else
message.AddError("err_mandatory")
end if
case "ResetStats"
checkCSRF()
customer.resetStats()
case "MoveUp"
checkCSRF()
page.moveUp()
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
case "MoveDown"
checkCSRF()
page.moveDown()
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
case "Move"
checkCSRF()
if isNumeriek(page.iId) then
Response.Redirect ("bs_selectPage.asp?"&QS_secCodeURL&"&btnaction=selectPage&iId="& EnCrypt(page.iId))
end if
case "Insert"
checkCSRF()
'in het oude parent moet deze child worden weggerekend in de rangen
page.parentPage.removeRang(page)
'in de nieuwe parent moet deze child achteraan worden toegevoegd.
page.iParentID=convertLng(DeCrypt(Request.QueryString ("insertInto")))
'free pages omzetten naar menu-pagina's
page.bLossePagina=false
if not page.parentPage.bOnline then
page.bOnline=false
end if
'nieuwe rang!
page.iRang=convertGetal(page.getRang)+1
'paswoord synchroniseren
if not isLeeg(page.parentPage.sPw) then
page.sPw=page.parentPage.sPw
page.resetAllSubPasswords(page.iId)
end if
page.save()
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
case "Copy"
checkCSRF()
page.bCopy=true
page.copy()
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
case "continue"
dim redirect
select case convertGetal(Request.Form ("fixedType"))
case sb_item
redirect="bs_editItem.asp?"
case sb_container
redirect="bs_editContainer.asp?"
case sb_externalURL
redirect="bs_editExternalURL.asp?"
case sb_list
redirect="bs_editList.asp?"
case sb_lossePagina
redirect="bs_editItem.asp?bLossePagina="& true & "&"
case sb_constant
redirect="bs_constantEdit.asp?"
case else
message.AddError("err_makeChoice")
end select
if not isLeeg(redirect) then Response.Redirect (redirect&"bIntranet="& Request.form("bIntranet") &"&iParentID="&Request.Form ("iParentID"))
case "removeFavicon"
checkCSRF()
dim fsoFAV
set fsoFAV=server.CreateObject ("scripting.filesystemobject")
if fsoFAV.FileExists (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "favicon.ico")) then
fsoFAV.DeleteFile (server.MapPath (C_VIRT_DIR & Application("QS_CMS_userfiles") & "favicon.ico"))
end if
set fsoFAV=nothing
customer.hasFavicon=false
customer.save
case "saveAdmin"
checkCSRF()
if logon.currentPW=customer.adminPassword then
customer.sUrl	= Request.Form ("sUrl")
customer.sAlternateDomains	= Request.Form ("sAlternateDomains")
end if
customer.sDescription	= Request.Form ("sDescription")
customer.siteName	= Request.Form ("siteName")
customer.siteTitle	= Request.Form ("siteTitle")
customer.sSiteSlogan	= Request.Form ("sSiteSlogan")
customer.copyRight	= Request.Form ("copyRight")
customer.keywords	= Request.Form ("keywords")
customer.googleAnalytics	= Request.Form ("googleAnalytics")
customer.sHeader	= Request.Form ("sHeader")
customer.language	= Request.Form ("language")
customer.sDatumFormat	= Request.Form ("sDatumFormat")
customer.webmaster	= Request.Form ("webmaster")
customer.webmasterEmail	= Request.Form ("webmasterEmail")
customer.sDefaultRSSLink	= Request.Form ("sDefaultRSSLink")
if customer.save() then
application(QS_textDirection)	= ""
clearMenuCache
message.Add ("fb_saveOK")
end if
case "saveAdminIntranetOSM"
checkCSRF()
customer.sWelcomeMessage	= removeEmptyP(Request.Form ("sWelcomeMessage"))
customer.sExplTicket	= removeEmptyP(Request.Form ("sExplTicket"))
customer.sExplProfile	= removeEmptyP(Request.Form ("sExplProfile"))
if customer.save() then
message.Add ("fb_saveOK")
end if
case "saveAdminIntranetEM"
checkCSRF()
customer.sMailTicket	= removeEmptyP(Request.Form ("sMailTicket"))
customer.sSubjectMailTicket	= Request.Form ("sSubjectMailTicket")
customer.intranetPWEmail	= removeEmptyP(Request.Form ("intranetPWEmail"))
set csShortList=new cls_contactStatusList
set csShortList=csShortList.list
csShortList.remove(cs_Silent)
dim key2, iMess
for each key2 in csShortList
set iMess=new cls_customerImess
iMess.pick(key2)
iMess.iStatus	= key2
iMess.sSubject	= removeEmptyP(Request.Form ("sSubject"&key2))
iMess.sBody	= removeEmptyP(Request.Form ("sBody"&key2))
iMess.bEnabled	= convertBool(Request.Form ("bEnabled"&key2))
iMess.save()
set iMess=nothing
next 
if customer.save() then
message.Add ("fb_saveOK")
end if
case l("delete_pw_total")
checkCSRF()
if customer.removeTotalPW() then
message.Add ("fb_saveOK")
end if
case l("save_pw_total")
checkCSRF()
if customer.applyTotalPW(Request.Form("sTotalPW")) then
message.Add ("fb_saveOK")
end if
case "makenormalpage"
checkCSRF()
if page.makeNormalPage() then
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
end if
case cSaveAdminPW
checkCSRF()
if Request.Form("adminPassword")=Request.Form("adminPasswordConfirm") then
if customer.saveAdminPW(Request.Form("adminPassword")) then
message.Add ("fb_saveOK")
end if
else
message.AddError("pwnomatch")
end if
case "savepeel"
checkCSRF()
if request.form("btnDelete")<>"" then
customer.sPeelURL=""
customer.sPeelImage=""
customer.bPeelOINW=false
customer.bPeelEnabled=false
customer.sPeelFlipColor=""
customer.sPeelIdleSize=0
customer.sPeelMOSize=0
customer.save
response.redirect("bs_Peel.asp?fbMessage=fb_saveOK")
else
customer.sPeelURL=left(request.form("sPeelURL"),255)
customer.bPeelEnabled=convertBool(request.form("bPeelEnabled"))
customer.bPeelOINW=convertBool(request.form("bPeelOINW"))
customer.sPeelFlipColor=request.form("sPeelFlipColor")
customer.sPeelIdleSize=convertGetal(request.form("sPeelIdleSize"))
customer.sPeelMOSize=convertGetal(request.form("sPeelMOSize"))
if isLeeg(customer.sPeelURL) or customer.sPeelURL="http://" or isLeeg(customer.sPeelFlipColor) then
message.AddError("err_mandatory")
elseif customer.save() then
if isLeeg(customer.sPeelImage) then
response.redirect("bs_selectPeel.asp")
else
response.redirect("bs_Peel.asp?fbMessage=fb_saveOK")
end if
end if
end if 
end select%>
