
<%class cls_secondAdmin
Public sPassword,sPasswordConfirm,bSetupGeneral,bSetupPageElements,bStats,bTemplates,bHomeConstants,bShoppingCart
Public bHomeVBScript,bFiles,bForms,bIntranet,bIntranetSetup,bIntranetContacts,bIntranetMail,bCatalog,bGuestbook
Public bFeed,bTheme,bFormExport,bApplicationpath,bGallery,bPagesPW,bPagesAdd,bPagesMove,bPagesDelete,bCustom404
Public bPageUrlRSS,bPageAdditionalHeader,bPageDescription,bPageKeywords,bPageTitleTag,bPageRefresh,bPageTemplate,bPageTheme,bAvailabilityCal
Public bPageForm,bPageCatalog,bPageFeed,bPageSetHomepage,bPagePublish,bPageOrder,bPageTitle,bPageBody, bPageUFL, bCalendar, bPoll, bPopup, bNewsletter, bPeel, bCookieWarning
Public iCustomerID
private sub class_initialize
pick()
end sub
Public Function Pick()
dim sql, RS
sql = "select * from tblSecondAdmin where iCustomerID="&cid
set RS = db.execute(sql)
if not rs.eof then
sPassword	= rs("sPassword")
sPasswordConfirm	= sPassword
bSetupGeneral	= rs("bSetupGeneral")
bSetupPageElements	= rs("bSetupPageElements")
bStats	= rs("bStats")
bTemplates	= rs("bTemplates")
bHomeConstants	= rs("bHomeConstants")
if customer.bApplication then
bHomeVBScript	= rs("bHomeVBScript")
else
bHomeVBScript	= false
end if
bFiles	= rs("bFiles")
bForms	= rs("bForms")
bIntranet	= rs("bIntranet")
bIntranetSetup	= rs("bIntranetSetup")
bIntranetContacts	= rs("bIntranetContacts")
bIntranetMail	= rs("bIntranetMail")
bCatalog	= rs("bCatalog")
bFeed	= rs("bFeed")
bTheme	= rs("bTheme")
bFormExport	= rs("bFormExport")
bApplicationpath	= rs("bApplicationpath")
bGallery	= rs("bGallery")
bPagesAdd	= rs("bPagesAdd")
bPagesPW	= rs("bPagesPW")
bPagesMove	= rs("bPagesMove")
bPagesDelete	= rs("bPagesDelete")
bPageUrlRSS	= rs("bPageUrlRSS")
bPageAdditionalHeader	= rs("bPageAdditionalHeader")
bPageDescription	= rs("bPageDescription")
bPageKeywords	= rs("bPageKeywords")
bPageTitleTag	= rs("bPageTitleTag")
bPageRefresh	= rs("bPageRefresh")
bPageTemplate	= rs("bPageTemplate")
bPageTheme	= rs("bPageTheme")
bPageForm	= rs("bPageForm")
bPageCatalog	= rs("bPageCatalog")
bPageFeed	= rs("bPageFeed")
bPageSetHomepage	= rs("bPageSetHomepage")
bPagePublish	= rs("bPagePublish")
bPageOrder	= rs("bPageOrder")
bPageTitle	= rs("bPageTitle")
bPageBody	= rs("bPageBody")
bPageUFL	= rs("bPageUFL")
bPoll	= rs("bPoll")
bGuestbook	= rs("bGuestbook")
bPopup	= rs("bPopup")
bNewsletter	= rs("bNewsletter")
iCustomerID	= rs("iCustomerID")
bAvailabilityCal	= rs("bAvailabilityCal")
bShoppingCart	= rs("bShoppingCart")
bCustom404	= rs("bCustom404")
bPeel	= rs("bPeel")
bCookieWarning	= rs("bCookieWarning")
end if
set RS = nothing
end function
private function check()
check=true
if isLeeg(sPassword) or isLeeg(sPasswordConfirm) then
message.AddError("err_mandatory")
check=false
elseif sPasswordConfirm<>sPassword then
message.AddError("pwnomatch")
check=false
elseif sPassword=customer.adminPassword then
message.AddError("err_firstandsecondpassword")
check=false
elseif sPassword=sha256(QS_defaultPW) then
check=false
message.AddError("err_backsitepw")
end if
if not check then
sPasswordConfirm=""
sPassword=""
end if
end function
Public Function Save
if check() then
save=true
else
save=false
exit function
end if
dim rs
set rs = db.GetDynamicRS
rs.Open "select * from tblSecondAdmin where iCustomerID="&cId
if rs.eof then
rs.AddNew()
end if
rs("sPassword")	= sPassword
rs("bSetupGeneral")	= convertBool(bSetupGeneral)
rs("bSetupPageElements")	= convertBool(bSetupPageElements)
rs("bStats")	= convertBool(bStats)
rs("bTemplates")	= convertBool(bTemplates)
rs("bHomeConstants")	= convertBool(bHomeConstants)
if customer.bApplication then
rs("bHomeVBScript")	= convertBool(bHomeVBScript)
else
rs("bHomeVBScript")	= false
end if
rs("bFiles")	= convertBool(bFiles)
rs("bForms")	= convertBool(bForms)
rs("bIntranet")	= convertBool(bIntranet)
rs("bIntranetSetup")	= convertBool(bIntranetSetup)
rs("bIntranetContacts")	= convertBool(bIntranetContacts)
rs("bIntranetMail")	= convertBool(bIntranetMail)
rs("bCatalog")	= convertBool(bCatalog)
rs("bFeed")	= convertBool(bFeed)
rs("bTheme")	= convertBool(bTheme)
rs("bFormExport")	= convertBool(bFormExport)
rs("bApplicationpath")	= convertBool(bApplicationpath)
rs("bGallery")	= convertBool(bGallery)
rs("bPagesAdd")	= convertBool(bPagesAdd)
rs("bPagesPW")	= convertBool(bPagesPW)
rs("bPagesMove")	= convertBool(bPagesMove)
rs("bPagesDelete")	= convertBool(bPagesDelete)
rs("bPageUrlRSS")	= convertBool(bPageUrlRSS)
rs("bPageAdditionalHeader")	= convertBool(bPageAdditionalHeader)
rs("bPageDescription")	= convertBool(bPageDescription)
rs("bPageKeywords")	= convertBool(bPageKeywords)
rs("bPageTitleTag")	= convertBool(bPageTitleTag)
rs("bPageRefresh")	= convertBool(bPageRefresh)
rs("bPageTemplate")	= convertBool(bPageTemplate)
rs("bPageTheme")	= convertBool(bPageTheme)
rs("bPageForm")	= convertBool(bPageForm)
rs("bPageCatalog")	= convertBool(bPageCatalog)
rs("bPageFeed")	= convertBool(bPageFeed)
rs("bPageSetHomepage")	= convertBool(bPageSetHomepage)
rs("bPagePublish")	= convertBool(bPagePublish)
rs("bPageOrder")	= convertBool(bPageOrder)
rs("bPageTitle")	= convertBool(bPageTitle)
rs("bPageBody")	= convertBool(bPageBody)
rs("bPageUFL")	= convertBool(bPageUFL)
rs("bPoll")	= convertBool(bPoll)
rs("bGuestbook")	= convertBool(bGuestbook)
rs("bPopup")	= convertBool(bPopup)
rs("bNewsletter")	= convertBool(bNewsletter)
rs("bAvailabilityCal")	= convertBool(bAvailabilityCal)
rs("bShoppingCart")	= convertBool(bShoppingCart)
rs("bPeel")	= convertBool(bPeel)
rs("bCookieWarning")	= convertBool(bCookieWarning)
rs("bCustom404") = convertBool(bCustom404)
if convertGetal(iCustomerID)<>convertGetal(cId) and convertGetal(iCustomerID)<>0 then
rs("iCustomerID")	= iCustomerID
else
rs("iCustomerID")	= cId
end if
rs.Update 
rs.close
Set rs = nothing
clearMenuCache
save=true
pick()
end function
Public sub getPasswordValues()
sPassword	= sha256(Request.Form ("sPassword"))
sPasswordConfirm	= sha256(Request.Form ("sPasswordConfirm"))
end sub
public sub getRequestValues()
bSetupGeneral	= convertBool(Request.Form ("bSetupGeneral"))
bSetupPageElements	= convertBool(Request.Form ("bSetupPageElements"))
bStats	= convertBool(Request.Form ("bStats"))
bTemplates	= convertBool(Request.Form ("bTemplates"))
bHomeConstants	= convertBool(Request.Form ("bHomeConstants"))
bHomeVBScript	= convertBool(Request.Form ("bHomeVBScript"))
bFiles	= convertBool(Request.Form ("bFiles"))
bForms	= convertBool(Request.Form ("bForms"))
bIntranet	= convertBool(Request.Form ("bIntranet"))
bIntranetSetup	= convertBool(Request.Form ("bIntranetSetup"))
bIntranetContacts	= convertBool(Request.Form ("bIntranetContacts"))
bIntranetMail	= convertBool(Request.Form ("bIntranetMail"))
bCatalog	= convertBool(Request.Form ("bCatalog"))
bFeed	= convertBool(Request.Form ("bFeed"))
bTheme	= convertBool(Request.Form ("bTheme"))
bFormExport	= convertBool(Request.Form ("bFormExport"))
bApplicationpath	= convertBool(Request.Form ("bApplicationpath"))
bGallery	= convertBool(Request.Form ("bGallery"))
bPagesAdd	= convertBool(Request.Form ("bPagesAdd"))
bPagesPW	= convertBool(Request.Form ("bPagesPW"))
bPagesMove	= convertBool(Request.Form ("bPagesMove"))
bPagesDelete	= convertBool(Request.Form ("bPagesDelete"))
bPageUrlRSS	= convertBool(Request.Form("bPageUrlRSS"))
bPageAdditionalHeader	= convertBool(Request.Form("bPageAdditionalHeader"))
bPageDescription	= convertBool(Request.Form("bPageDescription"))
bPageKeywords	= convertBool(Request.Form("bPageKeywords"))
bPageTitleTag	= convertBool(Request.Form("bPageTitleTag"))
bPageRefresh	= convertBool(Request.Form("bPageRefresh"))
bPageTemplate	= convertBool(Request.Form("bPageTemplate"))
bPageTheme	= convertBool(Request.Form("bPageTheme"))
bPageForm	= convertBool(Request.Form("bPageForm"))
bPageCatalog	= convertBool(Request.Form("bPageCatalog"))
bPageFeed	= convertBool(Request.Form("bPageFeed"))
bPageSetHomepage	= convertBool(Request.Form("bPageSetHomepage"))
bPagePublish	= convertBool(Request.Form("bPagePublish"))
bPageOrder	= convertBool(Request.Form("bPageOrder"))
bPageTitle	= convertBool(Request.Form("bPageTitle"))
bPageBody	= convertBool(Request.Form("bPageBody"))
bPageUFL	= convertBool(Request.Form("bPageUFL"))
bPoll	= convertBool(Request.Form ("bPoll"))
bGuestbook	= convertBool(Request.Form ("bGuestbook"))
bPopup	= convertBool(request.form("bPopup"))
bNewsletter	= convertBool(request.form("bNewsletter"))
bAvailabilityCal	= convertBool(request.form("bAvailabilityCal"))
bShoppingCart	= convertBool(request.form("bShoppingCart"))
bPeel	= convertBool(request.form("bPeel"))
bCookieWarning	= convertBool(request.form("bCookieWarning"))
bCustom404	= convertBool(request.form("bCustom404"))
end sub
public sub delete
dim rs
set rs=db.execute("delete from tblSecondAdmin where iCustomerID="& cId)
set rs=nothing
end sub
end class%>
