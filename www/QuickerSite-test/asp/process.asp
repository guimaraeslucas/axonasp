
<%if not devVersion then
if customer.adminPassword=sha256(QS_defaultPW) then
response.redirect("asp/" & QS_backsite_login_page)
end if
end if
'mobile browsers?
checkMobileBrowser()
'UFL?
if isNull(selectedPage.iId) then
selectedPage.pickByUserFriendlyURL()
end if
'variabelen declareren
dim pageTitle, pageBody, pagePrint, pageMail, pageSiteMap, theMail, pageZoomIn, pageZoomOut, pageListBody, pageZoomBack, pageSearch
dim search, catalogItem, body, formFields, fFieldKey, formItem, siteMap, siteMapContent
dim contactFields, contactField, fields, results, clickmenu, resultPage, listitems, listkey, encryptedListkey
dim catalog, catalogFields, catalogField, catalogFormFields, searchFields, sf, ifields, files, f
dim fileTypes, ft, itemSearch, inputFields, resultItems, itemKey, postBackCat, formDict, formKey, dateFields,  catForm
dim smTable,subPages, subPage,continue,gotoMenu,onlineSiblings,siblingmenu,hereList, isLPEU, selectedListItem,poll
dim pageListPage, pageCatalog, pageFeed, pageForm, pageRSS, runAPP, searchSnippet,cRSSlink,ticket,pageTheme,LPCounter


if selectedPage.b404 then

	pageTitle=convertStr(customer.sCustom404Title)
	pageTitle=replace(pageTitle,"[404FILENAME]",sanitize(selectedPage.sufl),1,-1,1)
	selectedPage.sTitle=pageTitle
	pageBody=convertStr(customer.sCustom404Body)
	pageBody=replace(pageBody,"[404FILENAME]",sanitize(selectedPage.sufl),1,-1,1)
	i404TemplateID=convertGetal(customer.i404TemplateID)
	Response.Status = "404 File Not Found"
	selectedPage.buildTemplate()
	
else

	runAPP=false
	postBackCat=convertBool(Request("postBackCat"))
	set search=new cls_search

	select case lcase(convertStr(request("pageAction")))

	case "avataredit"
	response.clear%><!-- #include file="process_avatar.asp"--><%response.flush
	CleanupASP()
	response.end 
	case "fileupload"
	response.clear%><!-- #include file="process_fileupload.asp"--><%response.flush
	CleanupASP()
	response.end 
	case "unsubscribe"%><!-- #include file="process_unsubscribe.asp"--><%case "send"%><!-- #include file="process_send.asp"--><%case "showpost"%><!-- #include file="process_showpost.asp"--><%case "showitem"%><!-- #include file="process_showitem.asp"--><%case "sitemap"%><!-- #include file="process_sitemap.asp"--><%case "itemform"%><!-- #include file="process_itemform.asp"--><%case "search"%><!-- #include file="process_search.asp"--><%case "nomobile","mobile","wap","nowap"
	Response.Redirect ("default.asp?iId=" & encrypt(selectedPage.iId))
	case "nlcatshowfb"
		
	dim nlcat
	set nlcat=new cls_newsletterCategory
	response.clear
	nlcat.registerNameAndEmail()
	response.flush
	CleanupASP()
	response.end
	case "vote"
	set poll=new cls_poll
	poll.registerVote()
	set poll=nothing
	Response.Flush
	CleanupASP()
	Response.end
	case "voteshowresults"
	set poll=new cls_poll
	Response.write poll.showresults()
	set poll=nothing
	Response.Flush
	CleanupASP()
	Response.end
	case "voteshowballot"
	set poll=new cls_poll
	Response.write poll.build(false)
	set poll=nothing
	Response.Flush
	CleanupASP()
	Response.end
	case "showpopup"
	dim showPopup
	set showPopup=new cls_popup
	response.clear
	response.write "<!DOCTYPE html><html><head><meta charSet=""utf-8""/><title>"  & quotrep(treatconstants(showPopup.sName,true)) &  "</title></head><body style=""margin:0px 0px 0px 0px;padding:0px 0px 0px 0px"">" & treatconstants(showPopup.sValue,true) & "</body></html>"
	set showPopup=nothing
	response.flush
	CleanupASP()
	Response.end
	case "rlog"
	response.clear
	if convertGetal(left(request.querystring("i"),12))<>0 then
	dim rsUL
	set rsUL=db.getDynamicRS
	rsUL.open "select * from tblNewsletterLog where iId=" & convertGetal(request.querystring("i")) & " and sKey='" & cleanup(left(request.querystring("k"),8)) & "'"
	if not rsUL.eof then
	rsUL("bRead")=true
	rsUL("dWhen")=now()
	rsUL.update
	rsUL.close
	end if
	set rsUL=nothing
	end if
	CleanupASP()
	Response.end
	case "gbedit" 
	dim gbEntry
	select case Request.QueryString ("ac")
	case "approve"
	set gbEntry=new cls_guestbookitem
	gbEntry.pickByKey(Request.QueryString ("sKey"))
	gbEntry.approve()
	set gbEntry=nothing
	Response.redirect("default.asp?fbMessage=fb_saveOK&iId=" & Request.QueryString ("redPage")) 
	case "delete"
	set gbEntry=new cls_guestbookitem
	gbEntry.pickByKey(Request.QueryString ("sKey"))
	gbEntry.remove()
	set gbEntry=nothing
	Response.redirect("default.asp?fbMessage=fb_topicremoved&iId=" & Request.QueryString ("redPage")) 
	case "edit"
	set gbEntry=new cls_guestbookitem
	gbEntry.pickByKey(Request.QueryString ("sKey"))
	if convertGetal(gbEntry.iId)<>0 then 
	Response.Redirect ("asp/" & QS_backsite_login_page & "?bs_page="& server.URLEncode ("bs_gbEditItems.asp?iGBID=" & encrypt(gbEntry.iGuestBookID)))
	end if
	end select
	case "editsite"
	if customer.intranetUse and logon.authenticatedIntranet then%><!-- #include file="process_editsite.asp"--><%else
	redirectToHP()
	end if
	case lcase(clogin)%><!-- #include file="process_login.asp"--><%case cWelcome
	if customer.intranetUse and customer.intranetUseMyProfile  then
	pageTitle=l("thankyou")
	pageBody=customer.sWelcomeMessage
	else
	redirectToHP()
	end if
	case cProfile
	if customer.intranetUse and customer.intranetUseMyProfile then%><!-- #include file="process_profile.asp"--><%else
	redirectToHP()
	end if
	case lcase(cloginIntranet)
	if customer.intranetUse then%><!-- #include file="process_loginIntranet.asp"--><%else
	redirectToHP()
	end if
	case lcase(cForgotPW)
	if customer.intranetUse then%><!-- #include file="process_forgotPW.asp"--><%else
	redirectToHP()
	end if
	case cRegister
	if customer.intranetUse and customer.bAllowNewRegistrations then%><!-- #include file="process_register.asp"--><%else
	redirectToHP()
	end if
	case lcase(cLogOff)
	checkCSRF()
	logon.logoff
	if convertGetal(logon.contact.iID)<>0 then logon.contact.lastLogoutTSSave
	if convertGetal(decrypt(request.querystring("iId")))<>0 then
	response.redirect("default.asp?iID="&request.querystring("iId"))
	else
	response.redirect("default.asp?iID="&EnCrypt(getHomepage.iId))
	end if
	case else
	dim getTperP,getBperP
	if request.querystring("fpv")="1" then
	if logon.authenticatedIntranet then
	set getTperP=logon.contact.getTper
	set getBperP=logon.contact.getBper
	if getTPerP.exists(selectedPage.iId) or getBPerP.exists(selectedPage.iId) then
	if selectedPage.sTitleToBeValidated<>selectedPage.sTitle and not isLeeg(selectedPage.sTitleToBeValidated) then
	selectedPage.sTitle=selectedPage.sTitleToBeValidated
	end if
	if selectedPage.sValueToBeValidated<>selectedPage.sValue and not isLeeg(selectedPage.sValueToBeValidated) then
	selectedPage.sValue=selectedPage.sValueToBeValidated
	end if
	end if
	set getTperP=nothing
	set getBperP=nothing
	end if
	end if
	dim hackofflinepages
	hackofflinepages=false

	'set online list items online!
	selectedPage.setOnlineListItemsOnline

	if not selectedPage.bOnline  then
	if Session(cId & "isAUTHENTICATEDSecondAdmin") or Session(cId & "isAUTHENTICATED") then
	hackofflinepages=true
	end if
	if not hackofflinepages then
	if logon.authenticatedIntranet then
	set getTperP=logon.contact.getTper
	set getBperP=logon.contact.getBper
	if getTPerP.exists(selectedPage.iId) or getBPerP.exists(selectedPage.iId) then
	hackofflinepages=true
	end if
	end if
	end if
	if hackofflinepages then
	selectedPage.sValue=selectedPage.sValue&"<script type=""text/javascript"">alert('This page is offline. It is not visible to the public.')</script>"
	end if
	end if
	set getTperP=nothing
	set getBperP=nothing
	if not (isNumeriek(selectedPage.iId) and (selectedPage.bOnline or editPage or hackofflinepages) and not selectedPage.bDeleted) then
	if right(Request.ServerVariables ("http_querystring"),12)<>"/favicon.ico" then 
	set selectedPage=getHomepage
	selectedPage.showFromCache()
	else
	cleanUPASP
	Response.End
	end if
	end if
	if isNumeriek(selectedPage.iId) and (selectedPage.bOnline or editPage or hackofflinepages) and not selectedPage.bDeleted then
	'check password
	if logon.logonItem(Request.Cookies(encrypt(selectedPage.iId)),selectedPage) then
	if selectedPage.bIntranet then
	if not logon.authenticatedIntranet then
	Response.Redirect ("default.asp?iPostID=" & request("iPostID") & "&pageAction="&cloginIntranet&"&iId="&encrypt(selectedPage.iId))
	else
	if not convertBool(customer.intranetUse) then
	Response.Redirect ("default.asp")
	end if
	if logon.contact.iStatus=cs_silent or logon.contact.iStatus=cs_profile or logon.contact.iStatus=cs_write then
	Response.Redirect ("default.asp?pageAction="&cWelcome)
	end if
	end if
	end if
	if selectedPage.bContainerPage then
	if not editPage then selectedPage.moveToFirstSubItem()
	elseif not isLeeg(selectedPage.sExternalURL) then
	if not editPage then Response.Redirect (selectedPage.sExternalURLPrefix&selectedPage.sExternalURL)
	else
	pageTitle=quotRep(selectedPage.sGetTitle)
	pageBody=selectedPage.listitemPicIMGTag & selectedPage.sValue
	if not isLeeg(selectedPage.sOrderBY) then%><!-- #include file="process_listpage.asp"--><%end if
	'form weergeven?
	if isNumeriek(selectedPage.iFormID) then
	pageForm=selectedPage.form.build("default.asp?iId="& encrypt(selectedPage.iID),selectedPage.sFormAlign,"submit",null)
	end if
	'thema weergeven?
	if isNumeriek(selectedPage.iThemeID) then
	pageTheme=selectedPage.theme.build()
	end if
	'cataloog weergeven?
	if isNumeriek(selectedPage.iCatalogID) then%><!-- #include file="process_catalog.asp"--><%end if
	'feed weergeven?
	if isNumeriek(selectedPage.iFeedID) then
	pageFeed=selectedPage.Feed.build()
	end if
	'mogen applications uitgevoerd worden?
	runAPP=true
	end if 'einde inhoudspagina
	else
	'pagina met paswoord
	if customer.bUserFriendlyURL and not isLeeg(getHomepage.sUserFriendlyURL) then
	'Response.Redirect (customer.sUrl&"/"&getHomepage.sUserFriendlyURL)
	Response.Redirect (customer.sQSUrl &"/default.asp?pageAction=" & clogin & "&iId="& encrypt(selectedPage.iId))
	else
	Response.Redirect (C_DIRECTORY_QUICKERSITE&"/default.asp?pageAction=" & clogin & "&iId="& encrypt(selectedPage.iId))
	end if
	end if
	end if
	end select
	selectedPage.AddHit()
	selectedPage.addVisit()
	pageSiteMap	= getIcon(l("sitemap"),"sitemap","default.asp?pageAction=sitemap","","sitemap")
	pageSearch	= getIcon(l("search"),"search","#","javascript:document.mainform.submit();","search")
	cRSSlink=selectedPage.RSSLink
	if not isLeeg(cRSSlink) then pageRSS = "<a target='RSS' href='"& treatConstants(cRSSlink,true) &"' title='RSS'><img border=0 alt='RSS' src='"&C_DIRECTORY_QUICKERSITE&"/fixedImages/public/feed.gif' /></a>"
	if convertBool(customer.bMonitor) then
	'customer.AddLog()
	end if
	set cPopup=customer.currentPopup
	if not printReplies then
	selectedPage.buildTemplate()
	elseif pagetoemail then
	pagetoemailbody=selectedPage.buildTemplate()
	else
	'response.write pagebody
	'do nothing
	end if
	pageBody=pageBody&pageListPage&pageForm&pageCatalog&pageFeed&pageTheme
	
end if
%>
