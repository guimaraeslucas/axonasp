<!-- #include file="asp/begin.asp"-->
<!-- #include file="asp/includes/rss_writer.asp"-->
<%
Response.ContentType="application/xml"
'response.end 
dim rss, keyLink, formattedDateTime
set rss = new kwRSS_writer	
		
if convertGetal(selectedPage.iId)=0 then
	
	if customer.bEnableMainRSS then

		if isEmpty(application("RSS" & cId)) or application("RSS" & cId)="" then

			'retrieve a list of recently updated pages
			rss.ChannelTitle = customer.siteName
			rss.ChannelURL   = customer.sUrl
			rss.ChannelDesc  = customer.sDescription
			rss.ChannelLanguage= langCode
	
			dim rs, RSScounter, rssPage
			RSScounter=0
			set rs=db.execute("select iId from tblPage where " & sqlCustId & " order by updatedTS desc")
	
			while not rs.eof and RSScounter<300
	
				set rssPage=new cls_page
				rssPage.pick(rs(0))
				
				if not isLeeg(rssPage.iId) and rssPage.bOnline and not rssPage.bDeleted and not rssPage.bContainerPage and not rssPage.bLossePagina and isLeeg(rssPage.iListPageID) and not rssPage.bIntranet and isLeeg(rssPage.sApplication) and isLeeg(rssPage.sPw) and isLeeg(rssPage.sExternalURL) then
				
					rss.AddNew		
					rss.SetTitle rssPage.sTitle
					rss.SetLink customer.sQSUrl & "/default.asp?iID="& encrypt(rssPage.iId)
					rss.SetDesc  rssPage.sValue				
					Format_RFC822_DateAndTime rssPage.updatedTS,formattedDateTime							
					rss.setPubdate formattedDateTime
					rss.SetAuthor  customer.webmasteremail & " (" & customer.webmaster &")"
					rss.setGUID customer.sQSUrl & "/default.asp?iID="& encrypt(rssPage.iId)
	
					RSScounter=RSScounter+1
					
				end if
				
				set rssPage=nothing
				
				rs.movenext
				
			wend 
	
			application("RSS" & cId)=treatConstants(prepareForExport(rss.GetRSS),true)
	
			set rs=nothing
	
		end if
	
		Response.Write application("RSS" & cId)
		
	end if

else

	if selectedPage.bIntranet and not logon.authenticatedIntranet then		
		
		rss.ChannelTitle = selectedPage.sTitle & " - " & customer.siteName
		rss.ChannelURL   = customer.sUrl
		rss.ChannelDesc  = customer.sDescription
		rss.ChannelLanguage= langCode
		rss.AddNew		
		rss.SetTitle "No Access"
		rss.SetLink customer.sUrl							
		rss.SetDesc  "No Access"								
		Format_RFC822_DateAndTime date(),formattedDateTime								
		rss.setPubdate formattedDateTime
		rss.SetAuthor  customer.webmasteremail & " (" & customer.webmaster &")"
		rss.setGUID generatePassword
		Response.Write rss.GetRSS
		set rss=nothing
	
	else

		
		'caching in application
		if isEmpty(application("RSS"&cId&selectedPage.iId & Request.QueryString ("viewtype"))) or application("RSS"&cId&selectedPage.iId & Request.QueryString ("viewtype"))="" or not isLeeg(Request.QueryString ("viewtype")) or 1=1 then
			
			rss.ChannelTitle = selectedPage.sTitle & " - " & customer.siteName
			rss.ChannelURL   = customer.sUrl
			rss.ChannelDesc  = customer.sDescription
			rss.ChannelLanguage= langCode

			if selectedPage.bPushRSS then

				dim rssItems, listkey
				set rssItems=selectedPage.fastlistitems
				for each listkey in rssItems
				
					if customer.bUserFriendlyURL and not isLeeg(rssItems(listkey).sUserfriendlyURL) then
						keyLink=customer.sQSUrl & "/" &  rssItems(listkey).sUserfriendlyURL
					else
						keyLink=customer.sQSUrl & "/default.asp?iID="& encrypt(selectedPage.iID) & "&amp;item="& encrypt(listkey) &"#"& encrypt(listkey) 
					end if
					
					rss.AddNew		
					rssItems(listkey).bHideDate=selectedPage.bHideDate			
					
					rss.SetTitle rssItems(listkey).sDateAndTitle
					
					if not isLeeg(rssItems(listkey).sLPExternalURL) and rssItems(listkey).sLPExternalURL<>"http://"  then
						rss.SetLink rssItems(listkey).sLPExternalURL
					else
						rss.SetLink	 keyLink
					end if
					
					rss.SetDesc  prepareForExport(rssItems(listkey).sValue &  rssItems(listkey).feed.build)
					
					if not isLeeg(rssItems(listkey).sItemPicture) then
						rss.SetEnclosure customer.sVDUrl & Application("QS_CMS_userfiles") & "listitemimages/" & listkey & "." & rssItems(listkey).sItemPicture
					end if
					
					Format_RFC822_DateAndTime rssItems(listkey).updatedTS,formattedDateTime
					
					rss.setPubdate formattedDateTime
					rss.SetAuthor  customer.webmasteremail & " (" & customer.webmaster &")"
					rss.setGUID keyLink
								
				next
				
				set rssItems=nothing

			end if
			
			
			if convertGetal(selectedPage.iThemeID)<>0 then
			
				dim theme
				set theme=selectedpage.theme
				
				if theme.bOnline and theme.bPushRSS then		
				
					theme.startpage=1
					theme.iPageSize=30
					
					dim posts,postkey
					set posts=theme.posts
						
					for each postkey in posts

						keyLink=customer.sQSUrl & "/default.asp?iID="& encrypt(selectedPage.iID) & "&iPostID=" & encrypt(postkey)

						rss.AddNew		
						rss.SetTitle quotrep(posts(postkey).sSubject)
						rss.SetLink	 keyLink
						
						if theme.bAllowHTML then
							rss.SetDesc prepareForExport(filterJS(addSmilies(posts(postkey).sBody)))
						else
							rss.SetDesc addSmilies(linkUrls(quotrep(posts(postkey).sBody)))
						end if	
						
						Format_RFC822_DateAndTime posts(postkey).dUpdatedTS,formattedDateTime
						
						rss.setPubdate formattedDateTime
						rss.SetAuthor  posts(postkey).contact.sNickName
						rss.setGUID keyLink
						rss.setComments keyLink & "#comments"
									
					next
				
					set posts=nothing
				
				end if
				
			end if

			dim copyCatalog
			set copyCatalog=selectedPage.catalog

			if convertGetal(copyCatalog.iId)<>0 and copyCatalog.bOnline then

				dim resultItems, itemSearch, itemKey, fileTypes, searchFields
				set searchFields=copyCatalog.fields("public,search")	
				
				set itemSearch=new cls_itemSearch	
				set itemSearch.searchFields=searchFields
				itemSearch.bOnline=true
				itemSearch.iCatalogID=copyCatalog.iId	

				set resultItems=itemSearch.resultItems	
													
				for each itemKey in resultItems
																			
					keyLink=customer.sQSUrl & "/default.asp?iID="& encrypt(selectedPage.iID) & "&amp;item="& encrypt(itemKey) &"#"& encrypt(itemKey) 

					rss.AddNew		
					rss.SetTitle resultItems(itemKey).sDateAndTitle
					rss.SetLink	 keyLink
					rss.SetDesc  prepareForExport(resultItems(itemKey).sFiche)
					
					
					
					Format_RFC822_DateAndTime resultItems(itemKey).dUpdatedTS,formattedDateTime
					
					rss.setPubdate formattedDateTime
					rss.SetAuthor  customer.webmasteremail & " (" & customer.webmaster &")"
					rss.setGUID keyLink		
												
				next
				
				set searchFields	= nothing
				set itemSearch		= nothing
				set resultItems		= nothing				
				
			end if

			set copyCatalog=nothing
			
			application("RSS"&cId&selectedPage.iId & Request.QueryString ("viewtype"))=rss.GetRSS
			
			Set rss = nothing
			
			
		end if

		response.write treatConstants(application("RSS"&cId&selectedPage.iId & Request.QueryString ("viewtype")),true)

		selectedPage.addHitRSS

		logReferer()
		
	end if

end if

cleanUPASP()
%>