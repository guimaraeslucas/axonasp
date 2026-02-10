
<%set catalogItem=new cls_catalogItem
catalogItem.pick(decrypt(Request.QueryString ("iItemID")))
catalogItem.checkOnline()
pageBody=pageBody&"<p class=""QScatItemBack""><a class=""QScatItemBack"" href=""default.asp?iId=" & sanitize(Request.QueryString("iPageID")) & "#" & encrypt(catalogItem.iId) &"""><- " & l("back")  &"</a></p>"
pageBody=pageBody&"<p>" & catalogItem.catalog.sItemName &": <b>" & catalogItem.sTitle & "</b></p>"
pageBody=pageBody & catalogItem.catalog.form.build("default.asp?pageAction=itemform&iPageID=" & sanitize(Request.QueryString("iPageID")) & "&iItemID=" & encrypt(catalogItem.iId),"left","submit",catalogItem.iId)
pageTitle=catalogItem.sTitle & " - " & catalogItem.catalog.sFormTitle
set catalogItem=nothing%>
