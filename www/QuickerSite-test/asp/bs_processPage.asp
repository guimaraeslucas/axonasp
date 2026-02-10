
<script type="text/javascript">function deletePage(){
if (confirm('<%=l("deletecomplete")%>')){
document.mainform.btnaction.value='<%=l("delete")%>';
document.mainform.submit ();
}
}
</script><%if isNumeriek(Decrypt(Request("iParentID"))) then
page.iParentID=Decrypt(Request("iParentID"))
end if
Select Case request("btnaction")
case l("save"),save_listpage,save_listitem
checkCSRF()
page.getRequestValues()
if page.save then
if request("btnaction")=l("save") then
message.Add("fb_saveOK")
end if
if request("btnaction")=save_listpage then
Response.Redirect ("bs_listpage.asp?iID="& EnCrypt(page.iID))
end if
if request("btnaction")=save_listitem then
Response.Redirect ("bs_listpage.asp?iID="& EnCrypt(page.iListPageID))
end if
end if
case l("delete")
checkCSRf()
page.bDeleted=true
page.parentPage.removeRang(page)
page.iRang=0
page.bOnline=false
page.bHomepage=false
page.deleteListItems()
if page.save then
if isLeeg(page.iListPageID) then
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
else
Response.Redirect ("bs_listPage.asp?fbMessage=fb_topicremoved&iID="&encrypt(page.iListPageID))
end if
end if
case l("save_pw")
checkCSRF()
page.sPw=Request.Form ("sPw")
if len(Request.Form ("sPw"))<3 then
Message.AddError("err_pw")
else
page.sPw=lcase(convertStr(Request.Form ("sPw")))
if page.save then
page.resetAllSubPasswords(page.iId)
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
end if
end if
case l("delete_pw")
checkCSRF()
page.sPw=""
if page.save then
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
end if
case l("delete_pw_all")
checkCSRF()
page.sPw=""
if page.save then
page.removeAllSubPasswords(page.iId)
if page.bIntranet then
Response.Redirect ("bs_intranet.asp")
else
Response.Redirect ("bs_default.asp")
end if
end if
end select%>
