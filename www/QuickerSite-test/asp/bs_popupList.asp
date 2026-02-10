<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bPopup%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Popup)%><%dim copyPopup
set copyPopup=new cls_popup
copyPopup.copy()
if convertGetal(copyPopup.iId)<>0 then Response.Redirect ("bs_popupEdit.asp?ipopupID=" & encrypt(copypopup.iId))
if convertGetal(decrypt(request.querystring("i")))<>0 then
copyPopup.pick(decrypt(request.querystring("i")))
copyPopup.iShows=0
copyPopup.save()
end if
set copyPopup=nothing
dim popups
set popups=customer.popups%><p align=center><%=getArtLink("bs_popupEdit.asp","New Popup","","","")%></p><%if popups.count>0 then%><table align=center cellpadding=4 cellspacing=0><%dim popupKey, popup
for each popupKey in popups%><tr><td style="border-top:1px solid #DDD"><a href="bs_popupEdit.asp?ipopupID=<%=encrypt(popupkey)%>"><%=popups(popupkey).sName%></a></td><td style="border-top:1px solid #DDD">(<%=convertGetal(popups(popupkey).iShows)%> times shown - <a href="bs_PopupList.asp?i=<%=encrypt(popupKey)%>">reset</a>)</td><td style="border-top:1px solid #DDD;text-align:center"><%if convertBool(popups(popupkey).bEnabled) then%><strong>Enabled</strong><%else%><i><font color=#AAA>Disabled</font></i><%end if%> </td><td style="border-top:1px solid #DDD"><%=getIcon(l("preview"),"search","#","javascript:window.open('" & C_DIRECTORY_QUICKERSITE & "/default.asp?iId="&encrypt(getHomePage.iId)&"&amp;forcePP="&popupkey&"','"&encrypt(popupkey)&"');","search"&popupkey)%></td><td style="border-top:1px solid #DDD"><%=getIcon(l("copyItem"),"copyItem","bs_popupList.asp?"&QS_secCodeURL&"&amp;ipopupID="&encrypt(popupkey),"javascript:return confirm('"& l("areyousuretocopy") &"');","copy"&popupkey)%></td><td style="border-top:1px solid #DDD"><i>iID: <%=popupkey%></i></td></tr><%next%></table><%else%><p align=center>No Popups available</p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
