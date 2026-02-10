<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bGuestbook%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><%=getBOHeader(btn_gb)%><%dim guestbook, entry, entryID
set guestbook=new cls_guestbook
dim postBack
postBack=convertBool(Request.Form("postBack"))
if postBack then
checkCSRF()
for each entryID in Request.Form ("iGBItemID")
set entry=new cls_guestbookitem
entry.pick(entryID)
if convertGetal(entry.iId)<>0 then
entry.sValue	= Request.Form ("sValue" & entryID)
entry.sReply	= Request.Form ("sReply" & entryID)
entry.sMessageBy	= Request.Form ("sAuthor" & entryID)
entry.bApproved	= convertBool(Request.Form ("approve" & entryID))
if convertBool(Request.Form ("remove" & entryID)) then
entry.remove()
else
entry.save()
end if
end if
set entry=nothing
next
Message.Add("fb_saveOK")
end if
dim entries
set entries=guestbook.entries%><table align=center><tr><td align=center>-> <b><a href="bs_gbList.asp"><%=l("back")%></a></b> <-</td></tr></table><form action="bs_gbEditItems.asp" method="post" name="mainform"><input type=hidden name=iGBId value="<%=encrypt(guestbook.iID)%>" /><input type="hidden" value="<%=true%>" name=postback /><%=QS_secCodeHidden%><table align=center style="width:95%" border=1 cellpadding=5 cellspacing=0><tr><td colspan="7" align="center"><input class="art-button" onclick="javascript:this.disabled=true;mainform.submit()" name="btnaction" type="submit" value="<%=l("save")%>" /></td></tr><tr><td align="center">iId</td><td align="center">IP</td><td>Author</td><td>Message</td><td>Reply by admin</td><td align="center">Approved</td><td align="center">Delete</td></tr><%dim arrE
for each entry in entries
arrE=split(entries(entry),QS_VBScriptIdentifier)%><tr><td align="center"><%=entry%><input type="hidden" value="<%=entry%>" name="iGBItemID" /></td><td align="center"><%=sanitize(arrE(3))%></td><td><input type="text" size="20" maxlength="50" name="sAuthor<%=entry%>" value="<%=sanitize(arrE(0))%>" /><%if not isLeeg(arrE(5)) then%><br /><a href="mailto:<%=sanitize(arrE(5))%>"><%=sanitize(arrE(5))%></a><%end if%></td><td><textarea rows="4" cols="20" name="sValue<%=entry%>"><%=sanitize(arrE(1))%></textarea></td><td><textarea rows="4" cols="20" name="sReply<%=entry%>"><%=sanitize(arrE(4))%></textarea></td><td align="center"><input type="checkbox" value="<%=true%>" name="approve<%=entry%>" <%if convertBool(arrE(2)) then Response.Write " checked='checked' "%> /></td><td align="center"><input type="checkbox" value="<%=true%>" name="remove<%=entry%>" /></td></tr><%next%><tr><td colspan="7" align="center"><input class="art-button" onclick="javascript:this.disabled=true;mainform.submit()" name="btnaction" type="submit" value="<%=l("save")%>" /></td></tr></table></form><table align=center><tr><td align=center>-> <b><a href="bs_gbList.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
