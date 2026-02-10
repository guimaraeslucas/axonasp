<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Newsletter)%><%dim newsletterCategory
set newsletterCategory=new cls_newsletterCategory
dim postBack
postBack=convertBool(request.form("postBack"))
if request.form("btnaction")=l("delete") then
newsletterCategory.remove
Response.Redirect ("bs_NewsletterCategoryList.asp")
end if
if postBack then
newsletterCategory.getRequestValues()
checkCSRF()
if newsletterCategory.save then 
message.Add("fb_saveOK")
end if
end if%><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterCategoryList.asp"><%=l("back")%></a></b> <-</td></tr></table><form action="bs_NewsletterCategoryEdit.asp" method="post" name="mainform"><input type=hidden name=iNewsletterCategoryID value="<%=encrypt(newsletterCategory.iID)%>" /><input type=hidden name=postBack value="<%=true%>" /><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td class="QSlabel">List name:*</td><td colspan="2"><input type="text" size="40" maxlength="50" name="sName" value="<%=quotrep(newsletterCategory.sName)%>" /></td></tr><tr><td class="QSlabel">Signup form:*</td><td colspan="2"><textarea cols="90" name="sSignupForm" rows="20"><%=quotrep(newsletterCategory.sSignupForm)%></textarea></td></tr><tr><td class="QSlabel">Require both name and email?</td><td colspan="2"><input type="checkbox" <%if newsletterCategory.bRequireBoth then response.write  " checked='checked' " %> value="<%=true%>" name="bRequireBoth" /></td></tr><tr><td class="QSlabel">Error message:*</td><td colspan="2"><input maxlength="250" type="text" size="90" name="sErrorMessage" value="<%=quotrep(newsletterCategory.sErrorMessage)%>" /></td></tr><tr><td class="QSlabel">Email notification for new signups/unsubscribes to:<br />(leave blanco if not needed)</td><td colspan="2"><input type="text" size="40" maxlength="50" name="sNotifEmail" value="<%=quotrep(newsletterCategory.sNotifEmail)%>" /></td></tr><tr><td class="QSlabel">Welcome message:*<br />(shown on website after registration)</td><td colspan="2"><%createFCKInstance newsletterCategory.sWelcomeMessage,"siteBuilderRichText","sWelcomeMessage"%></td></tr><tr><td class="QSlabel">Unsubscribe-feedback page-title:*</td><td colspan="2"><input type=text size=50 maxlength=50 name="sUnsubscribeFBTitle" value="<%=quotRep(newsletterCategory.sUnsubscribeFBTitle)%>" /></td></tr><tr><td class="QSlabel">Unsubscribe-feedback:*<br />(shown on website after unsubscribe)</td><td colspan="2"><%createFCKInstance newsletterCategory.sUnsubscribeFB,"siteBuilderRichText","sUnsubscribeFB"%></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit name="btnaction" value="<% =l("save")%>" /> 
<%if isNumeriek(newsletterCategory.iID) then%><input class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>" /><%end if%></td></tr></table></form><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterCategoryList.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
