<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Newsletter)%><%dim newsletterMailing
set newsletterMailing=new cls_newsletterMailing
if request.form("btnaction")=l("delete") then
newsletterMailing.remove
Response.Redirect ("bs_NewsletterList.asp")
end if
dim previewName,previewEmail,postBack
postBack=convertBool(request.form("postBack"))
if postBack then
newsletterMailing.getRequestValues()
checkCSRF()
if newsletterMailing.save then 
previewName=request.form("previewName")
previewEmail=request.form("previewEmail")
if request.form("btnaction")=l("send") then
newsletterMailing.Newsletter.send previewName,previewEmail,""
else
response.redirect("bs_newsletterMailingSend.asp?iNewsletterMailingID=" & encrypt(newsletterMailing.iID))
end if
end if
end if
if isLeeg(previewName) or isLeeg(previewEmail) then
previewName=customer.webmaster
previewEmail=customer.webmasterEmail
end if%><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterList.asp"><%=l("back")%></a></b> <-</td></tr></table><form action="bs_NewsletterMailingEdit.asp" method="post" name="mainform"><input type=hidden name=iNewsletterMailingID value="<%=encrypt(newsletterMailing.iID)%>" /><input type=hidden name=postBack value="<%=true%>" /><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td colspan="2" class="header"><%=l("general")%>:</td></tr><tr><td class="QSlabel">Select Newsletter:*</td><td colspan="2"><select name="iNewsletterID"><%=customer.showSelectedNewsletter("option",newsletterMailing.iNewsletterID)%></select></td></tr><tr><td class="QSlabel">Select email list:*</td><td style="padding:4px"><%dim NewsletterCategories, NC
set NewsletterCategories=customer.NewsletterCategories
dim arrNC,arrNCi
arrNC=split(newsletterMailing.sCategory,",")
for each NC in NewsletterCategories
response.write "<input name=""sCategory"" type=""radio"" "
if NewsletterCategories.count=1 then
response.write " checked='checked' "
end if
for arrNCi=lbound(arrNC) to ubound(arrNC)
if convertGetal(NC)=convertGetal(trim(arrNC(arrNCi))) then
response.write " checked='checked' "
exit for 
end if
next
response.write " value=""" & NC & """/>" & NewsletterCategories(NC).sName & " (" & NewsletterCategories(NC).nmbrSubscribers & " subscribers)<br/>"
next%></td></tr><tr><td class=QSlabel>Store mailing for future reporting?</td><td><input type="checkbox" name="bLog" value="<%=true%>" <%if newsletterMailing.bLog then response.write " checked='checked' "%>/></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel><a name="MS"></a>Send preview to:</td><td><%if request.form("btnaction")=l("send") then%><span style="padding:5px;background-color:Yellow;color:Green"><strong>Preview was sent!</strong></span>  <%end if%><input type="text" name="previewName" value="<%=previewName%>" /> - <input type="text" name="previewEmail" value="<%=previewEmail%>" /><input class="art-button" type=submit name=btnaction value="<% =l("send")%>" /></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input  class="art-button" type=submit name="btnaction" value="<% =l("continue")%>"> 
<%if isNumeriek(newsletterMailing.iID) then%><input  class="art-button" type=submit name=btnaction onclick="javascript:return confirm('<%=l("areyousure")%>');" value="<% =l("delete")%>"><%end if%></td></tr></table></form><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterList.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
