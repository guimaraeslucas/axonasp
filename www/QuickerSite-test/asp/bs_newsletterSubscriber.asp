<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Newsletter)%><p align=center><%if customer.bCanImportSubscribers then%><a href="bs_NewsletterSubscribers.asp"><b>Subscribers</b></a> - <%end if%><a href="bs_NewsletterList.asp"><b>Newsletter home</b></a></p><%dim NewsletterCategories,iCategoryID
set NewsletterCategories=customer.NewsletterCategories
dim rs
set rs=db.getDynamicRS
rs.open "select * from tblNewsletterCategorySubscriber where iCustomerID=" & cId & " and iId=" & convertGetal(decrypt(request("iSubscriptionID")))
if rs.recordcount=0 then response.redirect ("bs_newsletterSubscribers.asp")
iCategoryID=rs("iCategoryID")
if request.form("btnaction")=l("save") then
rs("sEmail")=lcase(trim(request.form("sEmail")))
rs("sName")=trim(request.form("sName"))
rs("bActive")=convertBool(request.form("bActive"))
if not checkEmailSyntax(request.form("sEmail")) then
message.AddError("err_email")
else
rs.update()
rs.close()
set rs=nothing
response.redirect ("bs_newsletterSubscribers.asp?iCategoryID=" & encrypt(iCategoryID))
end if
end if
if request.form("btnaction")=l("delete") then
db.execute ("delete from tblNewsletterLog where iSubscriberID=" & convertGetal(decrypt(request("iSubscriptionID"))))
db.execute ("delete from tblNewsletterCategorySubscriber where iCustomerID=" & cId & " and iId=" & convertGetal(decrypt(request("iSubscriptionID"))))
response.redirect ("bs_newsletterSubscribers.asp?iCategoryID=" & encrypt(iCategoryID))
end if%><form name=mainform method=post action="bs_newsletterSubscriber.asp"><input type=hidden name=iSubscriptionID value="<%=encrypt(rs("iId"))%>" /><input type=hidden name=postBack value="<%=true%>" /><%=QS_secCodeHidden%><table align=center cellpadding="2"><tr><td class="QSlabel">Name:</td><td><input type=text size=40 maxlength=50 name="sName" value="<%=quotRep(rs("sName"))%>" /></td></tr><tr><td class="QSlabel">Email:*</td><td><input type=email size=40 maxlength=50 name="sEmail" value="<%=quotRep(rs("sEmail"))%>" /></td></tr><tr><td class="QSlabel">Active?</td><td><input type=checkbox name="bActive" <%if convertBool(rs("bActive")) then response.write " checked='checked' "%> value="<%=true%>" /></td></tr><tr><td class="QSlabel">Category:</td><td><%=NewsletterCategories(convertGetal(rs("iCategoryID"))).sName%></td></tr><tr><td class="QSlabel">Key:</td><td><%=quotRep(rs("sKey"))%></td></tr><tr><td class="QSlabel">Date added:</td><td><%=convertEurodate(rs("dAdded"))%></td></tr><tr><td class="QSlabel">&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class="QSlabel">&nbsp;</td><td><input class="art-button" type=submit value="<%=l("save")%>" name="btnaction" /> <input class="art-button" onclick="javascript:return confirm('Are you sure to delete this subscription?')" type=submit value="<%=l("delete")%>" name="btnaction" /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
