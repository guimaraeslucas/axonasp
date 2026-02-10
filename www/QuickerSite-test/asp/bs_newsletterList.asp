<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Newsletter)%><%dim copyNewsletter
set copyNewsletter=new cls_Newsletter
copyNewsletter.copy()
if convertGetal(copyNewsletter.iId)<>0 then Response.Redirect ("bs_NewsletterEdit.asp?iNewsletterID=" & encrypt(copyNewsletter.iId))
dim NewsletterCategories
set NewsletterCategories=customer.NewsletterCategories
dim Newsletters
set Newsletters=customer.Newsletters%><p align=center><a class="art-button" href="bs_NewsletterEdit.asp"><b>New Newsletter</b></a> 
<%if customer.bCanImportSubscribers then%>  <a class="art-button" href="bs_NewsletterMailingEdit.asp"><b>New Mailing</b></a>  <a class="art-button" href="bs_NewsletterMailingHistory.asp"><b>Mailing History</b></a>  <a class="art-button" href="bs_NewsletterSubscribers.asp"><b>Subscribers</b></a> 
 <%end if%> <a class="art-button" href="bs_NewsletterCategoryList.asp"><b>Email lists</b></a></p><%if Newsletters.count>0 then%><table align=center cellpadding=3 cellspacing=0><%dim NewsletterKey, Newsletter
for each NewsletterKey in Newsletters%><tr><td style="border-top:1px solid #DDD"><a href="bs_NewsletterEdit.asp?iNewsletterID=<%=encrypt(Newsletterkey)%>"><%=Newsletters(Newsletterkey).sName%></a></td><td style="border-top:1px solid #DDD"><%=getIcon(l("copyItem"),"copyItem","bs_NewsletterList.asp?"&QS_secCodeURL&"&amp;iNewsletterID="&encrypt(Newsletterkey),"javascript:return confirm('"& l("areyousuretocopy") &"');","copy"&Newsletterkey)%></td><td style="border-top:1px solid #DDD"><i>iID: <%=Newsletterkey%></i></td></tr><%next%></table><%else%><p align=center>No Newsletters available</p><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
