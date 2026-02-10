<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><%dim newsletterMailing,newsletter,ccEmail,allEmails
set allEmails=server.createobject("scripting.dictionary")
set newsletterMailing=new cls_newsletterMailing
set newsletter=newsletterMailing.newsletter
'count
dim NewsletterCategories,postback
set NewsletterCategories=customer.NewsletterCategories
dim arrC,c,total,iInterval,iIntervalS
if request("iInterval")<>"" then
iInterval=convertGetal(request("iInterval"))
iIntervalS=convertGetal(request.form("iForceReload"))
else
iInterval=10
iIntervalS=15
end if
dim iStart
iStart=convertGetal(request("iStart"))
total=0
arrC=split(newsletterMailing.sCategory,",")
postback=convertBool(request("postback"))
bAddImageToEmail=convertBool(newsletterMailing.bLog)
if postback then
for c=lbound(arrC) to ubound(arrC)
if NewsletterCategories.exists(convertGetal(arrC(c))) then
dim rs
set rs=db.execute("select * from tblNewsletterCategorySubscriber where bActive=" & getSQLBoolean(true)& " and iCategoryID=" & convertGetal(arrC(c)) & " order by iId asc")
if iStart=0 then
session("NLstartTime")="<p align='center'>Mailing started on: <b>" & convert2(hour(time)) & ":" & convert2(minute(time)) & ":" & convert2(second(time)) & "</b></p>"
end if
dim stopon,rsRunner
rsRunner=0
stopon=iStart+iInterval
while not rs.eof and iStart<stopon
if rsRunner>=iStart then
ccEmail=rs("sEmail")
if not allEmails.exists(ccEmail) then
if bAddImageToEmail then
dim rsI
set rsI=db.getDynamicRS
rsI.open "select * from tblNewsletterLog where 1=2"
rsI.AddNew()
rsI("iSubscriberID")	= rs("iId")
rsI("iMailingID")	= newsletterMailing.iId
rsI("bRead")	= false
rsI("dWhen")	= null
rsI("sKey")	= generatePassword
rsI.update()
sAddImageUrl=customer.sQSUrl & "/default.asp?pageAction=rlog&k=" & rsI("sKey") & "&i=" & rsI("iId")
rsI.close
set rsI=nothing
else
sAddImageUrl=""
end if
newsletter.send rs("sName"),ccEmail, rs("sKey")
allEmails.add ccEmail,""
end if
iStart=iStart+1
end if
rsRunner=rsRunner+1
rs.movenext
wend
if rs.eof then
session("NLstopTime")="<p align='center'>Mailing completed on: <b>" & convert2(hour(time)) & ":" & convert2(minute(time)) & ":" & convert2(second(time)) & "</b></p>"
if newsletterMailing.bLog then
newsletterMailing.dSentDate=now()
newsletterMailing.save()
end if
response.redirect ("bs_newsletterMailingSend.asp?iStart=" & iStart & "&sent=1&iNewsletterMailingID=" & encrypt(newsletterMailing.iID))
end if
end if
next
end if
for c=lbound(arrC) to ubound(arrC)
if NewsletterCategories.exists(convertGetal(arrC(c))) then
total=total+NewsletterCategories(convertGetal(arrC(c))).nmbrSubscribers
end if
next
set NewsletterCategories=nothing
set allEmails=nothing
if request("sent")<>"1" then
iForceReload=convertGetal(request("iForceReload"))
sMetaTagRefresh="<META HTTP-EQUIV=""refresh"" CONTENT=" & """" & iForceReload &";URL=bs_newsletterMailingSend.asp?iStart=" & iStart & "&iInterval=" & iInterval & "&sent=0&postback=1&iForceReload=" & iForceReload & "&iNewsletterMailingID=" & encrypt(newsletterMailing.iID) & """/>"
end if%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Newsletter)%><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterList.asp"><%=l("back")%></a></b> <-</td></tr></table><%select case request("sent")
case ""
if not postback then%><form action="bs_NewsletterMailingSend.asp" method="post" name="mainform"><input type=hidden name=iNewsletterMailingID value="<%=encrypt(newsletterMailing.iID)%>" /><input type=hidden name=postBack value="<%=true%>" /><input type=hidden name=sent value="0" /><%=QS_secCodeHidden%><p align="center">You are about to send the newsletter <strong><%=newsletterMailing.newsletter.sName%></strong> to <b><%=total%></b> subscribers</p><p align="center">Send <select name="iInterval"><%=numberList(1,100,1,iInterval)%></select> messages every <select name="iForceReload"><%=numberList(3,600,3,iIntervalS)%></select> seconds</p><p align="center"><input class="art-button" type=submit name="btnaction" onclick="javascript:if(confirm('Are you sure to send the mailing now?')) {this.value='Please wait... Do not close this window!';this.disabled=true;mainform.submit();}else{return false;}" value="Send now!" /></p></form><%end if
case "0"%><p align="center"><table align=center cellpadding=15><tr><td align=center style="background-color:Yellow;color:Green"><b><%=iStart%> messages</b> sent until now... <br /><br /><b>STILL SENDING!</b><br /><br /><b>DO NOT CLOSE THIS WINDOW!</b></table></p><%case "1"%><p align="center"><b>Mailing finished!</b> - <b><%=iStart%></b> messages are sent!</p><%=session("NLstartTime") & session("NLstopTime")%><%if newsletterMailing.bLog then %><p align="center"><a href="bs_newsletterMailingHistory.asp">Go to mailing reporting</a></p><%else
newsletterMailing.remove()
end if
set newsletterMailing=nothing%><%end select%><table align=center><tr><td align=center>-> <b><a href="bs_NewsletterList.asp"><%=l("back")%></a></b> <-</td></tr></table><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
