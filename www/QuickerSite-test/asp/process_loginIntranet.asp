
<%if Request.Form ("pageAction") = cloginIntranet then
logon.mode=Request.Form ("mode")
If LCase(session("captcha")) <> LCase(Left(Request.Form("captcha"),4)) Then
message.AddError("err_captcha")
elseif logon.logonIntranet(Request.form("sEmail"),Request.form("sPw")) then
'response.end
'buildURL!
if not isLeeg(request("iPostID")) then
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId)&"&iPostID=" & sanitize(request("iPostID")) & "&item="& sanitize(request("item")))
else
if request("item")="" then
if customer.bUserfriendlyURL and not isleeg(selectedPage.sUserfriendlyURL) then
Response.Redirect (selectedPage.sUserfriendlyURL)
end if
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId))
else
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId) & "&item="& sanitize(request("item")))
end if
end if
else
message.AddError("err_login")
end if
end if
dim strChecked
if isLeeg(Request.Form ("mode")) then
strChecked=" checked='checked' "
end if
pageTitle=customer.intranetName
pageBody=pageBody& "<form action='default.asp' method='post' name='loginIntranetForm'>"
pageBody=pageBody& "<input type='hidden' name='iId' value='" & encrypt(selectedPage.iId) & "' />"
pageBody=pageBody& "<input type='hidden' name='item' value="""& sanitize(request("item")) &""" />"
pageBody=pageBody& "<input type='hidden' name='iPostID' value=""" & sanitize(Request("iPostID")) & """ />"
pageBody=pageBody& "<input type='hidden' name='pageAction' value='"&cloginIntranet&"' /><p>"&l("toGetAccessTo")&" <b>" & customer.intranetName & "</b>, "
pageBody=pageBody& l("needPWIntranet") &"</p>"
pageBody=pageBody& "<div id='QS_form'>"
pageBody=pageBody& "<div class='QS_fieldline'>"
select case convertGetal(customer.iLoginMode)
case 0
pageBody=pageBody& "<div class='QS_fieldlabel'>"&l("email")&":</div>"
case 1
pageBody=pageBody& "<div class='QS_fieldlabel'>"&l("nickname")&":</div>"
end select
pageBody=pageBody& "<div class='QS_fieldvalue'><input required type='text' style=""width:200px"" name='sEmail' size='30' value=" & """" & sanitize(Request("sEmail")) & """" &" /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>"&l("password")&":</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input type='password' style=""width:200px"" name='sPw' size='30' value='' required /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>"&l("captcha")&":</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><img src='" & C_DIRECTORY_QUICKERSITE & "/asp/includes/captcha.asp' style='vertical-align: middle;margin:0px 6px 0px 0px;border-style:none' alt='captcha image' /><input style=""width:63px"" type='text' name='captcha' size='6' maxlength='4' required /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>&nbsp;</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><a href='default.asp?pageAction="& cForgotPW &"'>"&l("forgotPW")&"</a>"
if customer.bAllowNewRegistrations then
pageBody=pageBody&"<br />"
pageBody=pageBody& "<a href='default.asp?pageAction="& cRegister &"'>"&customer.sLabelRegister&"</a>"
end if
pageBody=pageBody& "</div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>&nbsp;</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input class=""art-button"" type='submit' value='"&l("login")&"' name='dummy' /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>&nbsp;</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input type='radio' "
if not isLeeg(strChecked) then
pageBody=pageBody& strChecked
else
pageBody=pageBody& convertChecked(Request ("mode")="mode1") 
end if
pageBody=pageBody& " name='mode' value='mode1' />&nbsp;" & l("b1") & "</div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>&nbsp;</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input type='radio' " & convertChecked(Request ("mode")="mode2") & " name='mode' value='mode2' />&nbsp;"&l("b2") & "</div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>&nbsp;</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input type='radio' " & convertChecked(Request ("mode")="mode3") & " name='mode' value='mode3' />&nbsp;"&l("b3") & "</div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "</form>"
pageBody=pageBody& "<script type=""text/javascript"">document.loginIntranetForm.sEmail.focus();</script>"%>
