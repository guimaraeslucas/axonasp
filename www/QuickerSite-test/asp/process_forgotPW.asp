
<%if Request.Form ("pageAction") = cForgotPW then
If LCase(session("captcha")) <> LCase(Left(Request.Form("captcha"),4)) Then
message.AddError("err_captcha")
elseif logon.resetPW(Request.form("sEmail")) then
Response.Redirect ("default.asp?pageAction="& cloginIntranet&"&iId="&encrypt(getIntranetHomePage.iId)&"&fbMessage=fb_emailFound")
else
message.AddError("err_emailNotFound")
end if
end if
pageTitle=l("forgotPW")
pageBody=pageBody & "<form action='default.asp' method='post' name='forgotPW'>"
pageBody=pageBody & "<input type='hidden' name='pageAction' value='" & cForgotPW &"' />"
pageBody=pageBody & "<p>" & l("forgotPwExpl") & "</p>"
pageBody=pageBody & "<div id='QS_form'>"
pageBody=pageBody & "<div class='QS_fieldline'>"
pageBody=pageBody & "<div class='QS_fieldlabel'>"&l("email")&":</div>"
pageBody=pageBody & "<div class='QS_fieldvalue'><input required type='email' style=""width:200px"" name='sEmail' size='30' value="& """" & sanitize(Request.Form ("sEmail")) &"""" &" /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody & "<div class='QS_fieldline'>"
pageBody=pageBody & "<div class='QS_fieldlabel'>"&l("captcha")&":</div>"
pageBody=pageBody & "<div class='QS_fieldvalue'><img src='" & C_DIRECTORY_QUICKERSITE & "/asp/includes/captcha.asp' style='vertical-align: middle;border-style:none;margin:0px 6px 0px 0px'><input style=""width:63px"" type='text' name='captcha' size='6' maxlength='4' required /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody & "<div class='QS_fieldline'>"
pageBody=pageBody & "<div class='QS_fieldlabel'>&nbsp;</div>"
pageBody=pageBody & "<div class='QS_fieldvalue'><input class=""art-button"" type='submit' value=" & """" & sanitize(l("send")) & """" &" name='dummy' /></div>"
pageBody=pageBody & "</div>"
pageBody=pageBody & "</div>"
pageBody=pageBody & "</form>"
pageBody=pageBody& "<script type=""text/javascript"">document.forgotPW.sEmail.focus()</script>"%>
