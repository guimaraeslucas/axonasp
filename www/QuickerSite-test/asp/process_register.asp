
<%if Request.Form ("pageAction") = cRegister then
checkCSRF()
set ticket=new cls_ticket
ticket.sEmail=trim(Request.Form ("sEmail"))
if ticket.saveAndSend() then
Response.Redirect ("default.asp?pageAction="& cloginIntranet&"&iId="&encrypt(getIntranetHomePage.iId)&"&fbMessage=fb_activationlink")
end if
end if
pageBody=pageBody& "<form action='default.asp' method='post' name='registerForm'>"
pageBody=pageBody& QS_secCodeHidden
pageBody=pageBody& "<input type='hidden' name='pageAction' value='"& cRegister & "' />"
pageBody=pageBody & customer.sExplTicket
pageBody=pageBody& "<div id='QS_form'>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>"&l("email")&":</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input required style=""width:200px"" size='30' maxlength='100' type='email' name='sEmail' value=" & """" & sanitize(Request("sEmail")) & """" &" /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody & "<div class='QS_fieldline'>"
pageBody=pageBody&"<div class='QS_fieldlabel'>"& l("captcha") &":</div>"
pageBody=pageBody & "<div class='QS_fieldvalue'><img src='" & C_DIRECTORY_QUICKERSITE & "/asp/includes/captcha.asp' alt='This is a CAPTCHA Image' style='vertical-align: middle;border-style:none;margin:0px 6px 0px 0px' /><input type='text' size='5' style='width:63px' maxlength='4' name='captcha' required /></div>"
pageBody=pageBody & "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>&nbsp;</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input class=""art-button"" type='submit' value="""& sanitize(l("send")) &""" name='dummy' /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "</form>"
pageBody=pageBody& "<script type=""text/javascript"">document.registerForm.sEmail.focus()</script>"
pageTitle=customer.sLabelRegister%>
