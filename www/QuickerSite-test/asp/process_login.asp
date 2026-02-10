
<%if Request.Form ("pageAction") = clogin then
if logon.logonItem(Request.form("sPw"),selectedPage) then
Response.Redirect ("default.asp?iId="& encrypt(selectedPage.iId))
else
message.AddError("err_login")
end if
end if
pageBody=pageBody& "<form action='default.asp' method='post' name='loginForm'>"
pageBody=pageBody& "<input type='hidden' name='iId' value='" & encrypt(selectedPage.iId) & "' />"
pageBody=pageBody& "<input type='hidden' name='pageAction' value='"& clogin & "' />"
pageBody=pageBody& "<p>" & l("toGetAccessTo") &" <b>" & selectedPage.sTitle & "</b>, "& l("needPW") &" -> <i>"
pageBody=pageBody& "<a href=" & """" & "mailto:"& customer.webmasterEmail &"?subject=" & l("requestPW") & "&amp;body=" & l("pleasePW")& " " & selectedPage.sTitle & " " & l("on") & " " & customer.siteName & "? " & l("thankyou") & "." & """" & ">"& l("askPassword") &"</a></i></p>"
pageBody=pageBody& "<div id='QS_form'>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>"&l("password")&":</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input style=""width:200px"" size='8' maxlength='15' type='password' name='sPw' value='' required /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "<div class='QS_fieldline'>"
pageBody=pageBody& "<div class='QS_fieldlabel'>&nbsp;</div>"
pageBody=pageBody& "<div class='QS_fieldvalue'><input class=""art-button"" type='submit' value="""& sanitize(l("login")) &""" name='dummy' /></div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "</div>"
pageBody=pageBody& "</form>"
pageBody=pageBody& "<script type=""text/javascript"">document.loginForm.sPw.focus();</script>"
pageTitle=selectedPage.sTitle%>
