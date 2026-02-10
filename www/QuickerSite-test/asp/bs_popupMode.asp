<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%'logon.hasaccess secondAdmin.bTemplates%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%'=getBOHeader(btn_Setup)%><%'=getBOSetupMenu(btn_template)%><%if request.form("sPopupViewmode")<>"" then
customer.sPopupViewmode=request.form("sPopupViewmode")
customer.save()
response.redirect("bs_popupMode.asp?fbMessage=fb_saveOK")
end if%><p  align="center">Select <b>Popup-effect:</b></p><form method="post" action="bs_popupMode.asp" name="popupForm"><table align="center"><tr><td><input type=radio value=1 name=sPopupViewmode <%if customer.sPopupViewmode="1" then response.write "checked='checked'"%> />Dark Rounded Square<br /><br /><input type=radio value=2 name=sPopupViewmode <%if customer.sPopupViewmode="2" then response.write "checked='checked'"%> />Light Square<br /><br /><input type=radio value=3 name=sPopupViewmode <%if customer.sPopupViewmode="3" then response.write "checked='checked'"%> />Dark Square<br /><br /><input type=radio value=4 name=sPopupViewmode <%if customer.sPopupViewmode="4" then response.write "checked='checked'"%> />Light Rounded Square<br /><br /><input type=radio value=5 name=sPopupViewmode <%if customer.sPopupViewmode="5" then response.write "checked='checked'"%> />Gray Square<br /><br /></td></tr><tr><td><input class="art-button" type=submit value="<%=l("save")%>" name="btnaction" /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
