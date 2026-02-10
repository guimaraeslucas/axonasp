<!-- #include file="begin.asp"-->


<%bUseArtLoginTemplate=true%><!-- #include file="bs_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%if request.querystring("loggedoff")="" then
logon.logoffBO
response.cookies.item(cId&"hfsdsiiqqssdfjf")=""
Session(cId & "isAUTHENTICATEDSecondAdmin")= false
Session(cId & "isAUTHENTICATED")= false
response.redirect(QS_backsite_login_page)
end if%><p align=center><%=l("youareloggedoff")%><br /><a href="<%=QS_backsite_login_page%>"><%=l("loginagain")%></a></p><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
