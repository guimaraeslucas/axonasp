<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bCookieWarning%><!-- #include file="includes/header.asp"--><!-- #include file="includes/urlenCodeJS.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%dim postback
postback=convertBool(request.form("postback"))
if request.querystring("loaddefault")=1 then
select case convertGetal(customer.language)
case 2 'Nederlands
customer.sCWLocation="top"
customer.sCWNumber="0"
customer.sCWText="<p><strong>Cookie beleid</strong>&nbsp;-&nbsp;Deze website gebruikt cookies. Enkele van deze cookies zijn van belang om delen van deze website goed laten functioneren en zijn reeds aangemaakt.<br />U kan deze cookies blokkeren en verwijderen, maar dan zullen bepaalde pagina's op deze website niet werken. Lees meer over ons <a href=""" & C_DIRECTORY_QUICKERSITE & "/cookiepolicy/cookiebeleid.html"" class=""QSPP"">cookie beleid</a>.</p>"
customer.sCWContinue="Accepteren"
customer.sCWError="U dient de ""Ik accepteer de cookies van deze website"" box aanvinken"
customer.sCWAccept="Ik accepteer de cookies van deze website"
customer.sCWBackgroundColor="#000000"
customer.sCWButtonClass="art-button"
customer.sCWTextColor="#FFFFFF"
customer.sCWLinkColor="#FFFFFF"
customer.bCWUseAsNormalPP=false
case else
customer.sCWLocation="top"
customer.sCWNumber="0"
customer.sCWText="<p><strong>Cookies Policy</strong>&nbsp;-&nbsp;This site uses cookies. Some of the cookies we use are essential for parts of the site to operate and have already been set.<br />You may delete and block all cookies from this site, but parts of the site will not work. To find out more about cookies on this website, see our <a href=""" & C_DIRECTORY_QUICKERSITE & "/cookiepolicy/cookiespolicy.html"" class=""QSPP"">cookies policy</a>.</p>"
customer.sCWContinue="Continue"
customer.sCWError="You must tick the ""I accept cookies from this site"" box to accept"
customer.sCWAccept="I accept cookies from this site"
customer.sCWBackgroundColor="#000000"
customer.sCWButtonClass="art-button"
customer.sCWTextColor="#FFFFFF"
customer.sCWLinkColor="#FFFFFF"
customer.bCWUseAsNormalPP=false
end select
if customer.save() then
response.redirect("bs_cookiewarning.asp?fbMessage=fb_saveOK")
end if
end if
if request.querystring("loadminimal")="1" then
select case convertGetal(customer.language)
case 2 'Nederlands
customer.sCWLocation="top"
customer.sCWNumber="0"
customer.sCWText=""
customer.sCWContinue="Accepteren"
customer.sCWError="U dient de ""Ik accepteer de cookies van deze website"" box aanvinken"
customer.sCWAccept="Deze website gebruikt cookies. Lees meer over ons <a href=""" & C_DIRECTORY_QUICKERSITE & "/cookiepolicy/cookiebeleid.html"" class=""QSPP"">cookie beleid</a>. Ik accepteer de cookies van deze website"
customer.sCWBackgroundColor="#000000"
customer.sCWButtonClass=""
customer.sCWTextColor="#FFFFFF"
customer.sCWLinkColor="#FFFFFF"
customer.bCWUseAsNormalPP=false
case else
customer.sCWLocation="top"
customer.sCWNumber="0"
customer.sCWText=""
customer.sCWContinue="Continue"
customer.sCWError="You must tick the ""I accept cookies from this site"" box to accept"
customer.sCWAccept="This site uses cookies. To find out more about cookies on this website, see our <a href=""" & C_DIRECTORY_QUICKERSITE & "/cookiepolicy/cookiespolicy.html"" class=""QSPP"">cookies policy</a>. I accept cookies from this site"
customer.sCWBackgroundColor="#000000"
customer.sCWButtonClass=""
customer.sCWTextColor="#FFFFFF"
customer.sCWLinkColor="#FFFFFF"
customer.bCWUseAsNormalPP=false
end select
if customer.save() then
response.redirect("bs_cookiewarning.asp?fbMessage=fb_saveOK")
end if
end if
if postback then
customer.bCookieWarning=convertBool(request.form("bCookieWarning"))
customer.sCWLocation=request.form("sCWLocation")
customer.sCWNumber=convertGetal(request.form("sCWNumber"))
customer.sCWText=request.form("sCWText")
customer.sCWContinue=request.form("sCWContinue")
customer.sCWError=request.form("sCWError")
customer.sCWAccept=request.form("sCWAccept")
customer.sCWBackgroundColor=request.form("sCWBackgroundColor")
customer.sCWButtonClass=request.form("sCWButtonClass")
customer.sCWTextColor=request.form("sCWTextColor")
customer.sCWLinkColor=request.form("sCWLinkColor")
customer.bCWUseAsNormalPP=convertBool(request.form("bCWUseAsNormalPP"))
if customer.save() then
response.redirect("bs_cookiewarning.asp?fbMessage=fb_saveOK")
end if
end if%><p align="center"></p><form method="post" action="bs_cookiewarning.asp" name="mainform"><input type="hidden" name="postback" value="1" /><table align="center" cellpadding="2"><tr><td class=QSlabel valign=top>Enable cookie warning?</td><td><input name="bCookieWarning" type="checkbox" value="1" <%if convertBool(customer.bCookieWarning) then response.write "checked=""checked"""%> /><span style="float:right"><input type="submit" value="Save" class="art-button" />&nbsp;
<a class="art-button"  href="bs_cookieWarning.asp?loadminimal=1" onclick="javascript:return confirm('Are you sure to load the MINIMAL settings/texts for the cookie warning? This will overwrite/erase the current settings!')">load minimal settings</a>&nbsp;
<a class="art-button"  href="bs_cookieWarning.asp?loaddefault=1" onclick="javascript:return confirm('Are you sure to load the DEFAULT settings/texts for the cookie warning? This will overwrite/erase the current settings!')">load default settings</a>&nbsp;
<a target="_blank" class="art-button" href="<%=C_DIRECTORY_QUICKERSITE & "/default.asp?iId="& EnCrypt(getHomepage.iId)%>&amp;QSbcwp=1">Preview</a></span></td></tr><tr><td class=QSlabel valign=top>Warning position:</td><td><select name="sCWLocation"><option value="top" <%if isLeeg(customer.sCWLocation) or convertStr(customer.sCWLocation)="top" then response.write " selected=""selected"""%>>top</option><option value="bottom" <%if convertStr(customer.sCWLocation)="bottom" then response.write " selected=""selected"""%>>bottom</option></select></td></tr><tr><td class=QSlabel valign=top>Nmbr of displays:</td><td><select name="sCWNumber"><%=numberList(0,20,1,customer.sCWNumber)%></select> (0: the warning will keep on showing)</td></tr><tr><td class=QSlabel>Disclosure</td><td><%=dumpFCKInstance(customer.sCWText,"siteBuilderMailSource","sCWText")%></textarea></td></tr><tr><td class=QSlabel>Text-color:</td><td><input type="text" id="sCWTextColor" name="sCWTextColor" value="<%=quotrep(customer.sCWTextColor)%>" /><%=JQColorPicker("sCWTextColor")%></td></tr><tr><td class=QSlabel>Link-color:</td><td><input type="text" id="sCWLinkColor" name="sCWLinkColor" value="<%=quotrep(customer.sCWLinkColor)%>" /><%=JQColorPicker("sCWLinkColor")%></td></tr><tr><td class=QSlabel>Background-color:</td><td><input type="text" id="sCWBackgroundColor" name="sCWBackgroundColor" value="<%=quotrep(customer.sCWBackgroundColor)%>" /><%=JQColorPicker("sCWBackgroundColor")%></td></tr><tr><td class=QSlabel>Do not show acceptance-block(*)</td><td><input type="checkbox" value="1" name="bCWUseAsNormalPP" <%if convertBool(customer.bCWUseAsNormalPP) then response.write "checked=""checked"""%> /></td></tr><tr><td class=QSlabel>Acceptance:</td><td><input type="text" maxlength=250 size="80" value="<%=quotrep(customer.sCWAccept)%>" name="sCWAccept" /></td></tr><tr><td class=QSlabel>Continue:</td><td><input type="text" maxlength=50 size="20" value="<%=quotrep(customer.sCWContinue)%>" name="sCWContinue" /> class=<input type="text" name="sCWButtonClass" maxlength=50 value="<%=quotrep(customer.sCWButtonClass)%>" /></td></tr><tr><td class=QSlabel>Error:</td><td><input type="text" maxlength=250 size="80" value="<%=quotrep(customer.sCWError)%>" name="sCWError" /></td></tr><tr><td class=QSlabel>&nbsp;</td><td><i>(*) not showing the acceptance-block makes this behave as an eye-catcher, rather than a cookie warning.</i></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit value="<%=l("save")%>" name="btnaction" /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
