<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%dim selectedColor
selectedColor=request.form("sCWLinkColor")
if isLeeg(selectedColor) then
selectedColor=request.querystring("selectedColor")
end if
if isLeeg(selectedColor) then
selectedColor="#000000"
end if
if request.querystring("loaddark")="1" then
customer.sQSAccordionMain="font-family: inherit; " & vbcrlf & "font-size: 1em;"
customer.sQSAccordionHeader="border:1px solid #111111;" & vbcrlf & "background-color: #222222;" & vbcrlf & "color: #FFFFFF;" & vbcrlf & "font-weight: 700;"
customer.sQSAccordionContent="border:1px solid #111111;" & vbcrlf & "background-color: #666666;" & vbcrlf & "color: #FFFFFF;"
customer.save()
response.redirect("bs_AccordionSetup.asp?fbMessage=fb_saveOK")
end if
if request.querystring("loaddefault")="1" then
customer.sQSAccordionMain	= ""
customer.sQSAccordionHeader	= ""
customer.sQSAccordionContent	= ""
customer.save()
response.redirect("bs_AccordionSetup.asp?fbMessage=fb_saveOK")
end if
if request.form("suggestcss")="Suggest CSS" then
dim color1,color2,color3,helpColor,triad
color1=replace(selectedColor,"#","",1,-1,1)
helpColor=HEXCOL2RGB(color1,0)
triad=16777216/3
color2=triad+helpColor
if color2>16777216 then color2=color2-16777216
helpColor=color2
color2=left(hex(color2),6)
if len(color2)<6 then color2=left(color2,3)
if len(color2)<4 then color2=color2&color2&color2&color2&color2&color2
color2=left(color2,6)
color3=helpColor+triad
if color3>16777216 then color3=color3-16777216
color3=left(hex(color3),6)
if len(color3)<6 then color3=left(color3,3)
if len(color3)<4 then color3=color3&color3&color3&color3&color3&color3
color3=left(color3,6)
customer.sQSAccordionHeader="border:1px solid #" & color1 & ";" & vbcrlf & "background-color: #" & color1 & ";" & vbcrlf & "color: #" & color2 & ";" & vbcrlf & "font-weight: 700;"
customer.sQSAccordionContent="border:1px solid #" & color1 & ";" & vbcrlf & "background-color: #" & color2 & ";" & vbcrlf & "color: #" & color3 & ";"
customer.save()
response.redirect("bs_AccordionSetup.asp?selectedColor=" & server.urlencode(selectedColor) & "&fbMessage=fb_saveOK")
end if
if request.form("btnaction")<>"" then
customer.sQSAccordionMain	= trim(request.form("sQSAccordionMain"))
customer.sQSAccordionHeader	= trim(request.form("sQSAccordionHeader"))
customer.sQSAccordionContent	= trim(request.form("sQSAccordionContent"))
customer.save()
response.redirect("bs_AccordionSetup.asp?fbMessage=fb_saveOK")
end if%><script type="text/javascript">$(function() {$( "#accordionPREVIEWQS" ).accordion();});</script><p  align="center">Below you can customize the look and feel of the JQuery Accordion plugin (used in list pages)</p><form method="post" action="bs_AccordionSetup.asp" name="mainform"><table align="center" cellpadding="3"><tr><td valign="top" class=QSlabel>Main CSS</td><td><textarea cols="80" rows="7" name="sQSAccordionMain"><%=quotrep(customer.sQSAccordionMain)%></textarea></td><td valign="top" rowspan=4 style="width:300px"><a class="art-button" href="bs_AccordionSetup.asp?loaddefault=1" onclick="javascript:return confirm('Are you sure to load the default CSS?');">Load default CSS</a><a class="art-button" href="bs_AccordionSetup.asp?loaddark=1" onclick="javascript:return confirm('Are you sure to load the dark CSS?');">Load dark CSS</a><br /><br /><input type="text" id="sCWLinkColor" name="sCWLinkColor" value="<%=selectedColor%>" /><%=JQColorPicker("sCWLinkColor")%>&nbsp;
<input class="art-button" type="submit" class="submit" name="suggestcss" value="Suggest CSS" /><p>Preview:</p><div class="QSAccordion" id="accordionPREVIEWQS"><div class="QSAccordionHeader">Lorum Ipsum 1</div><div class="QSAccordionContent"><p>Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua...</p></div><div class="QSAccordionHeader">Lorum Ipsum 2</div><div class="QSAccordionContent"><p>Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua...</p></div><div class="QSAccordionHeader">Lorum Ipsum 3</div><div class="QSAccordionContent"><p>Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua...</p></div></div></td></tr><tr><td valign="top" class=QSlabel>Header CSS</td><td><textarea cols="80" rows="7" name="sQSAccordionHeader"><%=quotrep(customer.sQSAccordionHeader)%></textarea></td></tr><tr><td valign="top" class=QSlabel>Content CSS</td><td><textarea cols="80" rows="7" name="sQSAccordionContent"><%=quotrep(customer.sQSAccordionContent)%></textarea></td></tr><tr><td valign="top" class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit value="<%=l("save")%>" name="btnaction" /></td></tr></table></form><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
