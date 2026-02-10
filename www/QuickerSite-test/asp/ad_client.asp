<!-- #include file="begin.asp"-->


<!-- #include file="beginClient.asp"--><!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="ad_adminMenu.asp"--><%dim client,products
set client=new cls_client
set products=client.products
if Request.Form ("btnaction")=l("save") then
client.sName	= convertStr(Request.Form ("sName"))
client.sAddress	= convertStr(Request.Form ("sAddress"))
client.sContactPerson	= convertStr(Request.Form ("sContactPerson"))
client.sMainEmail	= convertStr(Request.Form ("sMainEmail"))
client.sOtherEmail	= convertStr(Request.Form ("sOtherEmail"))
if client.save() then
message.Add("fb_saveOK")
end if
end if
if Request.Form ("btnaction")=l("delete") then
checkCSRF()
client.remove()
Response.Redirect ("ad_clientList.asp")
end if%><form action="ad_client.asp" name="mainform" method="post"><%=QS_secCodeHidden%><input type=hidden name=iClientId value="<%=encrypt(client.iID)%>" /><table cellpadding="2" align="center" border="0" style="border-style:none"><tr><td valign="top"><table align="center" border="0" style="border-style:none"><tr><td class=header colspan=2><%=l("general")%>:</td></tr><tr><td class=QSlabel>Name:*</td><td><input type=text value="<%=quotRep(client.sName)%>" name=sName style="width:300px" maxlength=250 /></td></tr><tr><td class=QSlabel>Contact Person:*</td><td><input type=text value="<%=quotRep(client.sContactPerson)%>" name=sContactPerson style="width:300px" maxlength=250 /></td></tr><tr><td class=QSlabel>Main email:*</td><td><input type=text value="<%=quotRep(client.sMainEmail)%>" name=sMainEmail style="width:300px" maxlength=250 /></td></tr><tr><td class=QSlabel valign="top">Other email:</td><td><textarea name="sOtherEmail" style="width:300px"  rows=3><%=quotRep(client.sOtherEmail)%></textarea></td></tr><tr><td class=QSlabel valign="top">Full Address:</td><td><textarea name="sAddress" style="width:300px"  rows=8><%=quotRep(client.sAddress)%></textarea></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type="submit" value="<%=l("save")%>" name="btnaction" /><%if isNumeriek(client.iId)  then%>&nbsp;<input class="art-button" onclick="javascript: return confirm('<%=l("areyousure")%>');" type="submit" value="<%=l("delete")%>" name="btnaction" /><%end if%></td></tr></table></td><td valign=top><%if isNumeriek(client.iId)  then%><table cellpadding="2" align="center" border="0" style="border-style:none"><tr><td class=header>Products:</td></tr><%if products.count>0 then%><tr><td><ul><%dim p
for each p in products%><li><a href="ad_clientproduct.asp?iClientProductID=<%=encrypt(p)%>"><%=products(p).sName%></a></li><%next%></ul></td></tr><%end if%><tr><td><%=getArtLink("ad_clientproduct.asp?iClientID=" & encrypt(client.iID),"Add product","","","")%></td></tr></table><%end if%></td></tr></table></form><%if convertGetal(client.iId)=0  then%><script type="text/javascript">document.mainform.sName.focus();</script><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
