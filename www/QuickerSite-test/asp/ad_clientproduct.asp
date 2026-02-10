<!-- #include file="begin.asp"-->


<!-- #include file="beginClient.asp"--><!-- #include file="ad_security.asp"--><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><!-- #include file="ad_adminMenu.asp"--><%dim clientproduct
set clientproduct=new cls_clientproduct
if Request.Form ("btnaction")=l("save") then
clientproduct.iClientID	= convertGetal(decrypt(Request.Form ("iClientID")))
clientproduct.sName	= convertStr(Request.Form ("sName"))
clientproduct.iPrice	= convertGetal(Request.Form ("iPrice"))
clientproduct.iVat	= convertGetal(Request.Form ("iVat"))
clientproduct.iRenewal	= convertGetal(Request.Form ("iRenewal"))
clientproduct.sNotes	= convertStr(Request.Form ("sNotes"))
clientproduct.dLastRenewalDate	= convertDateFromPicker(Request.Form ("dLastRenewalDate"))
clientproduct.dStartService	= convertDateFromPicker(Request.Form ("dStartService"))
clientproduct.dEndService	= convertDateFromPicker(Request.Form ("dEndService"))
clientproduct.sInvoiceNr	= convertStr(Request.Form ("sInvoiceNr"))
if clientproduct.save() then
Response.Redirect ("ad_client.asp?iClientID=" & encrypt(clientproduct.iClientID))
end if
end if
if Request.Form ("btnaction")=l("delete") then
checkCSRF()
clientproduct.remove()
Response.Redirect ("ad_client.asp?iClientID=" & encrypt(clientproduct.iClientID))
end if
'iPrice,iVat,iRenewal,dLastRenewalDate,sNotes,%><form action="ad_clientproduct.asp" name="mainform" method="post"><%=QS_secCodeHidden%><input type=hidden name=iClientProductID value="<%=encrypt(clientproduct.iID)%>" /><%if convertGetal(clientproduct.iClientID)<>0 then%><p align=center><strong><a class="art-button" href="ad_client.asp?iClientID=<%=encrypt(clientproduct.iClientID)%>"><%=quotrep(clientproduct.client.sname)%></a></strong></p><%end if%><table cellpadding="2" align="center" border="0" style="border-style:none"><tr><td class=header colspan=2>Product info:</td></tr><tr><td class=QSlabel>Customer:*</td><td><select name="iClientID"><option value="">Select...</option><%=clientproduct.client.showSelected(clientproduct.iClientID)%></select></td></tr><tr><td class=QSlabel>Product:*</td><td><input type=text value="<%=quotRep(clientproduct.sName)%>" name=sName style="width:500px" maxlength=250 /></td></tr><tr><td class=QSlabel valign="top">Notes:</td><td><%createFCKInstance clientproduct.sNotes,"siteBuilderMailSource","sNotes"%></td></tr><tr><td class=QSlabel>Price:</td><td><input type=text value="<%=quotRep(clientproduct.iPrice)%>" name=iPrice style="width:50px" maxlength=50 /></td></tr><tr><td class=QSlabel>VAT:</td><td><input type=text value="<%=quotRep(clientproduct.iVat)%>" name=iVat style="width:30px" maxlength=50 />%</td></tr><tr><td class=QSlabel>Start service:</td><td><input type="text" id="dStartService" name="dStartService" value="<%=convertEuroDate(clientproduct.dStartService)%>" /><%=JQDatePicker("dStartService")%></td></tr><tr><td class=QSlabel>End service:</td><td><input type="text" id="dEndService" name="dEndService" value="<%=convertEuroDate(clientproduct.dEndService)%>" /><%=JQDatePicker("dEndService")%></td></tr><tr><td class=QSlabel>Renew every </td><td><select name=iRenewal><%=numberList(0,96,1,clientproduct.iRenewal)%></select> months (0: not renewable service)</td></tr><tr><td class=QSlabel>Last renewed on:</td><td><input type="text" id="dLastRenewalDate" name="dLastRenewalDate" value="<%=convertEuroDate(clientproduct.dLastRenewalDate)%>" /><%=JQDatePicker("dLastRenewalDate")%></td></tr><tr><td class=QSlabel>(Last) Invoice Number:</td><td><input type=text value="<%=quotRep(clientproduct.sInvoiceNr)%>" name=sInvoiceNr style="width:130px" maxlength=50 /></td></tr><tr><td class=QSlabel>&nbsp;</td><td>(*) <%=l("mandatory")%></td></tr><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type="submit" value="<%=l("save")%>" name="btnaction" /><%if isNumeriek(clientproduct.iId)  then%>&nbsp;<input class="art-button" onclick="javascript: return confirm('<%=l("areyousure")%>');" type="submit" value="<%=l("delete")%>" name="btnaction"  /><%end if%></td></tr></table></form><%if convertGetal(clientproduct.iId)=0  then%><script type="text/javascript">document.mainform.sName.focus();</script><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
