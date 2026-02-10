<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bIntranetContacts%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Intranet)%><%=getBOHeaderIntranet(btn_contacts)%><%dim contactFields, contactField
set contactFields=customer.contactFields(true)
if contactFields.count=0 then
set contactFields=customer.contactFields(false)
end if
dim contactSearch
set contactSearch=new cls_contactSearch
set contactSearch.searchFields=contactFields
contactSearch.getRequestValues()
dim inputFields
set inputFields=contactSearch.inputFields
dim dateFields
set dateFields=contactSearch.dateFields
dim lowerGreaterThanList
set lowerGreaterThanList=new cls_lowerGreaterThanList
dim resultTable
resultTable=contactSearch.resultTable%><p align=center><%if contactFields.count=0 then set contactFields=customer.contactFields(false)
if contactFields.count>0 then%><a href="bs_contactEdit.asp" class="bPopupFullWidthReload"><b><%=l("addcontact")%></b></a> - 
<%end if%><a href="bs_contactFields.asp"><b><%=l("specifyfields")%></b></a> - 
<a href="bs_tickets.asp"><b><%=l("pendingactivationlinks")%></b></a></p><form name="mainform" action="bs_contactHome.asp" method="post"><input type="hidden" name="postback" value="<%=true%>" /><table align=center cellpadding="2"><%for each contactField in contactFields%><tr><td class=QSlabel><%=contactFields(contactField).sFieldname%></td><td><%select case contactFields(contactField).sType
case sb_text,sb_textarea,sb_url,sb_email,sb_richtext%><input type=text size=15 name="<%=encrypt(contactField)%>" value="<%=quotRep(inputFields(contactField))%>" /><%case sb_checkbox%><input type=checkbox name="<%=encrypt(contactField)%>" value="checked" <%=inputFields(contactField)%> /><%case sb_select%><select name="<%=encrypt(contactField)%>"><option value=""></option><%=contactFields(contactField).showSelected(inputFields(contactField))%></select><%case sb_date%><i><%=l("between")%></i> 
<input type="text" id="from<%=encrypt(contactField)%>" name="from<%=encrypt(contactField)%>" value="<%=dateFields("from"&contactField)%>" /><%=JQDatePicker("from" & encrypt(contactField))%><i><%=l("and")%></i> 
<input type="text" id="untill<%=encrypt(contactField)%>" name="untill<%=encrypt(contactField)%>" value="<%=dateFields("untill"&contactField)%>" /><%=JQDatePicker("untill" & encrypt(contactField))%><%end select%></td></tr><%next%><tr><td class=QSlabel><%=l("email")%>:</td><td><input type=text name=sEmail size=20 value="<%=quotRep(contactSearch.sEmail)%>" />&nbsp;<i><%=l("nickname")%></i>: <input type=text name=sNickname size=20 value="<%=quotRep(contactSearch.sNickname)%>" /></td></tr><tr><td class=QSlabel><%=l("memberrole")%>:</td><td><select name=iStatus><option value=''></option><%=cslist.showSelected("option",contactSearch.iStatus)%></select></tr><%if customer.hasManyContacts then%><tr><td class=QSlabel><%=l("pagesize")%>:</td><td><select name=pageSize><option value="99999"><%=l("showall")%></option><%=numberList(100,500,100,contactSearch.origPageSize)%></select></td></tr><%end if%><tr><td class=QSlabel>&nbsp;</td><td><input class="art-button" type=submit value="<%=l("search")%>" name=btnaction />&nbsp;<input class="art-button" type=submit value="<%=l("excel")%>" name=btnaction />&nbsp;<input class="art-button" type=button value="<%=l("reset")%>" name=btnaction onclick="javascript:location.assign('bs_contacthome.asp')" /></td></tr></table></form><%select case contactSearch.mode
case l("search")%><form name=accessForm method=post action="bs_contactSelectionActions.asp"><input type=hidden name=sActionType /><%if contactSearch.recordCount>0 and secondAdmin.bIntranetMail then%><p align=center>-> <a href="#" onclick="javascript:if(confirm('<%=l("areyousure")%>\n<%=l("warningcontacts")%>')){accessForm.sActionType.value='<%=SB_massMail%>';accessForm.submit();};"><%=l("emailtoselection")%></a> <-</p><%end if%><%=resultTable%></form><br /><table align=center cellspacing=0 cellpadding=4><tr><td class="cs_silent" align=center valign=middle><%=cslist.showSelected("single",cs_silent)%></td></tr><tr><td class="cs_profile" align=center valign=middle><%=cslist.showSelected("single",cs_profile)%></td></tr><tr><td class="cs_read" align=center valign=middle><%=cslist.showSelected("single",cs_read)%></td></tr><tr><td class="cs_write" align=center valign=middle><%=cslist.showSelected("single",cs_write)%></td></tr><tr><td class="cs_readwrite" align=center valign=middle><%=cslist.showSelected("single",cs_readwrite)%></td></tr></table><%case l("excel")%><%=resultTable%><%end select%><%if convertGetal(request.querystring("iCPP"))<>0 then%><script type="text/javascript">$(window).load(function(){
var sGetUrl = "bs_contactEdit.asp?iContactID=<%=encrypt(request.querystring("iCPP"))%>";
$.colorbox({close: "Close", open:true, innerWidth:"90%", innerHeight:516, iframe:true, href:sGetUrl,onClosed:function(){ location.assign('bs_contactHome.asp');}}); 
});</script><%end if%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
