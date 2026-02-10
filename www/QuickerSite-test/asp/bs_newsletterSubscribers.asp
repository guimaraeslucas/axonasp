<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><!-- #include file="includes/header.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Newsletter)%><script  type="text/javascript"><!--
function SetChecked(val) {
dml=document.mainform;
len = dml.elements.length;
var i=0;
for( i=0 ; i<len ; i++) {
if (dml.elements[i].name=='deleteID') {
dml.elements[i].checked=val;}
}
}
//--></script><p align=center><%if customer.bCanImportSubscribers then%><a class="art-button" href="bs_NewsletterImport.asp"><b>Import Subscribers</b></a>  <%end if%><a class="art-button" href="bs_NewsletterList.asp"><b>Newsletter home</b></a></p><%dim NewsletterCategories,iCategoryID,cat,searchname,searchemail
searchname=trim(request.form("sName"))
searchemail=trim(request.form("sEmail"))
iCategoryID=convertGetal(decrypt(request("iCategoryID")))
set NewsletterCategories=customer.NewsletterCategories
if NewsletterCategories.count=1 then
for each cat in NewsletterCategories
iCategoryID=cat
exit for 
next
end if
'delete
dim delID,delNLRS
for each delID in request.form("deleteID")

set delNLRS=db.execute("delete from tblNewsletterLog where iSubscriberID=" & convertGetal(delID))
set delNLRS=nothing

set delNLRS=db.execute("delete from tblNewsletterCategorySubscriber where iCustomerID=" & cId & " and iId=" & convertGetal(delID))
set delNLRS=nothing

next
dim sActiveSQL
select case request.form("bActive")
case ""
sActiveSQL=" and bActive=" & getSQLBoolean(true) & " "
case "i"
sActiveSQL=" and bActive=" & getSQLBoolean(false) & " "
case "b"
sActiveSQL=""
end select
dim rs,counter,sql
sql="select count(*) from tblNewsletterCategorySubscriber where iCategoryID=" & iCategoryID & sActiveSQL & " and iCustomerID=" & cId
if not isLeeg(searchname) then
sql=sql& " and sName like '%" & left(cleanup(searchname),50) & "%'"
end if
if not isLeeg(searchemail) then
sql=sql& " and sEmail like '%" & left(cleanup(searchemail),50) & "%'"
end if
set rs=db.execute(sql)
counter=rs(0)
set rs=nothing
if request.form ("pcase")="1" then
set rs=db.getDynamicRS
sql="select * from tblNewsletterCategorySubscriber where iCategoryID=" & iCategoryID & sActiveSQL & " and iCustomerID=" & cId
rs.open(sql)
while not rs.eof
rs("sName")=PCase(convertStr(rs("sName")))
rs.update()
rs.movenext
wend 
rs.close()
set rs=nothing
end if
sql="select * from tblNewsletterCategorySubscriber where iCategoryID=" & iCategoryID & sActiveSQL & " and iCustomerID=" & cId 
if not isLeeg(searchname) then
sql=sql& " and sName like '%" & left(cleanup(searchname),50) & "%'"
end if
if not isLeeg(searchemail) then
sql=sql& " and sEmail like '%" & left(cleanup(searchemail),50) & "%'"
end if
sql=sql & " order by sName"
set rs=db.execute(sql)%><p align=center>Total: <strong><%=counter%></strong> subscribers</p>
<%
if not isLeeg(request.form("unsubscribers")) and iCategoryID<>0 then

dim arrUNSUBS,us,recordSET
arrUNSUBS=split(request.form("unsubscribers"),vbcrlf)

for us=lbound(arrUNSUBS) to ubound(arrUNSUBS)
set recordSET=db.getDynamicRS()
recordSET.open "select * from tblNewsletterCategorySubscriber where iCategoryID=" & iCategoryID & " and sEmail='" & replace(lcase(trim(convertStr(arrUNSUBS(us)))),"'","''",1,-1,1) & "'"
if not recordSET.eof then
recordSET("bActive")=false
recordSET.update()
end if
recordSET.close
set recordSET=nothing
next


end if
%>



<form action="bs_newsletterSubscribers.asp" method="post" name="mainform">
<input type="hidden" name="pcase" value="" /><input type="hidden" name="dummy" value="" />
<p align=center><a href="#" onclick="javascript:if(confirm('Are you sure to delete the selected subscribers?')){mainform.submit();}">Delete selected</a> (<i><a href="#" onclick="javascript:SetChecked(1);">Check all</a> / <a href="#" onclick="javascript:SetChecked(0);">Clear all</a></i>)</p><table align="center" border=0 cellpadding=3 cellspacing=0><tr><td colspan=6><a href="#" onclick="javascript:if(confirm('Do you wish to proper case all names? - all first characters will be uppercase, the other lower case...')){mainform.pcase.value='1';mainform.submit();}">(Proper case?)</a></td></tr><tr><td align=center><input type=text name="sName" size="20" maxlength="50" value="<%=quotrep(searchname)%>" /><br /><a href="#" onclick="javascript:mainform.submit();"><%=l("search")%></a></td><td align=center><input type=text name="sEmail" size="20" maxlength="50" value="<%=quotrep(searchemail)%>" /><br /><a href="#" onclick="javascript:mainform.submit();"><%=l("search")%></a></td><td><select onchange="javascript:mainform.submit()" name="bActive"><option <%if request.form("bActive")="" then response.write "selected='selected'"%> value="">Active</option><option <%if request.form("bActive")="i" then response.write "selected='selected'"%> value="i">Inactive</option><option <%if request.form("bActive")="b" then response.write "selected='selected'"%> value="b">Both</option></select></td><td><select onchange="javascript:mainform.submit()" name="iCategoryID"><option>Select list</option><%for each cat in NewsletterCategories%><option <%if iCategoryID=cat then response.write "selected='selected'"%> value="<%=encrypt(cat)%>"><%=quotrep(NewsletterCategories(cat).sname)%></option><%next%></select></td><td><b><%=l("modify")%></b></td><td><b>Delete?</b></td></tr><%dim subCounter
subCounter=0
while not rs.eof and subCounter<2500
on error resume next
response.write "<tr>" & vbcrlf
response.write "<td style=""border-top:1px solid #DDD"">" & left(quotrep(rs("sName")),40) & "&nbsp;</td>" & vbcrlf
response.write "<td style=""border-top:1px solid #DDD""><a href=""mailto:" & quotrep(rs("sEmail")) & """>" & left(quotrep(rs("sEmail")),40) & "</a></td>" & vbcrlf
response.write "<td style=""border-top:1px solid #DDD"" align=""center"">" & jaNeen(convertBool(rs("bActive"))) & "</td>" & vbcrlf
response.write "<td style=""border-top:1px solid #DDD"">" & NewsletterCategories(convertGetal(rs("iCategoryID"))).sName & "</td>" & vbcrlf
response.write "<td style=""border-top:1px solid #DDD"" align=""center""><a href=""bs_newsletterSubscriber.asp?iSubscriptionID=" & encrypt(rs("iId")) & """>" & l("modify") & "</a></td>" & vbcrlf
response.write "<td style=""border-top:1px solid #DDD"" align=""center""><input type=""checkbox"" value=""" & rs("iId") & """ name=""deleteID"" /></td>"	 & vbcrlf
response.write "</tr>" & vbcrlf
rs.movenext
subCounter=subCounter+1
on error goto 0
wend
if subCounter>2499 then%>
<script type="text/javascript">alert('Warning! Only the first 2500 records are shown.\nDefine your query to search for any specific subscriber.');</script>
<%end if%></table>
<%
if convertGetal(iCategoryID)<>0 then
%>
<table align="center">
	<tr>
		<td><textarea name="unsubscribers" cols="50" rows="5"><%=sanitize(request.form("unsubscribers"))%></textarea><br />
		<span style="font-size:8pt">enter-separated list of email addresses</span></td>
		<td><input onclick="javascript:return confirm('Are you sure to unsubscribe these email addresses from the selected list?')" type="submit" name="massUnsubs" class="art-button" value="Unsubscribe" /></td>
	</tr>
</table>
<%
end if
%>
</form>
<!-- #include file="bs_endBack.asp"-->
<!-- #include file="includes/footer.asp"-->
