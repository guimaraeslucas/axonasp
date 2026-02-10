<!-- #include file="begin.asp"-->


<!-- #include file="bs_security.asp"--><%logon.hasaccess secondAdmin.bNewsletter%><%if not customer.bCanImportSubscribers then response.write ("bs_NewsletterList.asp")%><!-- #include file="includes/header.asp"--><!-- #include file="includes/pb.asp"--><!-- #include file="bs_initBack.asp"--><!-- #include file="bs_header.asp"--><%=getBOHeader(btn_Newsletter)%><p align=center><a class="art-button" href="bs_NewsletterSubscribers.asp"><b>Subscribers</b></a>  <a class="art-button" href="bs_NewsletterList.asp"><b>Newsletter home</b></a></p><%dim NewsletterCategories,NC,pairs
set NewsletterCategories=customer.NewsletterCategories
if request.form("pairs")<>"" then
pairs=request.form("pairs")
else
pairs=""
end if
dim sCounter
sCounter=0
if request.form("btnaction")="Import" then
'some checking
if isLeeg(request.form("cat")) then
message.AddError("err_mandatory")
end if
if not message.HasErrors then
dim arrPairs,arrPair,activeorinactive
if convertBool(request.form("bimportinactive")) then
activeorinactive=false
else
activeorinactive=true
end if
arrPairs=split(pairs,vbcrlf)
dim ai,arr,p,sName,sEmail,rs,iCat
for each iCat in request.form("cat")
for ai=lbound(arrPairs) to ubound(arrPairs)
arrPair=split(arrPairs(ai),",")
on error resume next
select case request.form("importType")
case "","0"
sName=trim(arrPair(0))
sEmail=trim(lcase(arrPair(1)))
case "1"
sName=""
sEmail=trim(lcase(arrPair(0)))
end select 
if err.number=0 then
if checkEmailSyntax(sEmail) then
set rs=db.execute("select * from tblNewsletterCategorySubscriber where iCategoryID="  & iCat & " and sEmail='" & sEmail & "'")
if rs.eof then
set rs=nothing
set rs=db.getDynamicRS
rs.open ("select * from tblNewsletterCategorySubscriber where 1=2")
rs.Addnew
rs("sName")=sName
rs("sEmail")=sEmail
rs("sKey")=lcase(generatePassword & generatePassword & generatePassword)
rs("iCategoryID")=convertGetal(iCat)
rs("iCustomerID")=cId
rs("bActive")=activeorinactive
rs("dAdded")=date()
rs.update()
rs.close
set rs=nothing
sCounter=sCounter+1
end if
end if
end if
err.clear()
on error goto 0
next
next
end if
end if%><%if sCounter>0 then%><table align=center cellpadding=5><tr><td align=center style="text-align:center;background-color:Yellow;color:Green"><b><%=sCounter%> subscriptions were added</b></td></tr></table><%end if
if request.form("btnaction")="Import" and sCounter=0 then%><table align=center cellpadding=5><tr><td align=center style="text-align:center;background-color:Yellow;color:Red"><b>No subscriptions were added - no new name-email pairs were found</b></td></tr></table><%end if%><form name=mainform method=post action="bs_newsletterImport.asp"><table align="center"><tr><td style="width:50%"><p><input type=radio name=importType <%if request.form("importType")="" or request.form("importType")="1" then response.write " checked='checked' "%>value=1 /><b>ENTER SEPARATED</b> <u>email addresses only</u></p><p><input type=radio name=importType <%if request.form("importType")="0" then response.write " checked='checked' "%>value=0 /><b>ENTER SEPARATED</b> Name,Email (name comma email)</p><%if Request.Form ("btnaction")="Extract valid email addresses" then
dim pb
set pb=new privatebot
pb.quickSearch(Request.Form ("pairs"))
set pb=nothing
else%><textarea rows="10" cols="75" name="pairs"><%=quotrep(pairs)%></textarea><%end if%></td><td style="width:20px">&nbsp;</td><td style="width:50%" valign="top"><p>Select the email lists:*</p><%dim cat
for each NC in NewsletterCategories
response.write "<input type='checkbox' value='" & NC & "' name='cat' "
if NewsletterCategories.count=1 then
response.write " checked='checked' "
end if
for each cat in request.form("cat")
if convertStr(cat)=convertStr(NC) then
response.write " checked='checked' "
end if
next
response.write "/>" & NewsletterCategories(NC).sName & "<br />"
next%><p><hr /></p><p><input type=checkbox value=1 <%if convertBool(request.form("bimportinactive")) then response.write "checked='checked'"%> name=bimportinactive /> Import as inactive</p></td></tr><tr><td colspan=2 align="center">(*) <%=l("mandatory")%></td></tr><tr><td colspan=2 align="center"><input class="art-button" type=submit value="Extract valid email addresses" name="btnaction" /> <input class="art-button" type=submit value="Import" name="btnaction" /><br /><br /><small>"Extract valid email addresses": You can paste any text/html in the box and after you click the button, a list of email addresses will be extracted.</small></td></tr></table></form><%set NewsletterCategories=nothing%><!-- #include file="bs_endBack.asp"--><!-- #include file="includes/footer.asp"-->
