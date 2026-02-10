
<%class cls_search
public value
public orderBY
public includeURL
public allowEmptyString
public showAll
public recordCount
public includePasswordProtected
public includeListItems
public bIntranet
public includeIntranet
public bIncludeHideFromSearch
Private sub class_initialize
orderBY="sTitle"
includeURL=true
allowEmptyString=false
includePasswordProtected=false
includeListItems=false
includeIntranet=false
value=quotRep(l("search")) & "..."
bIncludeHideFromSearch=true
end sub
Public property get returnValue
if not isleeg(value) then 
returnValue=value
else 
returnValue=Request.Cookies ("searchvalue")
end if
end property
Public function results
response.cookies("searchvalue") = value
set results=server.CreateObject ("scripting.dictionary")
if not allowEmptyString and (isLeeg(cleanup(value)) or len(cleanup(value))<3) then exit function
dim splitValues, lV, continue
continue=false 
dim countSpaces, copyvalue
countSpaces=0
splitValues=split(value," ")
for lV=lbound(splitValues) to ubound(splitValues)
if len(trim(splitValues(lV)))>2 then
continue=true
countSpaces=countSpaces+1
end if
next
if not allowEmptyString then
if not continue then exit function
end if
if countSpaces>2 then 
copyvalue=""""&value&""""
else
copyvalue=value
end if
dim sql
sql="select iId from tblPage where "& sqlCustId
sql=sql&" and (bContainerPage="&getSQLBoolean(false)&" or bHomepage="&getSQLBoolean(true)&") "
sql=sql&" and (bDeleted="&getSQLBoolean(false)&" or bDeleted is null) "
sql=sql&" and (bHideFromSearch="&getSQLBoolean(false)&" or bHideFromSearch is null) "
splitValues=split(copyvalue," ")
if not isLeeg(cleanup(copyvalue)) then
sql=sql & " and ("
if left(copyvalue,1)="""" and right(copyvalue,1)="""" then
sql=sql&" sTitle like '%" & replace(left(cleanup(copyvalue),100),"""","",1,-1,1) & "%' or sValueTextOnly like '%" & replace(left(cleanup(copyvalue),100),"""","",1,-1,1) & "%') "
else
for lV=lbound(splitValues) to ubound(splitValues)
if len(trim(splitValues(lV)))>2 then
sql=sql&" sTitle like '%" & left(cleanup(splitValues(lV)),100) & "%' or sValueTextOnly like '%" & left(cleanup(splitValues(lV)),100) & "%' OR"
end if
next
sql=left(sql,len(sql)-2)& ")"
end if
end if
sql=sql&" and (bOnline="&getSQLBoolean(true)&" "
if includeListItems then
sql=sql&" OR ((iListPageID is not null and (dOnlineFrom is null or dOnlineFrom<="&getSQLDateFunction&") and (dOnlineUntill is null or dOnlineUntill>="&getSQLDateFunction&")) "
sql=sql&" AND iListPageID in (select iId from tblPage where "& sqlCustId &" and bOnline="&getSQLBoolean(true)&"))"
end if
sql=sql&")"
if not includePasswordProtected then
sql=sql&" and (sPw is null or sPw='' or sPw in ("& logon.getAllPasswordsCS &")) "
end if
if not includeURL then
 sql=sql&" and (sExternalURL='' or sExternalURL is null) "
end if
if (not logon.authenticatedIntranet or logon.contact.istatus=cs_write or logon.contact.istatus=cs_silent or logon.contact.istatus=cs_profile)  and not includeIntranet then
sql=sql&" and (bIntranet="&getSQLBoolean(false)&" or bIntranet is null) "
end if
sql=sql&" order by "& orderBy
'Response.Write sql
'Response.Flush 
'Response.End 
dim rs
set rs=db.execute(sql)
dim page,pCounter
pCounter=1
while not rs.eof and pCounter<250
set page=new cls_page
page.pick(rs(0))
results.Add page.iId, page
set page=nothing
rs.movenext
pCounter=pCounter+1
wend
set rs=nothing
'exit function
'do the themes now
if not isLeeg(cleanup(copyvalue)) then
dim whichST
if logon.authenticatedIntranet then
whichST="1,2"
else
whichST="1"
end if
sql="SELECT tblPost.iId,tblPost.sSubject,tblPost.sBody  FROM tblTheme INNER JOIN tblPost ON tblTheme.iId = tblPost.iThemeID "
sql=sql&" WHERE tblTheme.iSearchType in (" & whichST & ") "
sql=sql & " and (tblPost.bNeedsToBeValidated is null or bNeedsToBeValidated=" & getSQLBoolean(false) & ") and tblTheme.iCustomerID="&cId&" and ("
splitValues=split(copyvalue," ")
if left(copyvalue,1)="""" and right(copyvalue,1)="""" then
sql=sql&" tblPost.sSubject like '%" & left(cleanup(copyvalue),100) & "%' or tblPost.sBody like '%" & left(cleanup(copyvalue),100) & "%') "
else
for lV=lbound(splitValues) to ubound(splitValues)
if len(trim(splitValues(lV)))>2 then
sql=sql&" tblPost.sSubject like '%" & left(cleanup(splitValues(lV)),100) & "%' or tblPost.sBody like '%" & left(cleanup(splitValues(lV)),100) & "%' OR"
end if
next
sql=left(sql,len(sql)-2)& ")"
end if
'response.write sql
set rs=db.execute(sql)
dim post, encPI
pCounter=0
while not rs.eof and pCounter<250
set page=new cls_page
set post=new cls_post
post.pick(rs(0))
if isLeeg(post.sSubject) or post.sSubject="COMMENT" then
page.sTitle	= post.parentTopic.sSubject
page.iPostID	= post.parentTopic.iId
else
page.sTitle	= post.sSubject
page.iPostID	= post.iId
end if
page.sValueTextOnly	= removeHTML(rs(2))
encPI=encrypt(page.iPostID)
if not results.Exists (encPI) then
results.Add encPI,page
end if
set page=nothing
set post=nothing
rs.movenext
pCounter=pCounter+1
wend
end if
end function
Public function allPages
dim sql
sql="select "
if not showAll then
sql=sql& " top 250 "
end if
sql=sql&" tblPage.iId,tblPage.sTitle,tblCustomer.sName,tblPage.createdTS, tblPage.updatedTS, tblCustomer.sUrl, tblPage.iVisitors, tblCustomer.dResetStats from tblPage inner join tblCustomer on tblCustomer.iID=tblPage.iCustomerID where tblPage.bOnline="&getSQLBoolean(true)&" order by tblPage.updatedTS desc"
dim rs
set rs=db.getDynamicRS
rs.open sql
if not rs.eof then
allPages	= rs.getRows()
recordCount	= rs.recordCount
end if
set rs=nothing
end function
end class%>
