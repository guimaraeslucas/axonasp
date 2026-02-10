
<%class cls_poll
Public iId
Public sQuestion
Public sCode
Public iCustomerID
Public dCreatedTS
Public Questions
Public bShowTitle
Public bCanVote
Public label_viewresults
Public label_votenow
Public label_numberofvotes
Public dVoteFrom
Public dVoteUntil
Public dVoteDeadline
Public dResetDate
Private Sub Class_Initialize
On Error Resume Next
set Questions=server.CreateObject ("scripting.dictionary")
bCanVote=true
label_viewresults="View Results"
label_votenow="Vote Now!"
label_numberofvotes="Number of votes: [NMBR]"
dim arr(2)
arr(0)="#67645F"
arr(1)=""
Questions.Add 1,arr
arr(0)="#9A948F"
arr(1)=""
Questions.Add 2,arr
arr(0)="#CDC5BF"
arr(1)=""
Questions.Add 3,arr
arr(0)="#68AB7D"
arr(1)=""
Questions.Add 4,arr
arr(0)="#89C49C"
arr(1)=""
Questions.Add 5,arr
arr(0)="#B0DFBF"
arr(1)=""
Questions.Add 6,arr
arr(0)="#2577CF"
arr(1)=""
Questions.Add 7,arr
arr(0)="#6FAAE8"
arr(1)=""
Questions.Add 8,arr
arr(0)="#A6D1FF"
arr(1)=""
Questions.Add 9,arr
arr(0)="#6B0C3A"
arr(1)=""
Questions.Add 10,arr
arr(0)="#DE267F"
arr(1)=""
Questions.Add 11,arr
arr(0)="#F57ABA"
arr(1)=""
Questions.Add 12,arr
bShowTitle=true
pick(decrypt(request("iPollID")))
On Error Goto 0
end sub
Public Function getRequestValues()
sQuestion	= convertStr(Request.Form ("sQuestion"))
sCode	= convertStr(Request.Form ("sCode"))
bShowTitle	= convertBool(Request.Form ("bShowTitle"))
label_viewresults	= convertStr(Request.Form ("label_viewresults"))
label_votenow	= convertStr(Request.Form ("label_votenow"))
label_numberofvotes	= convertStr(Request.Form ("label_numberofvotes"))
dVoteFrom	= convertDateFromPicker(Request.Form ("dVoteFrom"))
dVoteUntil	= convertDateFromPicker(Request.Form ("dVoteUntil"))
dVoteDeadline	= convertDateFromPicker(Request.Form ("dVoteDeadline"))
Questions.RemoveAll 
dim i,aArr(2)
for i=1 to 15
aArr(0)=Request.Form ("sA"&i&"c")
aArr(1)=Request.Form ("sA"&i)
Questions.Add i,aArr
next
end Function
Public function getByCode(code)
on error resume next
getByCode=false
dim rs
set rs=db.execute("select iId from tblPoll where iCustomerID=" & cid & " and sCode='"& left(cleanup(code),50) & "'")
if not rs.eof then
pick(rs(0))
getByCode=true
end if
on error goto 0
end function
Public Function Pick(id)
dim sql, rs
dim aArr(2)
if isNumeriek(id) then
sql = "select * from tblPoll where iCustomerID="&cid&" and iId=" & id
set rs = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sQuestion	= rs("sQuestion")
sCode	= rs("sCode")
iCustomerID	= rs("iCustomerID")
dCreatedTS	= rs("dCreatedTS")
bShowTitle	= rs("bShowTitle")
label_viewresults	= rs("label_viewresults")
label_votenow	= rs("label_votenow")
label_numberofvotes	= rs("label_numberofvotes")
dVoteUntil	= rs("dVoteUntil")
dVoteFrom	= rs("dVoteFrom")
dVoteDeadline	= rs("dVoteDeadline")
dResetDate	= rs("dResetDate")
Questions.RemoveAll 
dim i
for i=1 to 15
aArr(0)	= convertStr(rs("sA" & i & "c"))
aArr(1)	= convertStr(rs("sA" & i))
Questions.Add i,aArr
next
end if
set RS = nothing
end if
end function
Public function getQuestion(key)
On Error Resume Next
getQuestion=Questions(key)(1)
On Error Goto 0
end function
Public function getQuestionColor(key)
On Error Resume Next
getQuestionColor=Questions(key)(0)
On Error Goto 0
end function
Public Function Check()
Check = true
if isLeeg(sQuestion) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sCode) then
check=false
message.AddError("err_mandatory")
end if
End Function
Public Function Save()
if check() then
Save=true
else
Save=false
exit function
end if
'set db=nothing
'set db=new cls_database
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblPoll where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblPoll where iId="& iId
end if
rs("sQuestion")	= sQuestion
rs("sCode")	= sCode
rs("dCreatedTS")	= dCreatedTS
rs("iCustomerID")	= cId
rs("bShowTitle")	= bShowTitle
rs("label_viewresults")	= label_viewresults
rs("label_votenow")	= label_votenow
rs("label_numberofvotes")	= label_numberofvotes
rs("dVoteFrom")	= dVoteFrom
rs("dVoteUntil")	= dVoteUntil
rs("dVoteDeadline")	= dVoteDeadline
rs("dResetDate")	= dResetDate
dim i
for i=1 to 15
rs("sA" & i)	= getQuestion(i)
rs("sA" & i & "c")	= getQuestionColor(i)
next
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblPollVote where iPollID=" & iId)
set rs=nothing
set rs=db.execute("delete from tblPoll where iId=" & iId)
set rs=nothing
end if
end function
public function reset
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblPollVote where iPollID=" & iId)
set rs=nothing
dResetDate=now()
save()
end if
end function
public function build(withdiv)
if convertGetal(iId)<>0 then
On Error Resume Next
if not bIsOnline then exit function
if canVote then
if withdiv then
build="<div id='div" & iId & "'>"
end if
build=build&"<table border=""0"" style='width:100%' class='QS_pollquestion'>"
if bShowTitle then
build=build&"<tr><td colspan=""2"">" & sQuestion & "</td></tr>"
end if
dim i
for i=1 to 15
if not isLeeg(getQuestion(i)) then
build=build & "<tr>"
build=build & "<td style=""width:10px""><input type=""radio"" name=""vote"" value=""" & i & """ "
build=build & "onclick=""javascript:"
build=build & "qs_div='" & iId & "';"
build=build & "mode='';getVote(this.value,'" & Encrypt(iId) & "');"" /></td>"
build=build & "<td style=""width:99%""> " & getQuestion(i) & "</td>"
build=build & "</tr>"
end if
next 
build=build&"<tr><td colspan=""2""><a href='#' onclick=""javascript:mode='view';"
build=build&"qs_div='" & iId & "';"
build=build&"getVote(this.value,'" & Encrypt(iId) & "');return false"">" & label_viewresults & "</a></td></tr>"
build=build&"</table>"
if withdiv then
build=build&"</div>"
end if
else
build=showresults
end if
On Error GOto 0
end if
end function 
private property get canVote
on error resume next
canVote=true
'exit property
if not bCanVote then
canVote=false
exit property
end if
if not bIsOnline then
canVote=false
exit property
end if
if not isLeeg(dVoteDeadline) then
if date>dVoteDeadline then
canVote=false
exit property
end if
end if
if convertGetal(iId)=0 then
canVote=false
exit property
end if
if UserIP="" then
canVote=false
exit property
end if
if convertBool(session("hasvoted"&iId)) then
canVote=false
exit property
end if
dim rs
set rs=db.execute("select count(*) from tblPollVote where iPollID=" & iId & " and IP='" & left(cleanup(UserIP),40) & "'")
if convertGetal(rs(0))>5 then
canVote=false
exit property
end if
if not isLeeg(dResetDate) then
if dResetDate>cdate(request.Cookies("hasvotedon"&iId)) then
canVote=true
exit property
end if
end if
if request.Cookies("hasvotedon"&iId)<>"" then
canVote=false
exit property
end if
if convertGetal(rs(0))=0 then
canVote=true
exit property
end if
set rs=nothing
if err.number <>0 then
canVote=false
exit property
end if
on Error Goto 0
end property
public sub registerVote
on error resume next
if canvote then
'set db=nothing
'set db=new cls_database
dim rs
set rs = db.GetDynamicRS
rs.Open "select * from tblPollVote where 1=2"
rs.AddNew
rs("dVoteTS")=date()
rs("iPollID")=iId
rs("iVote")=convertGetal(Request.QueryString ("vote"))
rs("IP")=left(UserIP,40)
rs.update()
set rs=nothing
session("hasvoted"&iId)=true
Response.Cookies("hasvotedon"&iId)=now()
Response.Cookies("hasvotedon"&iId).expires=dateAdd("d",10,date())
Response.write showresults()
end if
on error goto 0
end sub
private function bIsOnline
bIsOnline=true
if not isLeeg(dVoteUntil) then
if date>dVoteUntil then
bIsOnline=false
exit function
end if
end if
if not isLeeg(dVoteFrom) then
if date<dVoteFrom then
bIsOnline=false
exit function
end if
end if
end function
public function showresults
on error resume next
if convertGetal(iId)<>0 then
if not bIsOnline then exit function
dim total,percent,totalquot
total=0
'This will summerize and count the records in the db
dim rs,i,arr(15),cc
for i=1 to 15
set rs= db.execute("select count(*) from tblPollVote where iPollID="& iId & " and iVote=" & i)
cc=rs(0)
arr(i)=clng(cc)
total=total+(arr(i))
next
if total=0 then 
totalquot=1
else
totalquot=total
end if
showresults="<table border=""0"" style=""width:100%"" class=""QS_pollquestion"">"
if bShowTitle then
showresults=showresults & "<tr><td>" & sQuestion & "</td></tr>"
end if
for i=1 to 15
if not isLeeg(getQuestion(i)) then
percent=formatnumber ((arr(i)/totalquot)*100,0)
showresults=showresults & "<tr>"
showresults=showresults & "<td>" & getQuestion(i)  & "&nbsp;("& percent & "%)</td>"
showresults=showresults & "</tr>"
if convertGetal(percent)<>0 then
showresults=showresults & "<tr>"
showresults=showresults & "<td style=""width:100%"">"
showresults=showresults & "<div style=""height:10px;width:" & percent & "%;background-color:" & getQuestionColor(i)  & """></div></td>"
showresults=showresults & "</tr>"
end if
end if
next 
showresults=showresults & "<tr><td>" & replace (label_numberofvotes,"[NMBR]",total,1,-1,1) & "</td></tr>" 
if canvote then
showresults=showresults & "<tr><td><a href=""#"" onclick=""javascript:qs_div='" & iId & "';mode='vote';getVote(this.value,'" & Encrypt(iId) & "');return false;"">" & label_votenow & "</a></td></tr>" 
end if
showresults=showresults&"</table>"
end if
on error resume next
end function
public function copy()
if isNumeriek(iId) then
iId=null
sCode=GeneratePassword
save()
end if
end function
end class%>
