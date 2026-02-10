
<%function treatConstants(byref origValue,fill)
On Error Resume Next
if isLeeg(origValue) then 
treatConstants=""
exit function
end if
if not bLoadConstants then
treatConstants=origValue
exit function
end if
'treatConstants=origValue
'exit function
origValue=convertStr(origValue)
origValue=replace(origValue,"[CDATA[","{{CDATA{{",1,-1,1)
origValue=replace(origValue,"<![endif]-->","<!{endif]-->",1,-1,1)
origValue=replace(origValue,"<!--[if IE ","<!--{if IE ",1,-1,1)
origValue=replace(origValue,"[]","####GA-FIX####",1,-1,1)
if instr(1,origValue,"[",VBTextCompare)=0 then 
origValue=replace(origValue,"{{CDATA{{","[CDATA[",1,-1,1)
origValue=replace(origValue,"<!{endif]-->","<![endif]-->",1,-1,1)
origValue=replace(origValue,"<!--{if IE ","<!--[if IE ",1,-1,1)
origValue=replace(origValue,"|R|R|R|","[",1,-1,1)
origValue=replace(origValue,"####GA-FIX####","[]",1,-1,1)
treatConstants=origValue
exit function
end if
'origValue=addSmilies(origValue)
dim cNester
if convertGetal(customer.defaultTemplate)<>0 then
cNester=2
else
cNester=4
end if
if instr(1,origValue,"[",VBTextCompare)<>0 then 
'insert constants
if not isEmpty(application(QS_CMS_arrconstants)) then
dim iconstantKey, continue, nester, re, matches, mv, cname
continue=true
nester=0
set re=new regexp
re.Global=true
re.IgnoreCase=true
do while continue and nester<cNester
'Response.Write "nester: " & nester & "<br />"
continue=false
for iconstantKey=lbound(application(QS_CMS_arrconstants),2) to ubound(application(QS_CMS_arrconstants),2)
cname=application(QS_CMS_arrconstants)(0,iconstantKey)
re.Pattern ="\[+("&cname&")+[\(]+[\S| ]+[\)]+[\]]|\[+("&cname&")+[\]]"
if re.Test (origValue) then
'Response.Write "sGlobal: " & application(QS_CMS_arrconstants)(2,iconstantKey)
if instr(1,application(QS_CMS_arrconstants)(1,iconstantKey),QS_VBScriptIdentifier,VBTextCompare)<>0 then
'treat as ASP/VBSscript
set matches=re.Execute (replace(origValue,"]","]" & vbcrlf,1,-1,1))
re.Pattern ="\[|\]|\(|\)|("&cname&")"
for each mv in matches
if fill then

dim startposA,endposA,paramA
startposA=instr(mv.value,"(")
endposA=instrRev(mv.value,")")
'response.write endposA
'response.flush

if startposA<>0 then
	paramA=mid(mv.value,startposA+1,endposA-startposA-1)
else
	paramA=re.Replace (mv.value,"")
end if

'response.write paramA
'response.end


origValue=Replace(origValue, mv.value,executeConstant(application(QS_CMS_arrconstants)(1,iconstantKey),false,paramA,application(QS_CMS_arrconstants)(2,iconstantKey)),1,-1,1)
else
origValue=Replace(origValue, mv.value,"",1,-1,1)
end if
next
set matches=nothing
else
'treat as text/html
if fill then
origValue=Replace(origValue, "[" & cname & "]",application(QS_CMS_arrconstants)(1,iconstantKey),1,-1,1)
else
origValue=Replace(origValue, "[" & cname & "]","",1,-1,1)
end if
end if
end if
re.Pattern="(\[)"
if re.Test(origValue) then 
continue=true
else
exit for
end if
next
nester=nester+1
loop
set re=nothing
else
if isEmpty(application(QS_CMS_constantsreloaded)) then 
customer.cacheConstants
treatConstants=treatConstants(origValue,fill)
end if
end if
end if
if instr(1,origValue,"[qs_feed:",VBTextCompare)<>0 then
'insert feeds
if not isEmpty(application(QS_CMS_arrfeeds)) then
dim ifeedKey, feedObj
for ifeedKey=lbound(application(QS_CMS_arrfeeds),2) to ubound(application(QS_CMS_arrfeeds),2)
if instr(1,lcase(origValue),"[qs_feed:" & lcase(application(QS_CMS_arrfeeds)(1,ifeedKey)) &"]",VBTextCompare)<>0 then 
if fill then
set feedObj=new cls_feed
feedObj.pick(application(QS_CMS_arrfeeds)(0,ifeedKey))
origValue=Replace(origValue, "[qs_feed:" & application(QS_CMS_arrfeeds)(1,ifeedKey) &"]",feedObj.build(),1,-1,1)
set feedObj=nothing
else
origValue=Replace(origValue, "[qs_feed:" & application(QS_CMS_arrfeeds)(1,ifeedKey) &"]","",1,-1,1)
end if
end if
next
else
if isEmpty(application(QS_CMS_feedsreloaded)) then 
customer.cachefeeds()
treatConstants=treatConstants(origValue,fill)
end if
end if
end if
'###########################################
'fill POLL values
dim startPos,endpos,pollcode,pollcounter
pollcounter=0
startpos=instr(origValue,"[QS_POLL:")
while startPos<>0 and pollcounter<25
endpos=instr(startpos,origValue,"]")
pollcode=mid(origValue,startpos+9,endpos-startpos-9)
dim poll
set poll=new cls_poll
if poll.getByCode(pollcode) and fill then
origValue=Replace(origValue, "[QS_POLL:" & pollcode & "]",poll.build(true),1,1,1)
else
origValue=Replace(origValue, "[QS_POLL:" & pollcode & "]","",1,-1,1)
end if
set poll=nothing
startpos=instr(origValue,"[QS_POLL:")
pollcounter=pollcounter+1
wend
'###########################################
'###########################################
'fill GB values
dim gbcode
startpos=instr(origValue,"[QS_GUESTBOOK:")
while startPos<>0
endpos=instr(startpos,origValue,"]")
gbcode=mid(origValue,startpos+14,endpos-startpos-14)
dim gbk
set gbk=new cls_guestbook
if gbk.getByCode(gbcode) and fill then
origValue=Replace(origValue, "[QS_GUESTBOOK:" & gbcode & "]",gbk.build(),1,-1,1)
else
origValue=Replace(origValue, "[QS_GUESTBOOK:" & gbcode & "]","",1,-1,1)
end if
set gbk=nothing
startpos=instr(origValue,"[QS_GUESTBOOK:")
wend
'###########################################
if instr(1,origValue,"[qs_gallery:",VBTextCompare)<>0 then
'insert galleries
if not isEmpty(application(QS_CMS_arrgalleries)) then
dim iGalleryKey, galleryObj
for iGalleryKey=lbound(application(QS_CMS_arrgalleries),2) to ubound(application(QS_CMS_arrgalleries),2)
if instr(1,lcase(origValue),"[qs_gallery:" & lcase(application(QS_CMS_arrgalleries)(1,iGalleryKey)) &"]",VBTextCompare)<>0 then 
if fill then
set galleryObj=new cls_gallery
galleryObj.pick(application(QS_CMS_arrgalleries)(0,iGalleryKey))
while instr(1,origValue,"[qs_gallery:" & application(QS_CMS_arrgalleries)(1,iGalleryKey) &"]",VBTextCompare)<>0 
origValue=Replace(origValue, "[qs_gallery:" & application(QS_CMS_arrgalleries)(1,iGalleryKey) &"]",galleryObj.build(),1,1,1)
wend 
set galleryObj=nothing
else
origValue=Replace(origValue, "[qs_gallery:" & application(QS_CMS_arrgalleries)(1,iGalleryKey) &"]","",1,-1,1)
end if
end if
next
else
if isEmpty(application(QS_CMS_galleriesreloaded)) then 
customer.cachegalleries()
treatConstants=treatConstants(origValue,fill)
end if
end if
end if
'###########################################
'fill NLCATR values
startpos=instr(origValue,"[QS_NLCAT_")
dim nlCATCode,nlc,catRunner
catRunner=0
while startPos<>0 and catRunner<5
catRunner=catRunner+1
endpos=instr(startpos,origValue,"]")
nlCATCode=mid(origValue,startpos+10,endpos-startpos-10)
set nlc=new cls_newsletterCategory
'response.write "'" & nlCATCode & "'"
'response.end
nlc.pick(nlCATCode)
if convertGetal(nlc.iId)<>0 and fill then
origValue=Replace(origValue, "[QS_NLCAT_" & nlCATCode & "]",nlc.build(),1,-1,1)
else
origValue=Replace(origValue, "[QS_NLCAT_" & nlCATCode & "]","",1,-1,1)
end if
set nlc=nothing
startpos=instr(origValue,"[QS_NLCAT_")
wend
'###########################################
if instr(1,origValue,"[qs_theme:",VBTextCompare)<>0 then
'insert themes
if not isEmpty(application(QS_CMS_arrthemes)) then
dim ithemeKey, themeObj
for ithemeKey=lbound(application(QS_CMS_arrthemes),2) to ubound(application(QS_CMS_arrthemes),2)
if instr(1,lcase(origValue),"[qs_theme:" & lcase(application(QS_CMS_arrthemes)(1,ithemeKey)) &"]",VBTextCompare)<>0 then 
if fill then
set themeObj=new cls_theme
themeObj.pick(application(QS_CMS_arrthemes)(0,ithemeKey))
origValue=Replace(origValue, "[qs_theme:" & application(QS_CMS_arrthemes)(1,ithemeKey) &"]",themeObj.build(),1,-1,1)
set themeObj=nothing
else
origValue=Replace(origValue, "[qs_theme:" & application(QS_CMS_arrthemes)(1,ithemeKey) &"]","",1,-1,1)
end if
end if
next
else
if isEmpty(application(QS_CMS_themesreloaded)) then 
customer.cachethemes()
treatConstants=treatConstants(origValue,fill)
end if
end if
end if
On Error Goto 0
origValue=replace(origValue,"{{CDATA{{","[CDATA[",1,-1,1)
origValue=replace(origValue,"<!{endif]-->","<![endif]-->",1,-1,1)
origValue=replace(origValue,"<!--{if IE ","<!--[if IE ",1,-1,1)
origValue=replace(origValue,"|R|R|R|","[",1,-1,1)
origValue=replace(origValue,"####GA-FIX####","[]",1,-1,1)
treatConstants=origValue
end function
function executeConstant(sScript,testMode,sParameters,sGlobal)
on error resume next
if customer.bApplication then
if not isLeeg(sGlobal) then

sGlobal=removeCRB(sGlobal)

executeGlobal(treatConstants(sGlobal,true))
end if
dim arrSc
arrSc=split(sScript,QS_VBScriptIdentifier)
sScript=treatConstants(arrSc(0),true)

sScript=removeCRB(sScript)

dim pars
pars=treatConstants(arrSc(1),true)
if isLeeg(pars) then pars="dummyPar"
if isLeeg(sParameters) then sParameters="dummyPar"
executeConstant="function CustomFunction("&pars&") " & vbcrlf & replace(sScript,"response.write","CustomFunction=CustomFunction&",1,-1,1) & vbcrlf & "end function " & vbcrlf
dim fullcode
fullcode=executeConstant & vbcrlf & "executeConstant=CustomFunction("&sParameters&")"
execute(fullcode)
if Err.number<>0 then
if testMode then
executeConstant="<b>TEST FAILED!<br />Error number: "& Err.number & "<br />" & Err.Source & "<br />" & Err.Description  &"</b><br />"&executeConstant
executeConstant=linkUrls(executeConstant)
else
executeConstant=""
dumperror fullcode,err 
end if
Err.Clear()
else
if testMode then
executeConstant="<b>TEST OK!</b><br />Output below:<br /><br />"&executeConstant
end if
end if
end if
on error goto 0
end function%>
