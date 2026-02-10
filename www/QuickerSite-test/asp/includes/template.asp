
<%class cls_template
Public iId
Public sName
Public sValue
Public overruleCID
Public sCompressValue
Public bCompress
Public CustomFooter
Public customContact
Public bSetFooterVar
Public bSetRSSLink
Public bSetContactVar
Public bSetCustomHL
Public customHL
private p_pv
private p_ev
private p_mv
private p_wv
public property get sWAPValue
if isLeeg(p_wv) then initWAP
sWAPValue=convertStr(p_wv)
end property
public property let sWAPValue(wv)
p_wv=wv
end property
public property get sPrintValue
if isLeeg(p_pv) then initPrint
sPrintValue=convertStr(p_pv)
end property
public property let sPrintValue(sv)
p_pv=sv
end property
public property get sMobileValue
if isLeeg(p_mv) and isLeeg(iId) then initMobile
sMobileValue=convertStr(p_mv)
end property
public property let sMobileValue(sv)
p_mv=sv
end property
public property get sEmailValue
if isLeeg(p_ev) then initEmail
sEmailValue=convertStr(p_ev)
end property
public property let sEmailValue(sv)
p_ev=sv
end property
Private Sub Class_Initialize
On Error Resume Next
iId	= null
bSetFooterVar	= false
bSetRSSLink	= false
overruleCID	= null
bCompress	= true
bSetContactVar	= false
bSetCustomHL	= false
pick(decrypt(request("iTemplateID")))
On Error Goto 0
end sub
Public sub init()
dim fsoTemplate
set fsoTemplate=server.CreateObject ("scripting.filesystemobject")
sValue=fsoTemplate.OpenTextFile(server.MapPath ("template.txt"),1).ReadAll 
set fsoTemplate=nothing
dim cssLinks
cssLinks="<link rel=""STYLESHEET"" type=""text/css"" href=""" & C_DIRECTORY_QUICKERSITE & "/css/qs_ltr.css"" title=""QSStyle"" />" & vbcrlf
cssLinks=cssLinks&"<link rel=""STYLESHEET"" type=""text/css"" href=""" & C_DIRECTORY_QUICKERSITE & "/css/verticalMenu_ltr.css"" title=""QSStyle"" />"
sValue=replace(sValue,"{CSSLINKS}",cssLinks,1,-1,1)
sValue=replace(sValue,"[MYQS_urlTemplates]",MYQS_urlTemplates,1,-1,1)
end sub
public sub initPrint()
on error resume next
dim fsoTemplate
set fsoTemplate=server.CreateObject ("scripting.filesystemobject")
p_pv=fsoTemplate.OpenTextFile(server.MapPath ("template_print.txt"),1).ReadAll 
set fsoTemplate=nothing
on error goto 0
end sub
public sub initEmail()
on error resume next
dim fsoTemplate
set fsoTemplate=server.CreateObject ("scripting.filesystemobject")
p_ev=fsoTemplate.OpenTextFile(server.MapPath ("template_email.txt"),1).ReadAll 
set fsoTemplate=nothing
on error goto 0
end sub
public sub initMobile()
on error resume next
dim fsoTemplate
set fsoTemplate=server.CreateObject ("scripting.filesystemobject")
p_mv=fsoTemplate.OpenTextFile(server.MapPath ("template_mobile.txt"),1).ReadAll 
set fsoTemplate=nothing
on error goto 0
end sub
public sub initWAP()
on error resume next
dim fsoTemplate
set fsoTemplate=server.CreateObject ("scripting.filesystemobject")
p_wv=fsoTemplate.OpenTextFile(server.MapPath ("template_wap.txt"),1).ReadAll 
set fsoTemplate=nothing
on error goto 0
end sub
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
if not isNull(overruleCID) then
sql = "select * from tblTemplate where iId=" & id
else
sql = "select * from tblTemplate where iCustomerID="&cid&" and iId=" & id
end if
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
sValue	= rs("sValue")
p_pv	= rs("sPrintValue")
p_ev	= rs("sEmailValue")
p_mv	= rs("sMobileValue")
p_wv	= rs("sWAPValue")
sCompressValue	= rs("sCompressValue")
bCompress	= rs("bCompress")
end if
set RS = nothing
end if
end function
Public Function Check()
Check = true
if isLeeg(sName) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sValue) then
check=false
message.AddError("err_mandatory")
end if
End Function
Public Function Save
if check() then
save=true
else
save=false
exit function
end if
'set db=nothing
'set db=new cls_database
dim rs
set rs = db.GetDynamicRS
if isLeeg(iId) then
rs.Open "select * from tblTemplate where 1=2"
rs.AddNew
else
rs.Open "select * from tblTemplate where iId="& iId
end if
if bCompress then
sCompressValue=compressHTML(sValue)
else
sCompressValue=""
end if
rs("sName")	= left(sName,50)
rs("sValue")	= sValue
rs("sEmailValue")	= p_ev
rs("sPrintValue")	= p_pv
rs("sMobileValue")	= p_mv
rs("sWAPValue")	= p_wv
rs("sCompressValue")	= sCompressValue
rs("bCompress")	= bCompress
if not isNull(overruleCID) then
rs("iCustomerID")	= overruleCID
else
rs("iCustomerID")	= cId
end if
rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
end function
public function getRequestValues()
sName	= convertStr(Request.Form ("sName"))
sValue	= convertStr(Request.Form ("sValue"))
sPrintValue	= convertStr(Request.Form ("sPrintValue"))
sEmailValue	= convertStr(Request.Form ("sEmailValue"))
sMobileValue	= convertStr(Request.Form ("sMobileValue"))
sWAPValue	= convertStr(Request.Form ("sWAPValue"))
bCompress	= convertBool(Request.Form("bCompress"))
end function
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("update tblPage set iTemplateID=null where iTemplateID="& iId)
set rs=nothing
set rs=db.execute("update tblCustomer set defaultTemplate=null where defaultTemplate="& iId)
set rs=nothing
set rs=db.execute("update tblCustomer set iDefaultMobileTemplate=null where iDefaultMobileTemplate="& iId)
set rs=nothing
set rs=db.execute("delete from tblTemplate where iId="& iId)
set rs=nothing
end if
end function
public function copy()
if isNumeriek(iId) then
iId	= null
sName	= l("copyof") & " " & sName
save()
end if
end function
public function copyToCustomerID(oCID,bUpdateCustomer)
if isNumeriek(iId) then
iId	= null
overruleCID	= oCID
save()
if bUpdateCustomer then db.execute("update tblCustomer set defaultTemplate=" & iId & " where iId=" & oCID)
end if
end function
public sub importArtisteer4(folder)
sValue=replace(sValue,"images/flash.swf",folder & "images/flash.swf",1,-1,1)
sValue=replace(sValue,"expressInstall.swf",folder & "expressInstall.swf",1,-1,1)
sValue=replace(sValue,"container.swf",folder & "container.swf",1,-1,1)
sValue=replace(sValue,"href=""style","href=""" & folder & "style",1,-1,1)
sValue=replace(sValue,"src=""","src=""" & folder & "",1,-1,1)
sValue=replace(sValue,"<script src=""" & folder & "http://html5","<script src=""http://html5",1,-1,1)
sValue=replace(sValue,"<script src=""" & folder & "https://html5","<script src=""https://html5",1,-1,1)


sValue=replace(sValue,"<nav ","<nav style=""z-index:999"" ",1,-1,1)
dim  startpos,endpos
dim fso,t_header
set fso=server.CreateObject ("scripting.filesystemobject")
t_header=fso.OpenTextFile(server.MapPath (C_DIRECTORY_QUICKERSITE & "/asp/art_header.txt"),1).ReadAll
t_header=replace(t_header,"[JAVASCRIPT]","",1,-1,1)
svalue=replace(svalue,"</head>","[JAVASCRIPT]" & vbcrlf & "</head>",1,-1,1)
startpos=instr(sValue,"<title>")
endpos=instr(sValue,"</title>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & vbcrlf & vbcrlf & "<!--BEGIN CMS HEADER-->" & vbcrlf & t_header & vbcrlf & "<!-- END CMS HEADER -->" & vbcrlf & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
svalue=replace(svalue,"Enter Site Title","[SITENAME]",1,-1,1)
svalue=replace(svalue,"Enter Site Slogan","[SITESLOGAN]",1,-1,1)
svalue=replace(sValue,"</body>","[GOOGLEANALYTICS]<!--[PAGERENDERTIME]--></body>",1,-1,1)
svalue=replace(svalue,"SLOGAN TEXT","[SITESLOGAN]",1,-1,1)
svalue=replace(svalue,"NAAM","[SITENAME]",1,-1,1)
svalue=replace(svalue,"SLOGAN TEKST","[SITESLOGAN]",1,-1,1)
on error resume next
'if instr(sValue,"Artisteer v3")=0 then
'get rid of specific stylesheet (table related)
dim stylecss,newstylecss
set stylecss=fso.OpenTextFile(server.mappath(folder & "style.css"),1)
newstylecss=stylecss.readAll
set stylecss=nothing
set stylecss=fso.OpenTextFile(server.mappath(folder & "style.css"),2)
if instr(newstylecss,"/* Start tables */")<>0 then 
newstylecss=replace(newstylecss,"/* Start tables */","/* removed by QuickerSite",1,-1,1)
newstylecss=replace(newstylecss,"/* Finish tables */","end remove by QuickerSite */",1,-1,1)
else
newstylecss=replace(newstylecss,".art-article table",vbcrlf & "/* removed by QuickerSite" & vbcrlf &".art-article table",1,1,1)
newstylecss=replace(newstylecss,vbcrlf & "pre",vbcrlf & "end remove by QuickerSite */" & vbcrlf & "pre",1,-1,1)
end if
stylecss.writeline(newstylecss)
stylecss.close
set stylecss=nothing
'end if
on error goto 0
set fso=nothing
end sub
public sub treatAsJSTemplate(folder)
sValue=replace(sValue,"<h3>Search</h3>","<h3>[QSSEARCHLABEL]</h3>",1,-1,1)
sValue=replace(sValue,"href=""style","href=""" & folder & "style",1,-1,1)
sValue=replace(sValue,"href=""js/superfish","href=""" & folder & "js/superfish",1,-1,1)
sValue=replace(sValue,"href=""css/JQueryUI","href=""" & folder & "css/JQueryUI",1,-1,1)
sValue=replace(sValue,"<meta charset=""utf-8"" />","",1,-1,1)
sValue=replace(sValue,"<meta content="""" name=""description"" />","",1,-1,1)
sValue=replace(sValue,"<meta content="""" name=""keywords"" />","",1,-1,1)
sValue=replace(sValue,"// initialise superfish menus","",1,-1,1)
sValue=replace(sValue,"<!-- preheader -->","<!-- preheader - not used in QuickerSite - but you can use it to link to specific pages (remove the visibility:hidden) -->" & vbcrlf,1,-1,1)
sValue=replace(sValue,"<div id=""preheader"" class=""cf"">","<div style=""visibility:hidden"" id=""preheader"" class=""cf"">",1,-1,1)  
sValue=replace(sValue,"src=""//","DONOTREPLACESRC1",1,-1,1)
sValue=replace(sValue,"src=""http","DONOTREPLACESRC2",1,-1,1)
sValue=replace(sValue,"src=""","src=""" & folder & "",1,-1,1)
sValue=replace(sValue,"DONOTREPLACESRC1","src=""//",1,-1,1)
sValue=replace(sValue,"DONOTREPLACESRC2","src=""http",1,-1,1)
dim  startpos,endpos
dim fso,t_header
set fso=server.CreateObject ("scripting.filesystemobject")
t_header=fso.OpenTextFile(server.MapPath (C_DIRECTORY_QUICKERSITE & "/asp/art_header.txt"),1).ReadAll
t_header=replace(t_header,"[JAVASCRIPT]","",1,-1,1)
svalue=replace(svalue,"</head>","[JS_JAVASCRIPT]" & vbcrlf & "</head>",1,-1,1)
startpos=instr(sValue,"<title>")
endpos=instr(sValue,"</title>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & vbcrlf & vbcrlf & "<!--BEGIN CMS HEADER-->" & vbcrlf & t_header & vbcrlf & "<!-- END CMS HEADER -->" & vbcrlf & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
svalue=replace(svalue,"<div id=""headline"">JStemplates.com</div>","<div id=""headline"">[SITENAME]</div>",1,-1,1)
startpos=instr(sValue,"<div id=""slogan"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "<div id=""slogan"">[SITESLOGAN]</div>" &  right(sValue,len(sValue)-endpos-6)
end if
end if
'svalue=replace(svalue,"<div id=""slogan"">HTML5 &amp; CSS3 JavaScript&reg; template generator</div>","<div id=""slogan"">[SITESLOGAN]</div>",1,-1,1)
svalue=replace(sValue,"</body>","[GOOGLEANALYTICS]<!--[PAGERENDERTIME]--></body>",1,-1,1)
svalue=replace(sValue,"<h3>Title of page</h3>","<h3>[PAGETITLE]</h3>",1,-1,1)
svalue=replace(sValue,"<h3>News</h3>","<h3>[QSHIGHLIGHTSLABEL]</h3>",1,-1,1)
svalue=replace(sValue,"<h3>Contact</h3>","<h3>[QSCONTACTINFOLABEL]</h3>",1,-1,1)
svalue=replace(sValue,"<div id=""footerContentText"">This is the footer and a <a href=""#"">link</a>.</div>","<div id=""footerContentText"">[QSSITEFOOTER]</div>",1,-1,1)
startpos=instr(sValue,"<!-- START SEARCHFORM -->")
if startpos<>0 then
endpos=instr(startpos,sValue,"<!-- END SEARCHFORM -->")
if startpos<>0 then
if instr(sValue,"class=""searchbutton""")=0 then
	svalue=left(sValue,startpos-1) & vbcrlf & "<form method=""post"" name=Search action=""[SF_ACTIONURL]"">[SF_HIDDENFIELD]<input maxlength='100' style=""width:95%"" type='text' name='svalue' id='s' value=""[SF_TEXTVALUE]"" onclick=""javascript:this.select();""><input type=""submit"" value=""[QSSEARCHLABEL]"" name=""search""></form>" & vbcrlf & right(sValue,len(sValue)-endpos-23)
else
	svalue=left(sValue,startpos-1) & vbcrlf & "<div style=""text-align:center""><form method=""post"" name=Search action=""[SF_ACTIONURL]"">[SF_HIDDENFIELD]<input maxlength='100' class=""searchfield"" type='text' name='svalue' id='s' value=""[SF_TEXTVALUE]"" onclick=""javascript:this.select();""><input type=""submit"" class=""searchbutton"" value=""&#xf002;"" /></form></div>" & vbcrlf & right(sValue,len(sValue)-endpos-23)
end if
end if
end if
startpos=instr(sValue,"<!-- vermenuUL -->")
if startpos<>0 then
endpos=instr(startpos,sValue,"<!-- /vermenuUL -->")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[JS_VMENU]" &  right(sValue,len(sValue)-endpos-19)
end if
end if
startpos=instr(sValue,"<!-- hormenuUL1 -->")
if startpos<>0 then
endpos=instr(startpos,sValue,"<!-- /hormenuUL1 -->")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[JS_HMENU1]" &  right(sValue,len(sValue)-endpos-20)
end if
end if
startpos=instr(sValue,"<!-- hormenuUL2 -->")
if startpos<>0 then
endpos=instr(startpos,sValue,"<!-- /hormenuUL2 -->")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[JS_HMENU2]" &  right(sValue,len(sValue)-endpos-20)
end if
end if
while instr(svalue,"<p>")<>0
startpos=instr(sValue,"<p>")
if startpos<>0 then
endpos=instr(startpos,sValue,"</p>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-4)
end if
end if
wend
while instr(svalue,"<!--maincontent-->")<>0
startpos=instr(sValue,"<!--maincontent-->")
if startpos<>0 then
endpos=instr(startpos,sValue,"<!--/maincontent-->")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[PAGEBODY][PAGELIST][PAGEFORM][PAGECATALOG][PAGEFEED][PAGEAPPLICATION][PAGETHEME]<div class=""cf""></div><div style=""margin-top:20px;font-size:smaller"">[PAGEBREADCRUMBS]</div>" & right(sValue,len(sValue)-endpos-19)
end if
end if
wend
while instr(svalue,"<!-- banner 1-->")<>0
startpos=instr(sValue,"<!-- banner 1-->")
if startpos<>0 then
endpos=instr(startpos,sValue,"<!-- /banner 1-->")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[QSHIGHLIGHTS]" & right(sValue,len(sValue)-endpos-17)
end if
end if
wend
while instr(svalue,"<!-- banner 2-->")>0
startpos=instr(sValue,"<!-- banner 2-->")
if startpos<>0 then
endpos=instr(startpos,sValue,"<!-- /banner 2-->")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[BANNER]" & right(sValue,len(sValue)-endpos-17)
end if
end if
wend
svalue=replace(svalue,"<!--", vbcrlf & "<!--",1,-1,1)
svalue=replace(svalue,"-->", "-->" & vbcrlf,1,-1,1)
end sub
public sub treatAsArtisteer(folder)
if instr(sValue,"JStemplates.com")<>0 then
treatAsJSTemplate(folder)
exit sub
end if
if instr(sValue,"Artisteer")<>0 then
dim sOrigValue
sOrigValue=sValue
if instr(sValue,"Artisteer v4")<>0 then
importArtisteer4 folder
exit sub
end if
'replace clearing divs
sValue=replace(sValue,"<div class=""cleared""></div>","**QSREPLACEEMPTYDIVS**",1,-1,1)
'hacks for v3 beta
while instr(svalue,"<center>")<>0
startpos=instr(sValue,"<center>")
if startpos<>0 then
endpos=instr(startpos,sValue,"</center>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-9)
end if
end if
wend
while instr(svalue,"<h2 class=""art-postheader"">")<>0
startpos=instr(sValue,"<h2 class=""art-postheader"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</h2>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "###PAGETITLE###" & right(sValue,len(sValue)-endpos+1)
end if
end if
wend
while instr(svalue,"<div style")<>0
startpos=instr(sValue,"<div style")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-6)
end if
end if
wend 
svalue=replace(sValue,"###PAGETITLE###","<h2 class=""art-postheader"">[PAGETITLE]",1,-1,1)
startpos=instr(sValue,"<h2 class=""art-logo-text"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</h2>")
svalue=left(sValue,startpos-1) & "<h2 class=""art-logo-text"">[SITESLOGAN]</h2>" &  right(sValue,len(sValue)-endpos-4)
end if
startpos=instr(sValue,"<h1 class=""art-logo-name"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</h1>")
svalue=left(sValue,startpos-1) & "<h1 class=""art-logo-name"">[SITENAME]</h1>" &  right(sValue,len(sValue)-endpos-4)
end if
'footer down the page
startpos=instr(sValue,"<p class=""art-page-footer"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</p>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "<p class=""art-page-footer"">" & customfooter & "</p>" &  right(sValue,len(sValue)-endpos-4)
end if
end if
'weird addition in Art3
dim bart3
bart3=false
if convertBool(bSetCustomHL) then
startpos=instr(sValue,"<ul><li><b>Jun 14,")
if startpos<>0 then
svalue=replace(sValue,"Highlights","[QSHIGHLIGHTSLABEL]",1,-1,1)
endpos=instr(startpos,sValue,"</ul>")
bart3=true
svalue=left(sValue,startpos-1) & "[QSHIGHLIGHTS]" & vbcrlf & right(sValue,len(sValue)-endpos-5)
end if
startpos=instr(sValue,"<ul>" & vbcrlf & "    <li>" & vbcrlf & "      <b>Jun 14,")
if startpos<>0 then
svalue=replace(sValue,"Highlights","[QSHIGHLIGHTSLABEL]",1,-1,1)
endpos=instr(startpos,sValue,"**QSREPLACEEMPTYDIVS**")
bart3=true
if instr(sValue,"[QSHIGHLIGHTSLABEL]")<>0 then
svalue=left(sValue,startpos-1) & "[QSHIGHLIGHTS]" & vbcrlf & right(sValue,len(sValue)-endpos+1)
else
svalue=left(sValue,startpos-1) & "[QSHIGHLIGHTS]</div>" & vbcrlf & right(sValue,len(sValue)-endpos+1)
end if
end if
end if
'get rid of art-postfootericons
while instr(svalue,"art-postfootericons")<>0
startpos=instr(sValue,"<div class=""art-postfootericons")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-6)
end if
end if
wend
'get rid of art-postheadericons
while instr(svalue,"art-postheadericons")<>0
startpos=instr(sValue,"<div class=""art-postheadericons")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-6)
end if
end if
wend
'get rid of title
while instr(svalue,">Text, ")<>0
startpos=instr(sValue,">Text, ")
if startpos<>0 then
endpos=instr(startpos,sValue,"</h2>")
if startpos<>0 then
svalue=left(sValue,startpos) & "[PAGETITLE]</h2>" & right(sValue,len(sValue)-endpos-5)
end if
end if
wend
'v3.0
if instr(sValue,"Artisteer v3")<>0 then
startpos=instr(sValue,"<div class=""art-postcontent"">")
endpos=instr(sValue,"</table>")
if startpos<>0 and endpos<>0 then
svalue=left(sValue,startpos-1) & "<div class=""art-postcontent"">[PAGEBODY][PAGELIST][PAGEFORM][PAGECATALOG][PAGEFEED][PAGEAPPLICATION][PAGETHEME]<div class=""cleared""></div><div style=""margin-top:20px;font-size:smaller"">[PAGEBREADCRUMBS]</div>" & right(sValue,len(sValue)-endpos-8)
end if
startpos=instr(startpos,sValue,"<p>")
if startpos<>0 then
endpos=instr(startpos,sValue,"</p>")
if endpos<>0 then
svalue=left(sValue,startpos-1) & "" & right(sValue,len(sValue)-endpos-4)
end if
end if
end if
sValue=replace(sValue,"images/flash.swf",folder & "images/flash.swf",1,-1,1)
sValue=replace(sValue,"expressInstall.swf",folder & "expressInstall.swf",1,-1,1)
sValue=replace(sValue,"container.swf",folder & "container.swf",1,-1,1)
sValue=replace(sValue,"href=""style","href=""" & folder & "style",1,-1,1)
sValue=replace(sValue,"src=""","src=""" & folder & "",1,-1,1)
dim fso,startpos,endpos,t_header,t_searchform
set fso=server.CreateObject ("scripting.filesystemobject")
'v24
startpos=instr(sValue,"<!-- article-content -->")
endpos=instr(sValue,"<!-- /article-content -->")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[PAGEBODY][PAGELIST][PAGEFORM][PAGECATALOG][PAGEFEED][PAGEAPPLICATION][PAGETHEME]<div class=""cleared""></div><div style=""margin-top:20px;font-size:smaller"">[PAGEBREADCRUMBS]</div><!-- article-content -->" & right(sValue,len(sValue)-endpos-24)
end if
startpos=instr(sValue,"<!-- article-content -->")
endpos=instr(sValue,"<!-- /article-content -->")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "" & right(sValue,len(sValue)-endpos-24)
end if
'get data
t_header=fso.OpenTextFile(server.MapPath (C_DIRECTORY_QUICKERSITE & "/asp/art_header.txt"),1).ReadAll 
if instr(sValue,"Artisteer v3")<>0 then
t_searchform=fso.OpenTextFile(server.MapPath (C_DIRECTORY_QUICKERSITE & "/asp/art_searchform_v3.txt"),1).ReadAll 
else
t_searchform=fso.OpenTextFile(server.MapPath (C_DIRECTORY_QUICKERSITE & "/asp/art_searchform.txt"),1).ReadAll 
end if
'hack for the gallery problem with Artisteer 3x
t_header=replace(t_header,"[JAVASCRIPT]","",1,-1,1)
svalue=replace(svalue,"</head>","[JAVASCRIPT]" & vbcrlf & "</head>",1,-1,1)
'############################################################################
'############################################################################
'############################################################################
'header
startpos=instr(sValue,"<title>")
endpos=instr(sValue,"</title>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & vbcrlf & vbcrlf & "<!--BEGIN CMS HEADER-->" & vbcrlf & t_header & vbcrlf & "<!-- END CMS HEADER -->" & vbcrlf & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
dim re,matches,match
Set re = new regexp  
re.IgnoreCase	=true
re.Global	=false
'############################################################################
'############################################################################
'############################################################################
'remove headings
dim h
for h=1 to 7
re.Pattern ="\<h" & h & "\>([^\]]+)\<\/h" & h & "\>"
Set matches = re.Execute(sValue)
For Each match In matches
sValue=re.Replace (sValue,"")
Next
next
'############################################################################
'############################################################################
'############################################################################
'remove tables
re.Pattern ="\<table([^\]]+)\<\/table\>"
Set matches = re.Execute(sValue)
For Each match In matches
sValue=re.Replace (sValue,"")
Next
'############################################################################
'############################################################################
'############################################################################
'search form
svalue=replace(sValue,"Newsletter",l("search"),1,-1,1)
svalue=replace(sValue,"Subscribe",l("search"),1,-1,1)
startpos=instr(sValue,"<form ")
if startpos<>0 then
endpos=instr(sValue,"</form>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & t_searchform & right(sValue,len(sValue)-endpos-6)
end if
else
startpos=instr(sValue,"<input type=""text""")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
svalue=left(sValue,startpos-1) & t_searchform & right(sValue,len(sValue)-endpos+1)
end if
end if
while instr(svalue,"<blockquote")<>0
startpos=instr(sValue,"<blockquote")
if startpos<>0 then
endpos=instr(startpos,sValue,"</blockquote>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-13)
end if
end if
wend 
while instr(svalue,"<p style")<>0
startpos=instr(sValue,"<p style")
if startpos<>0 then
endpos=instr(startpos,sValue,"</p>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-4)
end if
end if
wend 
while instr(svalue,"<ul>")<>0
startpos=instr(sValue,"<ul>")
if startpos<>0 then
endpos=instr(startpos,sValue,"</ul>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-5)
end if
end if
wend
'while instr(svalue,"<div class=""art-layout-cell layout-item-0""")<>0
'	startpos=instr(sValue,"<div class=""art-layout-cell layout-item-0""")
'	if startpos<>0 then
'	endpos=instr(startpos,sValue,"</div>")
'	if startpos<>0 then
'svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-6)
'	end if
'	end if
'wend
 
'############################################################################
'############################################################################
'############################################################################
'remove line at the bottom
if not isLeeg(customFooter) then
svalue=replace(sValue,"<p class=""art-page-footer""><a href=""http://www.artisteer.com/"">Web Template</a> created with Artisteer.</p>",customFooter,1,-1,1)
else
svalue=replace(sValue,"<p class=""art-page-footer""><a href=""http://www.artisteer.com/"">Web Template</a> created with Artisteer.</p>","",1,-1,1)
end if
'############################################################################
'############################################################################
'############################################################################
'pagemenu
startpos=instr(sValue,"<ul class=""art-menu"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[QS_ARTISTEER_FULLMENU_V22]</div>" & right(sValue,len(sValue)-endpos-6)
end if
end if
startpos=instr(sValue,"<ul class=""art-hmenu"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[QS_ARTISTEER_FULLMENU_V3]</div>" & right(sValue,len(sValue)-endpos-6)
end if
end if
startpos=instr(sValue,"art-vmenublockcontent-body"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
if instr(sValue,"v3.0.0.37990")<>0 or instr(sValue,"v3.0.0.39952")<>0 or instr(sValue,"v3.0.0.41778")<>0 or instr(sValue,"v3.1.")<>0 then
svalue=left(sValue,startpos-1) & "art-vmenublockcontent-body"">[QS_ARTISTEER_FULLVMENU_V3]" & right(sValue,len(sValue)-endpos+1)
else
svalue=left(sValue,startpos-1) & "art-vmenublockcontent-body"">[QS_ARTISTEER_FULLVMENU_V24]" & right(sValue,len(sValue)-endpos+1)
end if
end if
end if
'include menu
re.Pattern ="\<ul class=""art-menu""\>([^\]]+)\<\/ul\>"
Set matches = re.Execute(sValue)
For Each match In matches
sValue=re.Replace (sValue,"[QS_ARTISTEER_FULLMENU_V22]")
Next
'v24
re.Pattern ="\<ul class=""art-vmenu""\>([^\]]+)\<\/ul\>"
Set matches = re.Execute(sValue)
For Each match In matches
sValue=re.Replace (sValue,"[QS_ARTISTEER_FULLVMENU_V24]")
Next
'############################################################################
'############################################################################
'############################################################################
'pagebody
if instr(sValue,"Artisteer v3")=0 then
startpos=instr(sValue,"<div class=""art-PostContent"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "<div class=""art-PostContent"">[PAGEBODY][PAGELIST][PAGEFORM][PAGECATALOG][PAGEFEED][PAGEAPPLICATION][PAGETHEME]</div>" & vbcrlf & "<div class=""cleared""></div>" & vbcrlf & "<div style=""margin-top:20px;font-size:smaller"">[PAGEBREADCRUMBS]</div>" & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
end if
end if 
if convertBool(bSetFooterVar) then
startpos=instr(sValue,"<div class=""art-Footer-text"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if convertBool(bSetRSSLink) then
svalue=replace(svalue,"<a href=""#"" class=""art-rss-tag-icon"" title=""RSS""></a>","[RSSLINKART]",1,-1,1)
svalue=left(sValue,startpos-1) & "<div class=""art-Footer-text"">[QSSITEFOOTER]</div>" & vbcrlf & right(sValue,len(sValue)-endpos-7)
else
svalue=left(sValue,startpos-1) & "<div class=""art-Footer-text"">[QSSITEFOOTER]</div>" & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
end if
'v24
startpos=instr(sValue,"<div class=""art-footer-text"">")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "<div class=""art-footer-text"">[QSSITEFOOTER]</div>" & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
end if
end if
if convertBool(bSetContactVar) then
startpos=instr(sValue,"<img src=""" & folder & "images/contact.jpg")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[BANNER]</div>" & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
end if
svalue=replace(sValue,"Contact Info","[QSCONTACTINFOLABEL]",1,-1,1)
startpos=instr(sValue,"<img src=""" & folder & "./images/contact.jpg")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[BANNER]</div>" & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
end if
'svalue=replace(sValue,"Contact Info","[QSCONTACTINFOLABEL]",1,-1,1)
end if
if convertBool(bSetCustomHL) and not bart3 then
svalue=replace(sValue,"Highlights","[QSHIGHLIGHTSLABEL]",1,-1,1)
startpos=instr(sValue,"<p><b>Jun 14")
if startpos<>0 then
endpos=instr(startpos,sValue,"</div>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & "[QSHIGHLIGHTS]</div>" & vbcrlf & right(sValue,len(sValue)-endpos-7)
end if
end if
end if
while instr(svalue,"<p>")<>0
startpos=instr(sValue,"<p>")
if startpos<>0 then
endpos=instr(startpos,sValue,"</p>")
if startpos<>0 then
svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-4)
end if
end if
wend
if convertBool(bSetRSSLink) then
svalue=replace(svalue,"<a href=""#"" class=""art-rss-tag-icon"" title=""RSS""></a>","[RSSLINKART]",1,-1,1)
end if
if instr(svalue,"[RSSLINKART]")=0 and instr(sOrigValue,"art-rss-tag-icon")<>0 then
svalue=replace(svalue,"<div class=""art-footer-text"">","[RSSLINKART]<div class=""art-footer-text"">",1,-1,1)
end if
'startpos=instr(sValue,"<ul>")
'if startpos<>0 then
'	endpos=instr(startpos,sValue,"</ul>")
'	if startpos<>0 then
'	svalue=left(sValue,startpos-1) & right(sValue,len(sValue)-endpos-7)
'	'	end if
'end if
if instr(sValue,"[PAGEBODY]")=0 then
svalue=replace(svalue,"<div class=""art-postcontent"">","<div class=""art-postcontent"">[PAGEBODY][PAGELIST][PAGEFORM][PAGECATALOG][PAGEFEED][PAGEAPPLICATION][PAGETHEME]" & vbcrlf & "<div class=""cleared""></div>" & vbcrlf & "<div style=""margin-top:20px;font-size:smaller"">[PAGEBREADCRUMBS]</div>",1,-1,1)
if instr(sValue,"v3.1")<>0 then
svalue=replace(svalue,"[QS_ARTISTEER_FULLMENU_V3]","[QS_ARTISTEER_FULLMENU_V31]",1,-1,1)
svalue=replace(svalue,"[QS_ARTISTEER_FULLVMENU_V3]","[QS_ARTISTEER_FULLVMENU_V31]",1,-1,1)
end if
end if
'various replaces
svalue=replace(svalue,"Enter Site Title","[SITENAME]",1,-1,1)
svalue=replace(svalue,"Enter Site Slogan","[SITESLOGAN]",1,-1,1)
svalue=replace(sValue,"Welcome","[PAGETITLE]",1,-1,1)
svalue=replace(sValue,"</body>","[GOOGLEANALYTICS]<!--[PAGERENDERTIME]--></body>",1,-1,1)
svalue=replace(sValue,"Search",l("search"),1,-1,1)
svalue=replace(sValue,l("search"),"[QSSEARCHLABEL]",1,-1,1)
svalue=replace(svalue,"HEADLINE","[SITENAME]",1,-1,1)
svalue=replace(svalue,"SLOGAN TEXT","[SITESLOGAN]",1,-1,1)
svalue=replace(svalue,"NAAM","[SITENAME]",1,-1,1)
svalue=replace(svalue,"SLOGAN TEKST","[SITESLOGAN]",1,-1,1)
sValue=replace(sValue,"**QSREPLACEEMPTYDIVS**","<div class=""cleared""></div>",1,-1,1)
'fix for html templates
'startpos=instr(svalue,">E html PUBLIC")
'if startpos<>0 then
'	endpos=instr(startpos,sValue,"</blockquote>")
'	svalue=left(sValue,startpos-1) & "<div><div><div>" & right(sValue,len(sValue)-endpos-12)
'end if
if instr(sValue,"art-Footer-text")<>0 then
sValue=replace(sValue,"<div class=""art-Footer-text"">[QSSITEFOOTER]</div>","[QSSITEFOOTER]</div></div>",1,-1,1)
end if
sValue=replace(sValue,"<meta name=""keywords"" content=""Keywords"" />","",1,-1,1)
sValue=replace(sValue,"<meta name=""description"" content=""Description"" />","",1,-1,1)
on error resume next
'if instr(sValue,"Artisteer v3")=0 then
'get rid of specific stylesheet (table related)
dim stylecss,newstylecss
set stylecss=fso.OpenTextFile(server.mappath(folder & "style.css"),1)
newstylecss=stylecss.readAll
set stylecss=nothing
set stylecss=fso.OpenTextFile(server.mappath(folder & "style.css"),2)
if instr(newstylecss,"/* Start tables */")<>0 then 
newstylecss=replace(newstylecss,"/* Start tables */","/* removed by QuickerSite",1,-1,1)
newstylecss=replace(newstylecss,"/* Finish tables */","end remove by QuickerSite */",1,-1,1)
else
newstylecss=replace(newstylecss,".art-article table",vbcrlf & "/* removed by QuickerSite" & vbcrlf &".art-article table",1,1,1)
newstylecss=replace(newstylecss,vbcrlf & "pre",vbcrlf & "end remove by QuickerSite */" & vbcrlf & "pre",1,-1,1)
end if
stylecss.writeline(newstylecss)
stylecss.close
set stylecss=nothing
'end if
on error goto 0
set fso=nothing
end if
end sub
end class%>
