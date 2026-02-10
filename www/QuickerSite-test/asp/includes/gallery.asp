
<%class cls_gallery
Public iId, sName, sCode, sPath, iType, iThumbSize, iFullImageSize, iPicsInRow, sStyleTable, bShowFileName, sStyleImage, sStyleTableCell, backsitePV, iBrowseBy, sFullImage, iSlideShowTimerQS, bFSR, bAutoStartSS, bRandom, sCustomLink
Public sType, sBorder, sBorderColor, sWidth, sHeight,  bOpenInNewWindow, iSpecialEffect, slideshowC, sCycleEffect, iSortImagesBy, sNSCss, iFixedH, sNSImgLinks
Public bNSControlNav,bNSdirectionNav
Private p_sNextLink, p_sPreviousLink, slideshow
Private Sub Class_Initialize
On Error Resume Next
iId	= null
iThumbSize	= 150
iFullImageSize	= 640
sStyleTable	= "margin:0 auto;border-style:none"
sStyleImage	= "border-style:none;padding:0px;margin:0px"
sStyleTableCell	= "padding:3px"
bShowFileName	= true
iPicsInRow	= 3
backsitePV	= false
iBrowseBy	= 9
sFullImage	= "Click image for full version"
iSlideShowTimerQS	= 3
bFSR	= false
bAutoStartSS	= false
p_sNextLink	= l("next") & "&nbsp;&gt;&gt;"
p_sPreviousLink	= "&lt;&lt;&nbsp;" & l("previous") 
bRandom	= false
sCustomLink	= ""
bOpenInNewWindow	= true
sCycleEffect	= "fade"
sHeight	= "150"
sWidth	= "150"

sBorder	= "1"
sBorderColor	= "#CCCCCC"
sNSCss="default"
bNSControlNav=true
bNSdirectionNav=true
sNSImgLinks=""
pick(decrypt(request("iGalleryID")))

On Error Goto 0
end sub
Public property let sNextLink(value)
p_sNextLink=value
End property
Public property get sNextLink
if backsitePV then
sNextLink=""
else
sNextLink=p_sNextLink
end if
end property
Public property let sPreviousLink(value)
p_sPreviousLink=value
End property
Public property get sPreviousLink
if backsitePV then
sPreviousLink=""
else
sPreviousLink=p_sPreviousLink
end if
end property
Public Function Pick(id)
dim sql, RS
if isNumeriek(id) then
sql = "select * from tblGallery where iCustomerID="&cid&" and iId=" & id
set RS = db.execute(sql)
if not rs.eof then
iId	= rs("iId")
sName	= rs("sName")
sCode	= rs("sCode")
sPath	= rs("sPath")
iType	= rs("iType")
iThumbSize	= rs("iThumbSize")
iFullImageSize	= rs("iFullImageSize")
iPicsInRow	= rs("iPicsInRow")
sStyleTable	= rs("sStyleTable")
sStyleImage	= rs("sStyleImage")
bShowFileName	= rs("bShowFileName")
sStyleTableCell	= rs("sStyleTableCell")
iBrowseBy	= rs("iBrowseBy")
sFullImage	= rs("sFullImage")
iSlideShowTimerQS	= rs("iSlideShowTimerQS")
bFSR	= rs("bFSR")
p_sNextLink	= rs("sNextLink")
p_sPreviousLink	= rs("spreviousLink")
bAutoStartSS	= rs("bAutoStartSS")
bRandom	= false
sCustomLink	= rs("sCustomLink")
sType	= rs("sType")
sHeight	= rs("sHeight")
sWidth	= rs("sWidth")
sBorder	= rs("sBorder")
sBorderColor	= rs("sBorderColor")
bOpenInNewWindow	= rs("bOpenInNewWindow")
iSpecialEffect	= rs("iSpecialEffect")
sCycleEffect	= rs("sCycleEffect")
iSortImagesBy	= rs("iSortImagesBy")
sNSCss	= rs("sNSCss")
bNSControlNav= rs("bNSControlNav")
bNSdirectionNav= rs("bNSdirectionNav")
sNSImgLinks=rs("sNSImgLinks")


if convertBool(bRandom) then iSortImagesBy=3
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
if isLeeg(sPath) then
check=false
message.AddError("err_mandatory")
end if
if isLeeg(sCode) then
check=false
message.AddError("err_mandatory")
end if
if not isLeeg(sCode) then
dim checkRS
set checkRS=db.execute("select count(iId) from tblGallery where iId<>" & convertGetal(iId) & " and iCustomerID=" & cID & " and sCode='" & cleanup(ucase(sCode)) & "'")
if clng(checkRS(0))>0 then
check=false
message.AddError("err_doublefeed")
end if
set checkRS=nothing
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
rs.Open "select * from tblGallery where 1=2"
rs.AddNew
rs("dCreatedTS")=now()
else
rs.Open "select * from tblGallery where iId="& iId
end if
rs("sName")	= left(sName,50)
rs("sCode")	= sCode
rs("sPath")	= sPath
rs("iType")	= iType
rs("iThumbSize")	= iThumbSize
rs("iFullImageSize")	= iFullImageSize
rs("iPicsInRow")	= iPicsInRow
rs("bShadowThumb")	= false
rs("sStyleTable")	= sStyleTable
rs("sStyleImage")	= sStyleImage
rs("bShowFileName")	= bShowFileName
rs("iCustomerID")	= cId
rs("sStyleTableCell")	= sStyleTableCell
rs("iBrowseBy")	= iBrowseBy
rs("sFullImage")	= sFullImage
rs("iSlideShowTimerQS")	= iSlideShowTimerQS
rs("bFSR")	= bFSR
rs("sNextLink")	= p_sNextLink
rs("sPreviousLink")	= p_sPreviousLink
rs("bAutoStartSS")	= bAutoStartSS
rs("bRandom")	= bRandom
rs("sCustomLink")	= sCustomLink
rs("sType")	= sType
rs("sHeight")	= sHeight
rs("sWidth")	= sWidth
rs("sBorder")	= sBorder
rs("sBorderColor")	= sBorderColor
rs("bOpenInNewWindow")	= bOpenInNewWindow
rs("iSpecialEffect")	= iSpecialEffect
rs("sCycleEffect")	= sCycleEffect
rs("iSortImagesBy")	= iSortImagesBy
rs("sNSCss") = sNSCss
rs("bNSdirectionNav")= bNSdirectionNav
rs("bNScontrolNav")= bNScontrolNav
rs("sNSImgLinks") = sNSImgLinks

rs.Update 
iId = convertGetal(rs("iId"))
rs.close
Set rs = nothing
customer.cacheGalleries()
application(QS_CMS_cacheGallery & iId)=""
end function
public function getRequestValues()
sName	= convertStr(Request.Form ("sName"))
sPath	= convertStr(Request.Form ("sPath"))
iType	= convertGetal(Request.Form ("iType"))
iThumbSize	= convertGetal(Request.Form ("iThumbSize"))
iFullImageSize	= convertGetal(Request.Form ("iFullImageSize"))
iPicsInRow	= convertGetal(Request.Form ("iPicsInRow"))
sCode	= ucase(convertStr(Request.Form ("sCode")))
bShowFileName	= convertBool(Request.Form ("bShowFileName"))
sStyleTable	= convertStr(Request.Form ("sStyleTable"))
sStyleImage	= convertStr(Request.Form ("sStyleImage"))
sStyleTableCell	= convertStr(Request.Form ("sStyleTableCell"))
iBrowseBy	= convertGetal(Request.Form ("iBrowseBy"))
sFullImage	= convertStr(Request.Form ("sFullImage"))
iSlideShowTimerQS	= convertGetal(Request.Form ("iSlideShowTimerQS"))
bFSR	= convertBool(Request.Form ("bFSR"))
p_sNextLink	= convertStr(Request.Form ("sNextLink"))
p_sPreviousLink	= convertStr(Request.Form ("sPreviousLink"))
bAutoStartSS	= convertBool(Request.Form("bAutoStartSS"))
bRandom	= convertBool(Request.Form ("bRandom"))
sCustomLink	= convertStr(Request.Form ("sCustomLink"))
sType	= convertStr(Request.Form ("sType"))
sHeight	= convertStr(Request.Form ("sHeight"))
sWidth	= convertStr(Request.Form ("sWidth"))
sBorder	= convertStr(Request.Form ("sBorder"))
sBorderColor	= convertStr(Request.Form ("sBorderColor"))
bOpenInNewWindow	= convertBool(Request.Form ("bOpenInNewWindow"))
iSpecialEffect	= convertGetal(request.form("iSpecialEffect"))
sCycleEffect	= convertStr(request.form("sCycleEffect"))
iSortImagesBy	= convertGetal(request.form("iSortImagesBy"))
sNSCss=convertStr(request.form("sNSCss"))
bNSControlNav= convertBool(Request.Form("bNSControlNav"))
bNSdirectionNav= convertBool(Request.Form("bNSdirectionNav"))

if isLeeg(sCycleEffect) then sCycleEffect="fade"
if convertGetal(iId)=0 then
if sHeight="" and sWidth="" then
sHeight=150
sWidth=150
end if
end if

if isLeeg(iId) then
	if request.form("sType")=QS_gallery_NS then
		if request.form("sHeight")="" and request.form("sWidth")="" then
			sHeight	= 0
			sWidth	= 0
		end if
	end if
end if

'save captions and urls in case of nivo slider
if sType=QS_gallery_NS then
sNSImgLinks=""
dim fso
set fso=server.createobject("scripting.filesystemobject")

dim imageFolder,imageFolderPath
imageFolderPath=C_VIRT_DIR & application("QS_CMS_userfiles") & sPath
imageFolderPath=replace(imageFolderPath,"//","/",1,-1,1)
set imageFolder=fso.getFolder(server.mappath(imageFolderPath))

dim file
for each file in  imageFolder.files
if allowedFileTypesforThumbing.exists(lcase(GetFileExtension(file.name))) then
sNSImgLinks=sNSImgLinks & removeDots(file.name) & vbtab & request.form("caption" & removeDots(sanitize(file.name))) & vbtab & request.form("url" & removeDots(sanitize(file.name))) & vbtab & request.form("oinw" & removeDots(sanitize(file.name))) & vbcrlf
end if
next

'response.clear
'response.write sNSImgLinks
'response.end 

end if

end function
Public property get iSlideShowTimerQSDefault
if convertGetal(iSlideShowTimerQS)=0 then
iSlideShowTimerQSDefault=4
else
iSlideShowTimerQSDefault=iSlideShowTimerQS
end if
end property
public function remove
if not isLeeg(iId) then
dim rs
set rs=db.execute("delete from tblGallery where iId="& iId)
set rs=nothing
customer.cacheGalleries()
application(QS_CMS_cacheGallery & iId)=""
end if
end function



public function build
 
fHeight=sHeight
sHeight=""
fWidth=sWidth
sWidth="" 
 
build=""
slideshow=""
slideshowC=""
dim genPW
genPW=generatePassWord&generatePassWord
if convertGetal(iId)<>0 then
'On Error Resume Next

dim files, folder, fso, sFolder, file, picLink
sFolder	= C_VIRT_DIR & Application("QS_CMS_userfiles") & sPath
sFolder=replace(sFolder,"//","/",1,-1,1)
set fso	= server.CreateObject ("scripting.filesystemobject")
if not fso.FolderExists (server.MapPath (sFolder)) then
build="<font color='Red'><b>" & server.MapPath (sFolder) & "</b> <u>does not exist</u>!</font>"
else
set folder	= fso.GetFolder (server.MapPath (sFolder))
set files	= folder.files
if files.count>0 then
build=""
dim filesDict
set filesDict=server.CreateObject ("scripting.dictionary")
'set random?
dim randomKey
for each file in files
'response.write convertCalcDateTime(file.DateCreated)  & "<br />"
if bRandom then iSortImagesBy=3
select case convertgetal(iSortImagesBy)
case 0
randomKey=file.name&generatePassword
case 1
randomKey=(99999999999999-(convertGetal(convertCalcDateTime(file.dateLastModified)))) & generatePassword
case 2
randomKey=convertCalcDateTime(file.dateLastModified) & generatePassword
case 4
randomKey=(99999999999999-(convertGetal(convertCalcDateTime(file.DateCreated)))) & generatePassword
case 5
randomKey=convertCalcDateTime(file.DateCreated) & generatePassword
case 3 'random
Randomize()
randomKey=cstr(Int(500 * Rnd + 1) & generatePassword & file.name)
end select
if not filesDict.Exists (randomKey) then filesDict.Add randomKey,file.name
next
set files=nothing
SortDictionary2 filesDict,1
dim counter, showNextLink, qs_gallery_start, colCounter, additionalImages
counter=1
colCounter=0
showNextLink=false
qs_gallery_start=convertGetal(Request.Querystring ("qs_gallery_start"))
if qs_gallery_start<0 then qs_gallery_start=0
if convertGetal(decrypt(Request.QueryString ("iGalleryID")))<>convertGetal(iId) then qs_gallery_start=0
dim fileKey, imagePath, sCopyCustomLink
for each fileKey in filesDict
imagePath=C_VIRT_DIR & Application("QS_CMS_userfiles") & sPath & "/"
imagePath=replace(imagePath,"//","/",1,-1,1)
set file=fso.GetFile (server.mappath(replace(imagePath,"/","\",1,-1,1))& "\" & filesDict(fileKey))
'check if image
if allowedFileTypesforThumbing.exists(lcase(GetFileExtension(file.name))) then
if counter>qs_gallery_start*iBrowseBy and counter<=(qs_gallery_start*iBrowseBy)+iBrowseBy then
'start new row if needed
if (colCounter MOD iPicsInRow)=0 and counter<>1 then 
build=build &"</tr><tr>"
colCounter=0
end if
colCounter=colCounter+1
if isLeeg(sCustomLink) then
select case lcase(GetFileExtension(file.name))
case "jpg","jpeg"
picLink="<a href="""& C_DIRECTORY_QUICKERSITE &"/showThumb.aspx?FSR=0&amp;"
picLink=picLink & "img=" & server.URLEncode (sFolder & "/" & file.name) &"&amp;maxSize="&iFullImageSize &"""  "
case else
picLink="<a href="""& sFolder & "/" & file.name & """ class=""QSPPIMG"""
end select
picLink=picLink & " rel=""lightbox" & replace(sCode,"'","\'",1,-1,1) & """ title="""
'show picturename if needed
if bShowFileName then
picLink=picLink & quotrep(left(file.name,len(file.name)-4)) '& " - " 
end if
picLink=picLink& """"
else
sCopyCustomLink=sCustomLink
sCopyCustomLink=replace(sCustomLink,"[FILENAME]",quotrep(file.name),1,-1,1)
sCopyCustomLink=replace(sCopyCustomLink,"[COUNTER]",counter,1,-1,1)
sCopyCustomLink=replace(sCopyCustomLink,"[PAGEID]",encrypt(selectedPage.iId),1,-1,1)
picLink="<a " & treatConstants(sCopyCustomLink,true)
end if
picLink=picLink &">" & vbcrlf & "<img style=" & """" & sStyleImage & """" 
if bShowFileName then
picLink=picLink & " alt=" & """" & quotrep(left(file.name,len(file.name)-4)) & """"
else
picLink=picLink & " alt='' "
end if
select case lcase(GetFileExtension(file.name))
case "jpg","jpeg"
picLink=picLink &" src="""& C_DIRECTORY_QUICKERSITE &"/showThumb.aspx?"
if bFSR then
picLink=picLink&"FSR=1&amp;"
else
picLink=picLink&"FSR=0&amp;"
end if
select case convertGetal(iSpecialEffect)
case QS_gallery_SE_BW
picLink=picLink&"SE=1&amp;"
case QS_gallery_SE_GS
picLink=picLink&"SE=2&amp;"
case QS_gallery_SE_SE
picLink=picLink&"SE=3&amp;"
case else
picLink=picLink&"SE=0&amp;"
end select
picLink=picLink &"maxSize=" & iThumbSize & "&amp;img="& server.URLEncode (sFolder & "/" & file.name) & """ /></a>"
case else
picLink=picLink &" src="""& sFolder & "/" & file.name & """"
if bFSR then
picLink=picLink &" style=""width:" & iThumbSize & "px;height:" & iThumbSize & "px"" width='" & iThumbSize & "' height='" & iThumbSize & "' /></a>"
else
picLink=picLink &" style=""width:" & iThumbSize & """ width='" & iThumbSize & "' /></a>"
end if
end select
if not isLeeg(sFullImage) then
picLink=picLink&"<div style=""width:100%;text-align:center;margin:3px;font-size:smaller"">" & treatConstants(sFullImage,true) & "</div>"
end if
'build cell
build=build & "<td valign=""top"" style=""" & sStyleTableCell & """ align=""center"">"
'add shadow if needed
'if bShadowThumb then
'	build=build&wrapInShadowCSS(piclink)
'else
build=build&piclink
'end if
if bShowFileName then
build=build & "<div style=""width:100%;text-align:center"">" & quotrep(left(file.name,len(file.name)-4)) & "</div>"
end if 
build=build&"</td>"
else
if isLeeg(sCustomLink) then
'foto toevoegen?
select case lcase(GetFileExtension(file.name))
case "jpg","jpeg"
additionalImages="<a href="""& C_DIRECTORY_QUICKERSITE &"/showThumb.aspx?FSR=0&amp;"
additionalImages=additionalImages&"img=" & server.URLEncode (sFolder & "/" & file.name) &"&amp;maxSize="&iFullImageSize &""" rel=""lightbox" & replace(sCode,"'","\'",1,-1,1) & """ title="""
case else
additionalImages="<a href="""& sFolder & "/" & file.name & """"
additionalImages=additionalImages&" rel=""lightbox" & replace(sCode,"'","\'",1,-1,1) & """ title="""
end select
if bShowFileName then
additionalImages=additionalImages & quotrep(left(file.name,len(file.name)-4)) ' & " - " 
end if
additionalImages=additionalImages & """></a>"
build=build & "<td style=""display:none"">" & additionalImages & "</td>"
end if
end if
select case lcase(GetFileExtension(file.name))
case "jpg","jpeg"


dim i,captionValue,urlValue,oinwValue,fHeight,fWidth
oinwValue=""
captionValue=""
urlValue=""

	if sType=QS_gallery_NS then		

		includeNS=true

		'captions
		dim sNSImgLinksARR,lineARR
		sNSImgLinksARR=split(sNSImgLinks,vbcrlf)

		for i=lbound(sNSImgLinksARR) to ubound(sNSImgLinksARR)
			if left(sNSImgLinksARR(i),len(removeDots(file.name)))=removeDots(file.name) then
				lineArr=split(sNSImgLinksARR(i),vbtab)
				captionValue=lineArr(1)
				urlValue=lineArr(2)
				if convertBool(lineArr(3)) then oinwValue=" target=""_blank"" "
			end if
		next
		
	end if

slideshow=slideshow & "fadeimages"&iId&genPW & "|R|R|R|" & counter-1 & "]=|R|R|R|""" & C_DIRECTORY_QUICKERSITE &"/showThumb.aspx?"

if not isLeeg(urlValue) then
	slideshowC=slideshowC & "<a " & oinwValue & " href=""" & urlValue & """>"
end if

slideshowC=slideshowC & "<img alt="""" border=""0"" style=""margin:0 !important;padding:0 !important;border-style:none"" src=""" & C_DIRECTORY_QUICKERSITE & "/showThumb.aspx?"

if convertGetal(fHeight)<>0 then
	slideshowC=slideshowC & "FH=" & fHeight & "&amp;"
end if
if convertGetal(fWidth)<>0 then
	slideshowC=slideshowC & "FW=" & fWidth & "&amp;"
end if



select case convertGetal(iSpecialEffect)
case QS_gallery_SE_BW
slideshow=slideshow&"SE=1&amp;"
slideshowC=slideshowC&"SE=1&amp;"
case QS_gallery_SE_GS
slideshow=slideshow&"SE=2&amp;"
slideshowC=slideshowC&"SE=2&amp;"
case QS_gallery_SE_SE
slideshow=slideshow&"SE=3&amp;"
slideshowC=slideshowC&"SE=3&amp;"
case else
slideshow=slideshow&"SE=0&amp;"
slideshowC=slideshowC&"SE=0&amp;"
end select
slideshow=slideshow & "maxSize=" & iThumbSize & "&amp;FSR="
if convertGetal(fWidth)=>convertGetal(fHeight) then
slideshowC=slideshowC & "maxSize=" & fWidth & "&amp;FSR="
else
slideshowC=slideshowC & "maxSize=" & fHeight & "&amp;FSR="
end if
if convertBool(bFSR) then
slideshow=slideshow&"1"
slideshowC=slideshowC&"1"
else
slideshow=slideshow&"0"
slideshowC=slideshowC&"0"
end if
slideshow=slideshow & "&amp;img="& server.URLEncode (sFolder & "/" & file.name) & """, """ & sCustomLink & """, """

if sType<>QS_gallery_NS then
slideshowC=slideshowC & "&amp;img="& server.URLEncode (sFolder & "/" & file.name) & """ height=""" & fHeight & """ width=""" & fWidth & """"
else
slideshowC=slideshowC & "&amp;img="& server.URLEncode (sFolder & "/" & file.name) & """ "
end if

if not isLeeg(captionValue) then
	slideshowC=slideshowC & " title="""   & sanitize(captionValue) &  """ "
end if

slideshowC=slideshowC & "/>" & vbcrlf

if not isLeeg(urlValue) then
slideshowC=slideshowC & "</a>" & vbcrlf
end if

if bOpenInNewWindow then
slideshow=slideshow & "_blank"
end if
slideshow=slideshow & """];"
case else
slideshow=slideshow & "fadeimages"&iId&genPW & "|R|R|R|" & counter-1 & "]=|R|R|R|""" & sFolder & "/" & file.name &""""
slideshowC=slideshowC & "<img alt="""" border=""0"" style=""width:" & fWidth & "px;height:" & fHeight & "px;border-style:none"" src=""" & sFolder & "/" & file.name & """"
slideshow=slideshow & ", """ & sCustomLink & """, """
slideshowC=slideshowC & " style=""height:" & fHeight & "px;width:" & fWidth & "px"" height=""" & fHeight & """ width=""" & fWidth & """ />" & vbcrlf
if bOpenInNewWindow then
slideshow=slideshow & "_blank"
end if
slideshow=slideshow & """];"
end select
counter=counter+1
end if
set file=nothing
next
if counter-1>convertGetal(qs_gallery_start*iBrowseBy)+iBrowseBy then
showNextLink=true
end if
if lcase(convertStr(sStyleTable))="width:100%;border-style:none" or isLeeg(sStyleTable) then
sStyleTable="margin:0 auto;border-style:none"
end if
build="<a name='QS_gallery" & iId &"'></a><table style=" & """" & sStyleTable & """" & ">[NAVBUILD]<tr>" & build & "</tr>[NAVBUILD]</table>"
dim navBuild
if showNextLink or qs_gallery_start>0 then
navBuild="<tr><td align=""center"" colspan=""" & iPicsInRow & """>"
if qs_gallery_start>0 then
navBuild=navBuild&"<a class=""art-button"" href='default.asp?iId=" & encrypt(selectedPage.iId) & "&amp;iGalleryID=" & encrypt(iId) & "&amp;qs_gallery_start=" & qs_gallery_start-1 & "#QS_gallery" & iId & "'>" & treatConstants(sPreviousLink,true) & "</a>&nbsp;&nbsp;"
end if
if showNextLink then
navBuild=navBuild&"&nbsp;&nbsp;<a class=""art-button"" href='default.asp?iId=" & encrypt(selectedPage.iId) & "&amp;iGalleryID=" & encrypt(iId) & "&amp;qs_gallery_start=" & qs_gallery_start+1 & "#QS_gallery" & iId & "'>" & treatConstants(sNextLink,true) & "</a>"
end if
navBuild=navBuild& "</td></tr>"
end if

if isLeeg(p_sPreviousLink) and isLeeg(p_sNextlink) then navBuild=""
build=Jscript & vbcrlf & build
build=replace(build,"[NAVBUILD]",navBuild,1,-1,1) '& additionalImages
build=replace(build,"<tr></tr>","",1,-1,1)

select case sType

case QS_gallery_NS

	dim generNSPW
	generNSPW=generatepassword
	build=vbcrlf & vbcrlf & "<div class=""slider-wrapper theme-"  & sNSCss &""">"
	build=build & "<div id=""slider" & generNSPW & """ class=""nivoSlider"">"
	
	slideshowC=replace(slideshowC,"height="""" width=""""","",1,-1,1)
	slideshowC=replace(slideshowC,"maxSize=","maxSize=1920",1,-1,1)
	slideshowC=replace(slideshowC,"  "," ",1,-1,1)
	
	build=build & slideshowC
	build=build & "</div>"
	build=build & "</div>"	
	
	dim nivoAdd
	nivoAdd=""
	
	nivoAdd=nivoAdd & "<script type=""text/javascript"">"
	nivoAdd=nivoAdd & "$(window).load(function() {"
	nivoAdd=nivoAdd & "$('#slider" & generNSPW & "').nivoSlider({controlNav:"& lcase(convertTF(bNSControlNav)) & ",directionNav:"& lcase(convertTF(bNSdirectionNav)) & ",effect:'" & sCycleEffect & "',pauseTime:" & iSlideShowTimerQS*1000 & "});"
	nivoAdd=nivoAdd & "});"
	nivoAdd=nivoAdd & "</script>"
	
	headerDictionary.add generatePassword,nivoAdd
	
case QS_gallery_SC
	dim generSSPW
	generSSPW=generatepassword
	build=vbcrlf & vbcrlf &"<script type=""text/javascript"">" & vbcrlf
	build=build& "$(document).ready(function() {" & vbcrlf
	build=build& "$('." & generSSPW & "').cycle({" & vbcrlf
	build=build& "fx: '" & sCycleEffect &"', " & vbcrlf
	build=build& "timeout: " & iSlideShowTimerQS*1000 & "," & vbcrlf
	build=build&"pause:1" & vbcrlf
	build=build& "});"& vbcrlf
	build=build& "});"& vbcrlf
	build=build&"</script>"& vbcrlf & vbcrlf
	build=build& "<div [ONCLICK]" & generSSPW & " class=""" & generSSPW &  """ style=""[STYLE]height: " & fHeight & "px; width: " & fWidth & "px; margin: auto"">" & vbcrlf & slideshowC & vbcrlf &  "</div>"
	
	if not isLeeg(sCustomLink) then
		if bOpenInNewWindow then
			build=replace(build,"[ONCLICK]" & generSSPW,"onclick=""javascript:window.open('" & sCustomLink & "');"" ",1,-1,1)
		else
			build=replace(build,"[ONCLICK]" & generSSPW,"onclick=""javascript:location.assign('" & sCustomLink & "');"" ",1,-1,1)
		end if
			build=replace(build,"[STYLE]","cursor:hand; cursor:pointer;",1,-1,1)
		else
			build=replace(build,"[ONCLICK]" & generSSPW,"",1,-1,1)
			build=replace(build,"[STYLE]","",1,-1,1)
	end if
	
case QS_gallery_SS
if isLeeg(sBorderColor) then
sBorderColor="#AAAAAA"
end if
build=Jscript&"<script type=""text/javascript"">"
build=build& "var bColor='"&sBorderColor&"';"
build=build& "var fadeimages"&iId&genPW&"=new Array();"
build=build& slideshow
build=build& "new fadeshow(fadeimages"&iId&genPW&", " & fWidth & ", " & fHeight & ", " & sBorder & ", " & iSlideShowTimerQS*1000 & ", "
if bRandom then
build=build& "1"
else
build=build& "0"
end if 
build=build& ")"
build=build& "</script>"& vbcrlf
end select
else
build=Jscript
end if
'Cleanup
set filesDict=nothing
set folder=nothing
end if
set fso=nothing
On Error Goto 0
end if
dim sAddHeaderQS
if instr(build,"lightbox")<>0 then
sAddHeaderQS="<script type=""text/javascript"">" & vbcrlf
sAddHeaderQS=sAddHeaderQS & "$(document).ready(function(){" & vbcrlf
sAddHeaderQS=sAddHeaderQS & "$(""a[rel='lightbox" & replace(sCode,"'","\'",1,-1,1) & "']"").colorbox({slideshowSpeed:" & iSlideShowTimerQS*1000 & ", initialWidth:150, initialHeight:100, photo:true, slideshow:true, transition:""elastic"""
if bAutoStartSS then
sAddHeaderQS=sAddHeaderQS & ", slideshowAuto:true"
else
sAddHeaderQS=sAddHeaderQS & ", slideshowAuto:false"
end if
sAddHeaderQS=sAddHeaderQS & "});" & vbcrlf
sAddHeaderQS=sAddHeaderQS & "});" & vbcrlf
sAddHeaderQS=sAddHeaderQS & "</script>" & vbcrlf
headerDictionary.Add generatePassword, sAddHeaderQS
end if
end function
public property get Jscript
Jscript="<script type=""text/javascript"">" & vbcrlf & "var slideShowTimerQS=" & iSlideShowTimerQSDefault & ";"
if convertBool(bAutoStartSS) then
Jscript=Jscript & "var bAutoStartSS=true;" & vbcrlf
else
Jscript=Jscript & "var bAutoStartSS=false;" & vbcrlf
end if
Jscript=Jscript & "</script>"
end property
public function copy()
if isNumeriek(iId) then
iId=null
sName=l("copyof") & " " & sName
sCode=GeneratePassWord()
save()
end if
end function
end class%>
