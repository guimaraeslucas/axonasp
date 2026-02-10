<%

function filterImages(v)

	while instr(v,"data:image/png;base64,")<>0

		dim startPos : startPos=instr(v,"data:image/png;base64,")		
		
		dim endPos : endPos=instr(startPos+1,v,"""")
		startPos=startPos+22	
		
		dim binText : binText=mid(v,startPos,endPos-startPos)
		
		dim oXML : Set oXML = CreateObject("Msxml2.DOMDocument")
		dim oNode : Set oNode = oXML.CreateElement("base64")
		oNode.dataType = "bin.base64"
		oNode.text = binText

		dim BinaryStream : Set BinaryStream = CreateObject("ADODB.Stream")
		BinaryStream.Type = 1 'adTypeBinary
		BinaryStream.Open
		BinaryStream.Write oNode.nodeTypedValue
			
		dim filename : filename=GeneratePassWord & GeneratePassWord
		BinaryStream.SaveToFile server.mappath(C_VIRT_DIR & Application("QS_CMS_userfiles")) & "\" & filename & ".png"

		v=replace(v,"data:image/png;base64," & binText,C_VIRT_DIR & Application("QS_CMS_userfiles") & filename & ".png",1,-1,1)			
	
	wend
	
	filterImages=v

end function 

function removeCRB(value)

	'remove Code Render Blocks

	value=replace(value,"<" & "%","",1,-1,1)
	removeCRB=replace(value,"%" & ">","",1,-1,1)	

end function	

function QS_iFrameYT(input)
	'this function converts a youtube url to an iframed youtube movie
	if not isLeeg(input) then
		on error resume next
		dim myRegExp,replacement
		Set myRegExp = New RegExp
		myRegExp.IgnoreCase = True
		myRegExp.Global = True
		myRegExp.Pattern = "(?:https?:\/\/)?(?:www\.)?(?:(?:(?:youtube.com\/watch\?[^?]*v=|youtu.be\/)([\w\-]+))(?:[^\s?]+)?)"
		replacement = "<div style='margin:10px 0px 10px 0px;clear:both'><iframe title='YouTube video player' width='480' height='390' src='http://www.youtube.com/embed/$1' frameborder='0' allowfullscreen='1'></iframe></div>"
		iFrameYT=myRegExp.Replace(input, replacement)
	end if
	on error goto 0
end function

function compressHTML(value)

	compressHTML=value
	
	while instr(compressHTML,vbtab)<>0
		compressHTML=replace(compressHTML,vbtab," ",1,-1,1)
	wend
	
	while instr(compressHTML,vbcrlf)<>0
		compressHTML=replace(compressHTML,vbcrlf," ",1,-1,1)
	wend
	
	while instr(compressHTML,vbCr)<>0
		compressHTML=replace(compressHTML,vbCr," ",1,-1,1)
	wend
	
	while instr(compressHTML,vbLf)<>0
		compressHTML=replace(compressHTML,vbLf," ",1,-1,1)
	wend
	
	while instr(compressHTML,"  ")<>0
		compressHTML=replace(compressHTML,"  "," ",1,-1,1)
	wend

end function


Function isLeeg(byval value)
isLeeg=false
if isNull(value) then
isLeeg=true
else
if isEmpty(value) or trim(value)="" then isLeeg=true
end if
End Function
Function convertGetal(value)
if not isnull(value) then
if IsNumeric(value) then
if trim(value)<>"" then
on error resume next
convertGetal=cdbl(value)
dumperror "convertGetal:"&value,err
err.clear
on error goto 0
else
convertGetal=0
end if
else
convertGetal=0
end if
else
convertGetal=0
end if
End Function
'clng
Function convertLng(value)
if not isnull(value) then
if IsNumeric(value) then
if trim(value)<>"" then
on error resume next
convertLng=cdbl(value)
dumperror "convertLng:"&value,err
err.clear
on error goto 0
else
convertLng=null
end if
else
convertLng=null
end if
else
convertLng=null
end if
End Function
'cstr
Function convertStr(value)
on error resume next
if not isnull(value) then
convertStr=cstr(value)
else
convertStr=""
end if
if err.number<>0 then
convertStr=value
end if
on error goto 0
End Function
'clng
Function convertBool(value)
On Error Resume Next
if isLeeg(value) then
convertBool=false
exit function
end if
if convertGetal(value)=1 then
convertBool=true
exit function
end if
if cstr(value)="0" then
convertBool=false
exit function
end if
if lcase(cstr(value))="true" then
convertBool=true
exit function
end if
if lcase(cstr(value))="false" then
convertBool=false
exit function
end if
if value=true or cBool(value) then
convertBool=true
exit function
end if
convertBool=false
On Error Goto 0
End Function
Function convertTF(value)
if convertBool(value) then
convertTF="true"
else
convertTF="false"
end if
End Function
Function convertTFS(value)
if convertBool(value) then
convertTFS="background-color:Green;color:#FFFFFF"
else
convertTFS="background-color:Red;color:#FFFFFF"
end if
End Function
function isNumeriek(value)
if isLeeg(value) then
isNumeriek=false
else
isNumeriek=isNumeric(value)
end if
end function
function leestekens
dim arrLt(26)
arrLt(0)="\"
arrLt(1)="'"
arrLt(2)="''"
arrLt(3)="§"
arrLt(4)="`"
arrLt(5)=":"
arrLt(6)=";"
arrLt(7)=","
arrLt(8)="¤"
arrLt(9)="("
arrLt(10)=")"
arrLt(11)="="
arrLt(12)=""""
arrLt(13)="#"
arrLt(14)="%"
arrLt(15)="&"
arrLt(16)="*"
'arrLt(17)="+"
arrLt(18)="<"
arrLt(19)=">"
arrLt(20)="?"
arrLt(21)="["
arrLt(22)="]"
arrLt(23)="{"
arrLt(24)="}"
arrLt(25)="|"
leestekens=arrLt
end function

function cleanUp(str)
	if isLeeg(str) then
		cleanUp=""
	else
		cleanUp=replace(str,"'","''",1,-1,1)
	end if
end function

function cleanUpStr(str)
On Error Resume Next
cleanUpStr=str
dim lt
lt=leestekens
dim i
for i=lbound(lt) to ubound(lt)-1
cleanUpStr=replace(cleanUpStr,convertStr(lt(i)),"",1,-1,1)
next
On Error Goto 0
end function



Function RemoveHTML( strText )
    Dim TAGLIST
    TAGLIST = ";!--;!DOCTYPE;A;ACRONYM;ADDRESS;APPLET;AREA;B;BASE;BASEFONT;" &_
              "BGSOUND;BIG;BLOCKQUOTE;BODY;BR;BUTTON;CAPTION;CENTER;CITE;CODE;" &_
              "COL;COLGROUP;COMMENT;DD;DEL;DFN;DIR;DIV;DL;DT;EM;EMBED;FIELDSET;" &_
              "FONT;FORM;FRAME;FRAMESET;HEAD;H1;H2;H3;H4;H5;H6;HR;HTML;I;IFRAME;IMG;" &_
              "INPUT;INS;ISINDEX;KBD;LABEL;LAYER;LAGEND;LI;LINK;LISTING;MAP;MARQUEE;" &_
              "MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;OL;OPTION;P;PARAM;PLAINTEXT;" &_
              "PRE;Q;S;SAMP;SCRIPT;SELECT;SMALL;SPAN;STRIKE;STRONG;STYLE;SUB;SUP;" &_
              "TABLE;TBODY;TD;TEXTAREA;TFOOT;TH;THEAD;TITLE;TR;TT;U;UL;VAR;WBR;XMP;"
    dim BLOCKTAGLIST
     BLOCKTAGLIST = ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"
    
    Dim nPos1
    Dim nPos2
    Dim nPos3
    Dim strResult
    Dim strTagName
    Dim bRemove
    Dim bSearchForBlock
    
    nPos1 = InStr(strText, "<")
    Do While nPos1 > 0
        nPos2 = InStr(nPos1 + 1, strText, ">")
        If nPos2 > 0 Then
            strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
    strTagName = Replace(Replace(strTagName, vbCr, " ",1,-1,1), vbLf, " ",1,-1,1)
            nPos3 = InStr(strTagName, " ")
            If nPos3 > 0 Then
                strTagName = Left(strTagName, nPos3 - 1)
            End If
            
            If Left(strTagName, 1) = "/" Then
                strTagName = Mid(strTagName, 2)
                bSearchForBlock = False
            Else
                bSearchForBlock = True
            End If
            
            If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                bRemove = True
                If bSearchForBlock Then
                    If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                        nPos2 = Len(strText)
                        nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
                        If nPos3 > 0 Then
                            nPos3 = InStr(nPos3 + 1, strText, ">")
                        End If
                        
                        If nPos3 > 0 Then
                            nPos2 = nPos3
                        End If
                    End If
                End If
            Else
                bRemove = False
            End If
            
            If bRemove Then
                strResult = strResult & Left(strText, nPos1 - 1)
                strText = Mid(strText, nPos2 + 1)
            Else
                strResult = strResult & Left(strText, nPos1)
                strText = Mid(strText, nPos1 + 1)
            End If
        Else
            strResult = strResult & strText
            strText = ""
        End If
        
        nPos1 = InStr(strText, "<")
    Loop
    
    strResult = strResult & strText
    
    convertTo_iso_8859_1(strResult) 
   
    strResult=Replace(strResult, "&euro;", "€",1,-1,1)
    strResult=Replace(strResult, "<o:p>", "",1,-1,1)
    strResult=Replace(strResult, "</o:p>", "",1,-1,1)       
    strResult=replace(strResult, vbcrlf," ",1,-1,1)  
        
    RemoveHTML = remove2whites(strResult)
End Function
Function RemoveJS( strText )
dim TAGLIST
   TAGLIST = ";!--;!DOCTYPE;ACRONYM;APPLET;AREA;BASE;BASEFONT;" &_
              "BGSOUND;BODY;BUTTON;CODE;" &_
              "COMMENT;DFN;DIR;EMBED;" &_
              "FORM;FRAME;FRAMESET;HEAD;HTML;IFRAME;" &_
              "INPUT;ISINDEX;KBD;LABEL;LAYER;LAGEND;LINK;LISTING;MAP;" &_
              "MENU;META;NOBR;NOFRAMES;NOSCRIPT;OBJECT;PARAM;PLAINTEXT;" &_
              "Q;S;SAMP;SCRIPT;STYLE;" &_
              "TITLE;VAR;WBR;XMP;"
              
    dim BLOCKTAGLIST 
    BLOCKTAGLIST= ";APPLET;EMBED;FRAMESET;HEAD;NOFRAMES;NOSCRIPT;OBJECT;SCRIPT;STYLE;"
    
    Dim nPos1
    Dim nPos2
    Dim nPos3
    Dim strResult
    Dim strTagName
    Dim bRemove
    Dim bSearchForBlock
    
    nPos1 = InStr(strText, "<")
    Do While nPos1 > 0
        nPos2 = InStr(nPos1 + 1, strText, ">")
        If nPos2 > 0 Then
            strTagName = Mid(strText, nPos1 + 1, nPos2 - nPos1 - 1)
    strTagName = Replace(Replace(strTagName, vbCr, " "), vbLf, " ",1,-1,1)
            nPos3 = InStr(strTagName, " ")
            If nPos3 > 0 Then
                strTagName = Left(strTagName, nPos3 - 1)
            End If
            
            If Left(strTagName, 1) = "/" Then
                strTagName = Mid(strTagName, 2)
                bSearchForBlock = False
            Else
                bSearchForBlock = True
            End If
            
            If InStr(1, TAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                bRemove = True
                If bSearchForBlock Then
                    If InStr(1, BLOCKTAGLIST, ";" & strTagName & ";", vbTextCompare) > 0 Then
                        nPos2 = Len(strText)
                        nPos3 = InStr(nPos1 + 1, strText, "</" & strTagName, vbTextCompare)
                        If nPos3 > 0 Then
                            nPos3 = InStr(nPos3 + 1, strText, ">")
                        End If
                        
                        If nPos3 > 0 Then
                            nPos2 = nPos3
                        End If
                    End If
                End If
            Else
                bRemove = False
            End If
            
            If bRemove Then
                strResult = strResult & Left(strText, nPos1 - 1)
                strText = Mid(strText, nPos2 + 1)
            Else
                strResult = strResult & Left(strText, nPos1)
                strText = Mid(strText, nPos1 + 1)
            End If
        Else
            strResult = strResult & strText
            strText = ""
        End If
        
        nPos1 = InStr(strText, "<")
    Loop
    
    strResult = strResult & strText
    
    convertTo_iso_8859_1(strResult) 
   
    strResult=Replace(strResult, "&euro;", "€",1,-1,1)
    strResult=Replace(strResult, "<o:p>", "",1,-1,1)
    strResult=Replace(strResult, "</o:p>", "",1,-1,1)       
    strResult=replace(strResult, vbcrlf," ",1,-1,1)  
        
    removeJS = remove2whites(strResult)
End Function
function filterJS(s)
dim arrJS, jsKey
arrJS=array("javascript:", "vbscript:", "onabort", "onactivate", "onafterprint", "onafterupdate", "onbeforeactivate", "onbeforecopy", "onbeforecut", "onbeforedeactivate", "onbeforeeditfocus", "onbeforepaste", "onbeforeprint", "onbeforeunload", "onbeforeupdate", "onblur", "onbounce", "oncellchange", "onchange", "onclick", "oncontextmenu", "oncontrolselect", "oncopy", "oncut", "ondataavailable", "ondatasetchanged", "ondatasetcomplete", "ondblclick", "ondeactivate", "ondrag", "ondragend", "ondragenter", "ondragleave", "ondragover", "ondragstart", "ondrop", "onerror", "onerrorupdate", "onfilterchange", "onfinish", "onfocus", "onfocusin", "onfocusout", "onhelp", "onkeydown", "onkeypress", "onkeyup", "onlayoutcomplete", "onload", "onlosecapture", "onmousedown", "onmouseenter", "onmouseleave", "onmousemove", "onmouseout", "onmouseover", "onmouseup", "onmousewheel", "onmove", "onmoveend", "onmovestart", "onpaste", "onpropertychange", "onreadystatechange", "onreset", "onresize", "onresizeend", "onresizestart", "onrowenter", "onrowexit", "onrowsdelete", "onrowsinserted", "onscroll", "onselect", "onselectionchange", "onselectstart", "onstart", "onstop", "onsubmit", "onunload")
for jsKey=lbound(arrJS) to ubound(arrJS)
s=replace(s,arrJS(jsKey),"[" & arrJS(jsKey) & "]",1,-1,1)
next
filterJS=s
end function
function remove2whites(byref svalue)
if instr(svalue,"  ")<>0 then
svalue=replace(svalue,"  "," ",1,-1,1)
svalue=remove2whites(svalue)
end if
remove2whites=svalue
end function
dim p_getHomePage
p_getHomePage=null
dim p_getIntranetHomePage
p_getIntranetHomePage=null
Function getHomePage()
On Error Resume Next
if isNull(p_getHomePage) then
if convertGetal(Application("iHomePageID"))=0 then
dim rs
set rs=db.execute("select iId from tblPage where bIntranet="&getSQLBoolean(false)&" and bHomepage="&getSQLBoolean(true)&" and "& sqlCustId)
Application("iHomePageID")=convertGetal(rs(0))
set rs=nothing
end if
set p_getHomePage=new cls_page
p_getHomePage.pick(Application("iHomePageID"))
end if
set getHomePage=p_getHomePage
On Error Goto 0
end Function
Function getHomePageNOCID(custID)
'On Error Resume Next
dim rs
set rs=db.execute("select iId from tblPage where bIntranet="&getSQLBoolean(false)&" and bHomepage="&getSQLBoolean(true)&" and iCustomerID="& custID)
set getHomePageNOCID=new cls_page
getHomePageNOCID.overruleCID=custID
getHomePageNOCID.pick(convertGetal(rs(0)))
set rs=nothing
'On Error Goto 0
end Function
Function getIntranetHomePage()
if isNull(p_getIntranetHomePage) then
if convertGetal(Application("iIntranetHomePageID"))=0 then
dim rs
set rs=db.execute("select iId from tblPage where bIntranet="&getSQLBoolean(true)&" and bHomepage="&getSQLBoolean(true)&" and "& sqlCustId)
Application("iIntranetHomePageID")=convertGetal(rs(0))
set rs=nothing
end if
set p_getIntranetHomePage=new cls_page
p_getIntranetHomePage.pick(Application("iIntranetHomePageID"))
end if
set getIntranetHomePage=p_getIntranetHomePage
end Function
function cId
cId=convertGetal(Application("QS_CMS_iCustomerID"))
end function
function sqlCustId
sqlCustId=" tblPage.iCustomerID="& Application("QS_CMS_iCustomerID") & " "
end function
function sqlCustIdCF
sqlCustIdCF=" tblContactField.iCustomerID="& Application("QS_CMS_iCustomerID") & " "
end function
'error
Function ErrorReport(value,byref error)
if not isLeeg(C_ADMINEMAIL) then
dim theMail
set theMail=new cls_mail_message
theMail.receiver=C_ADMINEMAIL
theMail.subject="Error op "& customer.siteName
theMail.body=value & "<hr /><p>Visitor details:</p>" & getVisitorDetails
theMail.send
set theMail=nothing
end if
End Function
function dumperror(value,byref error)
if error.number<>0 then
 
dumperror="<p><b>Page Error:</b><br />"
dumperror=dumperror& "Error Number: " & error.number & "<BR />"
dumperror=dumperror& "Error Description: " & error.Description & "<BR />"
dumperror=dumperror& "Error Source: " & error.Source & "<BR />"
dumperror=dumperror& "Session Variables: " & "<BR />"
dim si
for each si in session.Contents
dumperror=dumperror & si & ": "& session.Contents(si) & "<br />"
next
dumperror=dumperror&"</p>"
ErrorReport "<p>"&value&"</p>"&dumperror,error
error.Clear
'sResponse.Flush 
'Response.End 
end if
end function
function dumpConnError(value,byref errorColl)
'exit function
if errorColl.count>0 then
dim e
For each e in errorColl
dumpConnError=dumpConnError&"Error No: " & e.Number &"<BR />"
dumpConnError=dumpConnError&"Description: " & e.Description &"<BR />"
dumpConnError=dumpConnError&"Source: " & e.Source & "<BR />"
dumpConnError=dumpConnError&"SQLState: " & e.SQLState & "<BR />"
dumpConnError=dumpConnError&"NativeError: " & e.NativeError & "<br />"
ErrorReport "<p>"&value&"</p>"&dumpConnError,err
Next
Response.Write "<p>"&value&"</p>"&dumpConnError
Response.Flush 
Response.End 
end if
end function
function convertEuroDate(theDate)
if not isLeeg(theDate) then
select case customer.sDatumFormat
case QS_dateFM_US
convertEuroDate=convert2(month(theDate)) & "/" & convert2(day(theDate)) & "/" & year(theDate)
case else
convertEuroDate=convert2(day(theDate)) & "/" & convert2(month(theDate)) & "/" & year(theDate)
end select
else
convertEuroDate=""
end if
end function
Function formatTimeStamp(timestamp)
Dim date_part, time_part
date_part = convertEuroDate(timestamp)
time_part = right("0" & hour(timestamp), 2) & ":" & right("0" & minute(timestamp), 2) & ":" & right("0" & second(timestamp), 2)
formatTimeStamp = date_part & " " & time_part
End Function
function convertDateFromPicker(theDate)
if not isLeeg(theDate) then
dim arrDate
arrDate=split(theDate,"/")
select case customer.sDatumFormat
case QS_dateFM_US
convertDateFromPicker=dateserial(arrDate(2),arrDate(0),arrDate(1))
case else
convertDateFromPicker=dateserial(arrDate(2),arrDate(1),arrDate(0))
end select
else
convertDateFromPicker=null
end if
end function
public function convertEuroDateTime(theDate)
if not isLeeg(theDate) then
select case customer.sDatumFormat
case QS_dateFM_US
convertEuroDateTime=month(theDate) & "/" & day(theDate) & "/" & year(theDate) & " " & convert2(hour(theDate)) & ":"& convert2(minute(theDate))	& ":"& convert2(second(theDate))
case else
convertEuroDateTime=day(theDate) & "/" & month(theDate) & "/" & year(theDate) & " " & convert2(hour(theDate)) & ":"& convert2(minute(theDate))	& ":"& convert2(second(theDate))
end select
else
convertEuroDateTime=""
end if
end function
function convertDateToPicker(theDate)
if not isLeeg(theDate) then
select case customer.sDatumFormat
case QS_dateFM_US
convertDateToPicker=mid(theDate,5,2) & "/" & right(theDate,2) & "/" & left(theDate,4)
case else
convertDateToPicker=right(theDate,2) & "/" & mid(theDate,5,2) & "/" & left(theDate,4)
end select
else
convertDateToPicker=null
end if
end function
function convertCalcDate(theDate)
if not isLeeg(theDate) then
dim arrDate
arrDate=split(theDate,"/")
select case customer.sDatumFormat
case QS_dateFM_US
convertCalcDate=arrDate(2)&arrDate(0)&arrDate(1)
case else
convertCalcDate=arrDate(2)&arrDate(1)&arrDate(0)
end select
else
convertCalcDate=null
end if
end function
function convertCalcDateTime(theDate)
convertCalcDateTime=year(theDate) & convert2(month(theDate))  & convert2(day(theDate)) & convert2(hour(theDate)) & convert2(minute(theDate)) & convert2(second(theDate))
'response.write convertCalcDateTime &"<br />"
end function
function roomDate(theDate,mode)
if isLeeg(theDate) then
select case mode
case "down"
roomDate=dateserial(1900,1,1)
case "up"
roomDate=dateAdd("yyyy",10,date())
end select
else
roomDate=theDate
end if
end function
dim INNOVALOADED
INNOVALOADED=false
function dumpFCKInstance(sValue,sToolbar,sFieldName)
'sValue=filterImages(sValue)
dumpFCKInstance=""
if isLeeg(sValue) then sValue=""
select case QS_EDITOR
case 4 'ckeditor
dim editor,bAllowUpload
bAllowUpload=true
set editor = New CKEditor
editor.returnOutput = true
editor.basePath = C_DIRECTORY_QUICKERSITE & "/" & C_CKEDITOR & "/"
editor.instanceConfig("skin") = "kama"

'editor.config("forcePasteAsPlainText")=true
'editor.config("pasteFromWordRemoveFontStyles")=false
'editor.config("pasteFromWordRemoveStyles")=false
editor.config("entities")=false
select case sToolbar
case "siteBuilderBanner"
editor.config("width") = 400
editor.config("height") = 500
case "siteBuilder","siteBuilderMail","template"
editor.config("width") = 700
editor.config("height") = 400
case "siteBuilderRichText","siteBuilderMailNoSource","siteBuilderRichTextSmall","siteBuilderMailNoSourceUpload","siteBuilderMailSource"
if sToolbar<>"siteBuilderMailNoSourceUpload" and sToolbar<>"siteBuilderMailSource" then 
bAllowUpload=false
end if
if sToolbar="siteBuilderMailNoSource" or sToolbar="siteBuilderMailNoSourceUpload" then
if bThemeEmbed then
sToolbar="siteBuilderMailSource"
editor.config("width") = 520
else
sToolbar="siteBuilderMailNoSource"
editor.config("width") = 520
end if
else
sToolbar="siteBuilderMail"
editor.config("width") = 520
end if
if bThemeEmbed then
bAllowUpload=true
end if
case "siteBuilderFooter"
editor.config("width") = 700
editor.config("height") = 400
sToolbar="siteBuilder"
end select
if bAllowUpload then
editor.instanceConfig("filebrowserBrowseUrl")=C_DIRECTORY_QUICKERSITE & "/asp/assetmanager/assetmanagerIF.asp?showInsertButton=true"
editor.instanceConfig("filebrowserWindowWidth")=820
editor.instanceConfig("filebrowserWindowHeight")=580
end if
editor.config("toolbar")=sToolbar
dumpFCKInstance=editor.editor (sFieldName,sValue)
end select
end function
sub createFCKInstance(sValue,sToolbar,sFieldName)
Response.Write dumpFCKInstance(sValue,sToolbar,sFieldName)
end sub
function numberList(startNr,stopNr,interval,selected)
dim i
for i=startNr to stopNr step interval
numberList=numberList& "<option value=" & i 
if convertGetal(selected)=convertGetal(i) then numberList=numberList & " selected"
numberList=numberList& ">" & i & "</option>"
next
end function
function convert2 (byref getal)
if len(getal)=1 then 
convert2="0"&getal
else
convert2=getal
end if
end function
function jaNeen(value)
if value then
jaNeen=l("yes")
else
jaNeen=l("no")
end if
end function
function convertCheckedYesNo(sValue)
if sValue="checked" then
convertCheckedYesNo=l("yes")
else
convertCheckedYesNo=l("no")
end if
end function

Function GetEmailValidator()

  Set GetEmailValidator = New RegExp

  GetEmailValidator.Pattern = "^((?:[A-Z0-9_%+-]+\.?)+)@((?:[A-Z0-9-]+\.)+[A-Z]{2,40})$"

  GetEmailValidator.IgnoreCase = True

End Function

' Action: checks if an email is correct. 
' Parameter: Email address 
' Returned value: on success it returns True, else False. 
Function CheckEmailSyntax(strEmail) 

Dim EmailValidator : Set EmailValidator = GetEmailValidator()

 CheckEmailSyntax=EmailValidator.Test(strEmail)
 
 exit function
 
On Error Resume Next
    Dim strArray 
    Dim strItem 
    Dim i 
    Dim c 
  
  
    ' assume the email address is correct  
    CheckEmailSyntax = True 
    
    ' split the email address in two parts: name@domain.ext 
    strArray = Split(strEmail, "@") 
  
    ' if there are more or less than two parts  
    If UBound(strArray) <> 1 Then 
        CheckEmailSyntax = False               
        Exit Function 
    End If 
  
    ' check each part 
    For Each strItem In strArray 
        ' no part can be void 
        If Len(strItem) <= 0 Then 
            CheckEmailSyntax = False           
            Exit Function 
        End If 
        
        ' check each character of the part 
        ' only following "abcdefghijklmnopqrstuvwxyz_-." 
        ' characters and the ten digits are allowed 
        For i = 1 To Len(strItem) 
               c = LCase(Mid(strItem, i, 1)) 
               ' if there is an illegal character in the part 
               If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then 
                   CheckEmailSyntax = False                                    
                   Exit Function 
               End If 
        Next 
   
      ' the first and the last character in the part cannot be . (dot) 
        If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then 
           CheckEmailSyntax = False	   
           Exit Function 
        End If 
    Next 
  
    ' the second part (domain.ext) must contain a . (dot) 
    If InStr(strArray(1), ".") <= 0 Then 
       CheckEmailSyntax = False	 
        Exit Function 
    End If 
  
    ' check the length oh the extension  
    i = Len(strArray(1)) - InStrRev(strArray(1), ".") 
    ' the length of the extension can be only 2, 3, or 4 
    ' to cover the new "info" extension 
    If i <> 2 And i <> 3 And i <> 4 Then 
        CheckEmailSyntax = False
        Exit Function 
    End If 
    ' after . (dot) cannot follow a . (dot) 
    If InStr(strEmail, "..") > 0 Then 
        CheckEmailSyntax = False
        Exit Function 
    End If 
On Error Goto 0
End Function
' generate a random password
Function GeneratePassWord()
     
    Dim Lenght,Index,LetterNumber,CapitalNumber,Password
    Lenght = 0
    
    Randomize
    Lenght = Int(7 * Rnd)
    Lenght = Lenght + 6
    
    'Randomize(cint(lenght))
    
    For Index = 1 To cint(lenght)
             
        LetterNumber = Int(25 * Rnd)
        
        CapitalNumber = Rnd
        
        If (CapitalNumber < 0.5) Then
            Password = Password & Chr(65 + LetterNumber)
        Else
            Password = Password & Chr(97 + LetterNumber)
        End If
    
    Next 
        
    GeneratePassWord = ucase(Password)
    GeneratePassWord=left(GeneratePassWord,6)
    GeneratePassWord=left(GeneratePassWord,2)&convertStr(int(rnd*10))&mid(GeneratePassWord,3,2)&convertStr(int(rnd*10))&mid(GeneratePassWord,5,2)
   
End Function
function convertChecked(bValue)
if convertBool(bValue) then
convertChecked=" checked='checked' "
else
convertChecked=" "
end if
end function
Function LinkURLs(tempTxt)
tempTxt=convertStr(tempTxt)
'Make <br /> from linebreaks
tempTxt=replace(tempTxt,vbcrlf,"<br />",1,-1,1)
Dim regEx
Set regEx = New RegExp
regEx.Global = True
regEx.IgnoreCase = True
'Hyperlink Email Addresses
regEx.Pattern = "([_.a-z0-9-]+@[_.a-z0-9-]+\.[a-z]{2,3})" 
tempTxt = regEx.Replace(tempTxt, "<a href=""mailto:$1"">$1</a>")
'Hyperlink URL's
regEx.Pattern = "((www\.|(http|https|ftp|news|file)+\:\/\/)[_.a-z0-9-]+\.[a-z0-9\/_:@=.+?,##%&~-]*[^.|\'|\# |!|\(|?|,| |>|<|;|\)])"
tempTxt = regEx.Replace(tempTxt, "<a href=""$1"" target = '_blank'>$1</a>")          
'Make <a href="www = <a href="http://www
tempTxt = Replace(tempTxt, "href=""www", "href=""http://www",1,-1,1)	  
LinkURLs = tempTxt
End Function 
Function isBetween(startdate,checkdate,enddate)
if isLeeg(enddate) and isLeeg(startdate) then
isBetween=true
exit function
end if
if not isLeeg(enddate) and not isLeeg(startdate) then
if date<=enddate and date>=startdate then
isBetween=true
else
isBetween=false
end if
exit function
end if 
if not isLeeg(startdate) and isLeeg(enddate) then
if date<startdate then
isBetween=false
else
isBetween=true
end if
exit function
end if
if not isLeeg(enddate) and isLeeg(startdate) then
if date>enddate then
isBetween=false
else
isBetween=true
end if
end if
end function
function convertV(bool)
if bool then
convertV="V"
else
convertV="X"
end if
end function
Const f_dictKey  = 1
Const f_dictItem = 2
Function SortDictionary(objDict)
  ' declare our variables
  Dim strDict()
  Dim objKey
  Dim strKey,strItem
  Dim X,Y,Z
  ' get the dictionary count
  Z = objDict.Count
  ' we need more than one item to warrant sorting
  If Z > 1 Then
    ' create an array to store dictionary information
    ReDim strDict(Z,2)
    X = 0
    ' populate the string array
    For Each objKey In objDict
        strDict(X,f_dictKey)  = CStr(objKey)
        strDict(X,f_dictItem) = CStr(objDict(objKey))
        X = X + 1
    Next
    ' perform a a shell sort of the string array
    For X = 0 to (Z - 2)
      For Y = X to (Z - 1)
        If StrComp(strDict(X,f_dictKey),strDict(Y,f_dictKey),vbTextCompare) > 0 Then
            strKey  = strDict(X,f_dictKey)
            strItem = strDict(X,f_dictItem)
            strDict(X,f_dictKey)  = strDict(Y,f_dictKey)
            strDict(X,f_dictItem) = strDict(Y,f_dictItem)
            strDict(Y,f_dictKey)  = strKey
            strDict(Y,f_dictItem) = strItem
        End If
      Next
    Next
    ' erase the contents of the dictionary object
    objDict.RemoveAll
    ' repopulate the dictionary with the sorted information
    For X = 0 to (Z - 1)
      objDict.Add strDict(X,f_dictKey), strDict(X,f_dictItem)
    Next
  End If
End Function
  Function SortDictionary2(objDict, intSort)
    ' declare constants
    Const dictKey  = 1
    Const dictItem = 2
    ' declare our variables
    Dim strDict()
    Dim objKey
    Dim strKey,strItem
    Dim X,Y,Z
    ' get the dictionary count
    Z = objDict.Count
    ' we need more than one item to warrant sorting
    If Z > 1 Then
      ' create an array to store dictionary information
      ReDim strDict(Z,2)
      X = 0
      ' populate the string array
      For Each objKey In objDict
          strDict(X,dictKey)  = CStr(objKey)
          strDict(X,dictItem) = CStr(objDict(objKey))
          X = X + 1
      Next
      ' perform a a shell sort of the string array
      For X = 0 To (Z - 2)
        For Y = X To (Z - 1)
          If StrComp(strDict(X,intSort),strDict(Y,intSort),vbTextCompare) > 0 Then
              strKey  = strDict(X,dictKey)
              strItem = strDict(X,dictItem)
              strDict(X,dictKey)  = strDict(Y,dictKey)
              strDict(X,dictItem) = strDict(Y,dictItem)
              strDict(Y,dictKey)  = strKey
              strDict(Y,dictItem) = strItem
          End If
        Next
      Next
      ' erase the contents of the dictionary object
      objDict.RemoveAll
      ' repopulate the dictionary with the sorted information
      For X = 0 To (Z - 1)
        objDict.Add strDict(X,dictKey), strDict(X,dictItem)
      Next
    End If
  End Function
Function PrintTimer(byVal dblTimer) 
' Developed by Ryan Trudelle-Schwarz for www.mamanze.com  
   
' This is a wrapper function for the PrintInterval function. Use  
' this function if you want to check for the midnight rollover.  
   
  Dim dblTimer2 
   
' Grab the end time.  
  dblTimer2 = Timer() 
   
' Check for the midnight rollover.  
  If dblTimer2 < dblTimer Then 
    dblTimer2 = dblTimer2 + 86400 
  End If 
   
' Call the function that will give us a pretty result.  
  PrintTimer = PrintInterval(dblTimer2 - dblTimer) 
End Function 
   
Function PrintInterval(byVal dblMilliSecond) 
' Developed by Ryan Trudelle-Schwarz for www.mamanze.com  
  Dim intMilliSecond, intSecond, intMinute, intHour 
  Dim strReturn 
   
  strReturn =""   
   
' Determine the number of milliseconds.  
  intMilliSecond = Int(dblMilliSecond*1000) mod 1000 
   
' Determine the number of seconds. This is not the second value  
' yet, just the number of seconds.  
  intSecond = Int(dblMilliSecond) 
   
' Determine the number of minutes, simply divide the total number  
' of seconds by 60 and get the real number result.  
  intMinute = Int(intSecond / 60) 
   
' Now we modulus the seconds by 60 to form the seconds value.  
  intSecond = intSecond mod 60 
   
' Compute the Hours value by dividing the minutes by 60.  
  intHour = Int(intMinute / 60) 
   
' Compute the actual minute value by getting the modulus of the  
' total number of minutes and 60.  
  intMinute = intMinute mod 60 
   
' If the timer took more then a hour then display the hours.  
  If intHour > 0 Then 
    If intHour = 1 Then 
      strReturn = strReturn & intHour &" Hour "   
    Else 
      strReturn = strReturn & intHour &" Hours "   
    End If 
  End If 
   
' If the timer took more then a minute then display the minutes.  
  If intMinute > 0 Then 
    If intMinute = 1 Then 
      strReturn = strReturn & intMinute &" Minute "   
    Else 
      strReturn = strReturn & intMinute &" Minutes "   
    End If 
  End If 
   
' If the timer took more then a second then display the seconds.  
  If intSecond > 0 Then 
    If intSecond = 1 Then 
      strReturn = strReturn & intSecond &" Second "   
    Else 
      strReturn = strReturn & intSecond &" Seconds "   
    End If 
  End If 
   
' If the timer took more then a millisecond then display the  
' milliseconds. Also, if the script took no time then display 0  
' milliseconds.  
   
  If strReturn ="" OR intMilliSecond > 0 Then 
    If intMilliSecond = 1 Then 
      strReturn = strReturn & intMilliSecond &" MilliSecond"   
    Else 
      strReturn = strReturn & intMilliSecond &" MilliSeconds"   
    End If 
  End If 
   
  PrintInterval = strReturn 
End Function 
function quotRep(sValue)
if isLeeg(sValue) then
quotRep=sValue
else
quotRep=replace(sValue,"""","&quot;",1,-1,1)
quotRep=replace(quotRep,"<","&lt;",1,-1,1)
quotRep=replace(quotRep,">","&gt;",1,-1,1)
end if
end function
function quotRepJS(sValue)
quotRepJS=replace(sValue,"'","\'",1,-1,1)
end function
function prepareForExport(value)
prepareForExport=replace(value,vbtab,"",1,-1,1)
prepareForExport=replace(prepareForExport,"src="& """" & C_VIRT_DIR & Application("QS_CMS_userfiles"),"src=" & """" & customer.sVDUrl & Application("QS_CMS_userfiles"),1,-1,1)
prepareForExport=replace(prepareForExport,"src='" & C_VIRT_DIR & Application("QS_CMS_userfiles"),"src='" & customer.sVDUrl & Application("QS_CMS_userfiles"),1,-1,1)
prepareForExport=replace(prepareForExport,"src=" & Application("QS_CMS_userfiles"),"src=" & customer.sVDUrl & Application("QS_CMS_userfiles"),1,-1,1)
prepareForExport=replace(prepareForExport,"src="& """" & Application("QS_CMS_userfiles"),"src=" & """" & customer.sVDUrl & Application("QS_CMS_userfiles"),1,-1,1)
prepareForExport=replace(prepareForExport,"src='"& Application("QS_CMS_userfiles"),"src='" & customer.sVDUrl & Application("QS_CMS_userfiles"),1,-1,1)
prepareForExport=replace(prepareForExport,"href="& """" & "default.asp","href=" & """" & customer.sQSUrl & "/default.asp",1,-1,1)
prepareForExport=replace(prepareForExport,"href="& """" & "/default.asp","href=" & """" & customer.sQSUrl & "/default.asp",1,-1,1)
prepareForExport=replace(prepareForExport,"href="& Application("QS_CMS_userfiles"),"href=" & customer.sVDUrl & Application("QS_CMS_userfiles"),1,-1,1)
prepareForExport=replace(prepareForExport,"href="& """" & Application("QS_CMS_userfiles"),"href=" & """" & customer.sVDUrl & Application("QS_CMS_userfiles"),1,-1,1)
prepareForExport=replace(prepareForExport,"href='default.asp","href='"& customer.sQSUrl & "/default.asp",1,-1,1)
prepareForExport=replace(prepareForExport,"href='/default.asp","href='"& customer.sQSUrl & "/default.asp",1,-1,1)
prepareForExport=replace(prepareForExport,"src='"& C_DIRECTORY_QUICKERSITE & "/fckeditor","src='"& customer.sQSUrl & "/fckeditor",1,-1,1)
prepareForExport=replace(prepareForExport,"src='"& C_DIRECTORY_QUICKERSITE & "/fixedImages","src='"& customer.sQSUrl & "/fixedImages",1,-1,1)
prepareForExport=replace(prepareForExport,"src="& """" & C_DIRECTORY_QUICKERSITE & "/fckeditor","src="& """" & customer.sQSUrl & "/fckeditor",1,-1,1)
prepareForExport=replace(prepareForExport,"src="& """" & C_DIRECTORY_QUICKERSITE & "/fixedImages","src="& """" & customer.sQSUrl & "/fixedImages",1,-1,1)
prepareForExport=replace(prepareForExport,"action="& """" & "default.asp","action=" & """" & customer.sQSUrl & "/default.asp",1,-1,1)
prepareForExport=replace(prepareForExport,"action='default.asp","action='" & customer.sQSUrl & "/default.asp",1,-1,1)
prepareForExport=replace(prepareForExport,"action=default.asp","action=" & customer.sQSUrl & "/default.asp",1,-1,1)
prepareForExport=replace(prepareForExport,"src=" & """" & C_DIRECTORY_QUICKERSITE & "/showThumb.aspx","src=" & """" & customer.sQSUrl & "/showThumb.aspx",1,-1,1)
prepareForExport=replace(prepareForExport,"href=" & """" & C_DIRECTORY_QUICKERSITE & "/showThumb.aspx","href=" & """" & customer.sQSUrl & "/showThumb.aspx",1,-1,1)
prepareForExport=replace(prepareForExport,"src='" & C_DIRECTORY_QUICKERSITE & "/showThumb.aspx","src='" & customer.sQSUrl & "/showThumb.aspx",1,-1,1)
prepareForExport=replace(prepareForExport,"href='" & C_DIRECTORY_QUICKERSITE & "/showThumb.aspx","href='" & customer.sQSUrl & "/showThumb.aspx",1,-1,1)
prepareForExport=replace(prepareForExport,"'" & C_DIRECTORY_QUICKERSITE & "/showThumb.aspx","'" & customer.sQSUrl & "/showThumb.aspx",1,-1,1)
end function
function prepareforEmail(value)
prepareforEmail=value
if left(trim(prepareforEmail),1)<>"<" then
prepareforEmail="<p>"&prepareforEmail&"</p>"
end if
prepareforEmail=prepareForExport(prepareforEmail)
end function
function show(value)
if isleeg(value) then
show=""
else
show=LinkURLs(value)
end if
end function
function sqlNull(field,value)
if isLeeg(value) then
sqlNull=" " & field & " is null "
else
sqlNull=" " & field & "=" & value & " "
end if
end function
function getIcon(sLabel,sType,sUrl,sJs,sUniqueKey)
getIcon="<script type='text/javascript'>getIcon('"& quotRepJS(sLabel) & "','" & quotRepJS(sType)& "','" &quotRepJS(sUrl)& "','" &quotRepJS(sJs)& "','" &quotRepJS(sUniqueKey) &"','');</script>" & vbcrlf
end function
function getIconPP(sLabel,sType,sUrl,sJs,sUniqueKey,sClassname)
getIconPP="<script type='text/javascript'>getIcon('"& quotRepJS(sLabel) & "','" & quotRepJS(sType)& "','" &quotRepJS(sUrl)& "','" &quotRepJS(sJs)& "','" &quotRepJS(sUniqueKey) &"','" & sClassname & "');</script>" & vbcrlf
end function
function openLinksInNW (value)
openLinksInNW=replace(value,"<a ","<a target='"&generatePassword()&"' ",1,-1,1)
end function
function getCP(formField)
getCP="<a href='#' title="& """" & quotRep(l("select")) & """" & " onclick=""javascript: openPopUpWindow('colorPicker','colorPicker/picker.asp?"
getCP=getCP& "paramCode="& formField &"&amp;cColor='+ URLEncode(document.mainform."&formField&".value),450,380);"">"
getCP=getCP& "<img alt="& """" & quotRep(l("select")) & """" & " align='middle' border='0' src=" & """" & C_DIRECTORY_QUICKERSITE & "/fixedImages/pickColor.gif" & """" & " /></a>"
end function
Public function aantalDagenWithDate(resetDate)
aantalDagenWithDate=convertGetal(dateDiff("d",resetDate,date))
if aantalDagenWithDate=0 then aantalDagenWithDate=1
end function
function getImageLink(iAutoClose,sResizePicTo,sPicExt,iTSize,iId,sTitle)
getImageLink="<a class=""QSPPIMG"" href=""" & C_DIRECTORY_QUICKERSITE &"/showThumb.aspx?FSR=0&amp;maxSize="&sResizePicTo 
getImageLink=getImageLink&"&amp;img=" &server.URLEncode(C_VIRT_DIR & Application("QS_CMS_userfiles") & sPicExt)
getImageLink=getImageLink&""">"
getImageLink=getImageLink&"<img style='margin:0px;border-style:none' id='itemP"&iId&"' border='0' src='"&C_DIRECTORY_QUICKERSITE&"/showThumb.aspx?FSR=0&amp;maxSize="&iTSize&"&amp;img="
getImageLink=getImageLink& server.URLEncode(C_VIRT_DIR & Application("QS_CMS_userfiles") & sPicExt) 
getImageLink=getImageLink& "' alt=" & """" & sanitize(sTitle) & """" & " /></a>"
end function
function getImageLinkLB(sResizePicTo,sPicExt,iTSize,iId,sTitle,iThumbSize)
getImageLinkLB="<a href='"& C_DIRECTORY_QUICKERSITE & "/showThumb.aspx?FSR=0&amp;img=" & server.URLEncode(C_VIRT_DIR & Application("QS_CMS_userfiles") & sPicExt)
getImageLinkLB=getImageLinkLB & "&amp;maxSize=" & sResizePicTo & "' class=""QSPPIMG"" "
getImageLinkLB=getImageLinkLB & " title=""" & sanitize(sTitle) & """>"
getImageLinkLB=getImageLinkLB & "<img border=""0"" style=""margin:0px;border-style:none"" alt=""" & sanitize(sTitle) & """"
getImageLinkLB=getImageLinkLB & " src="""& C_DIRECTORY_QUICKERSITE &"/showThumb.aspx?FSR=0&amp;maxSize=" & iThumbSize & "&amp;img="& server.URLEncode (C_VIRT_DIR & Application("QS_CMS_userfiles") & sPicExt) & """ /></a>"
dim sAddHeaderQS
sAddHeaderQS="<script type=""text/javascript"">" & vbcrlf
sAddHeaderQS=sAddHeaderQS & "$(document).ready(function(){" & vbcrlf
sAddHeaderQS=sAddHeaderQS & "$(""a[rel='lightbox']"").colorbox({initialWidth:150, initialHeight:100, photo:true, slideshow:false, transition:""elastic"""
sAddHeaderQS=sAddHeaderQS & "});" & vbcrlf
sAddHeaderQS=sAddHeaderQS & "});" & vbcrlf
sAddHeaderQS=sAddHeaderQS & "</script>" & vbcrlf
headerDictionary.Add generatePassword, sAddHeaderQS
end function
function getDateWithoutSlash
getDateWithoutSlash=convertEuroDate(date)
getDateWithoutSlash=replace(getDateWithoutSlash,"/","")
end function
function getSQLDateFunction()
select case QS_DBS
case 1 'Access
getSQLDateFunction="date()"
case 2 'SQL Server
getSQLDateFunction="getdate()"
case 3 'MySQL
getSQLDateFunction="curdate()"
end select
end function
function getSQLDate(datevalue)
select case QS_DBS
case 1 'Access
getSQLDate="#" & year(datevalue) & "-" & month(datevalue) & "-" & day(datevalue)& "# "
case 2 'SQL Server
getSQLDate=date2sql(datevalue)
case 3 'MySQL
getSQLDate=date2mysql(datevalue)
end select
end function
function getSQLBoolean(boolValue)
select case QS_DBS
case 1 'Access
if boolValue then
getSQLBoolean="True"
else
getSQLBoolean="False"
end if
case 2,3 'SQL Server
if boolValue then
getSQLBoolean="1"
else
getSQLBoolean="0"
end if
end select
end function
function getTOPSyntax(topvalue,sqlstring)
select case QS_DBS
case 1,2 'Access and sql server
getTOPSyntax=replace(sqlstring,"[TOP]"," top "&topvalue,1,-1,1)
getTOPSyntax=replace(getTOPSyntax,"[LIMIT]","",1,-1,1)
case 3 'SQL Server
getTOPSyntax=replace(sqlstring,"[TOP]","",1,-1,1)
getTOPSyntax=replace(getTOPSyntax,"[LIMIT]"," limit "&topvalue,1,-1,1)
end select
end function
sub cleanUPASP()
set selectedPage	= nothing
set message	= nothing
set customer	= nothing
set logon	= nothing
set db	= nothing
end sub
sub clearMenuCache
dim copyCID, copyUfiles
copyCID	= cId
copyUfiles	= Application("QS_CMS_userfiles")
Application.Contents.removeAll()
Application("QS_CMS_iCustomerID")	= copyCID
Application("QS_CMS_userfiles")	= copyUfiles
Application("QS_CMS_C_VIRT_DIR")	= C_VIRT_DIR
Application("QS_CMS_C_DIRECTORY_QUICKERSITE")	= C_DIRECTORY_QUICKERSITE
Application("QS_ASPX")	= QS_ASPX
end sub
function isValidFolderName(elementName)
 isValidFolderName=true
Dim invalcount, invalidList,i
    invalcount = 0 
    invalidList = ",<.>?;:'@#~]}[{=+)(*&^%$£!`¬| -_" 
 
    ' check for " which can't be inside the string 
 
    if instr(elementName,chr(34))>0 then 
        invalcount = 1 
    else 
        ' loop through, making sure no characters 
        ' are in the 'reserved characters' list 
 
        for i = 1 to len(invalidList) 
            if instr(elementName,mid(invalidList,i,1))>0 then 
                invalcount = 1 
                exit for 
            end if 
        next 
    end if 
 
    if invalcount > 0 then 
       isValidFolderName=false
    end if 
end function
Function GetFileExtension(sFileName)
Dim Pos
Pos = instrrev(sFileName, ".")
If Pos>0 Then GetFileExtension = Mid(sFileName, Pos+1)
End Function
function allowedFileTypesforThumbing
set allowedFileTypesforThumbing=server.CreateObject ("scripting.dictionary")
'image types
allowedFileTypesforThumbing.Add "jpg",""
allowedFileTypesforThumbing.Add "jpeg",""
allowedFileTypesforThumbing.Add "jpe",""
allowedFileTypesforThumbing.Add "jp2",""
allowedFileTypesforThumbing.Add "jfif",""
allowedFileTypesforThumbing.Add "gif",""
allowedFileTypesforThumbing.Add "bmp",""
allowedFileTypesforThumbing.Add "png",""
allowedFileTypesforThumbing.Add "tif",""
allowedFileTypesforThumbing.Add "tiff",""
end function
function allowedFileTypes
set allowedFileTypes=server.CreateObject ("scripting.dictionary")
'image types
allowedFileTypes.Add "jpg",""
allowedFileTypes.Add "jpeg",""
allowedFileTypes.Add "jpe",""
allowedFileTypes.Add "jp2",""
allowedFileTypes.Add "jfif",""
allowedFileTypes.Add "gif",""
allowedFileTypes.Add "bmp",""
allowedFileTypes.Add "png",""
allowedFileTypes.Add "psd",""
allowedFileTypes.Add "eps",""
allowedFileTypes.Add "ico",""
allowedFileTypes.Add "tif",""
allowedFileTypes.Add "tiff",""
allowedFileTypes.Add "ai",""
allowedFileTypes.Add "raw",""
allowedFileTypes.Add "tga",""
allowedFileTypes.Add "mng",""
allowedFileTypes.Add "svg",""
'doctype
allowedFileTypes.Add "doc",""
allowedFileTypes.Add "rtf",""
allowedFileTypes.Add "txt",""
allowedFileTypes.Add "wpd",""
allowedFileTypes.Add "wps",""
allowedFileTypes.Add "csv",""
allowedFileTypes.Add "xml",""
allowedFileTypes.Add "xsd",""
allowedFileTypes.Add "sql",""
allowedFileTypes.Add "pdf",""
allowedFileTypes.Add "xls",""
allowedFileTypes.Add "mdb",""
allowedFileTypes.Add "ppt",""
allowedFileTypes.Add "docx",""
allowedFileTypes.Add "xlsx",""
allowedFileTypes.Add "pptx",""
allowedFileTypes.Add "ppsx",""
allowedFileTypes.Add "artx",""
'media types
allowedFileTypes.Add "mp3",""
allowedFileTypes.Add "wma",""
allowedFileTypes.Add "mid",""
allowedFileTypes.Add "midi",""
allowedFileTypes.Add "mp4",""
allowedFileTypes.Add "mpg",""
allowedFileTypes.Add "mpeg",""
allowedFileTypes.Add "wav",""
allowedFileTypes.Add "ram",""
allowedFileTypes.Add "ra",""
allowedFileTypes.Add "avi",""
allowedFileTypes.Add "mov",""
allowedFileTypes.Add "flv",""
allowedFileTypes.Add "m4a",""
allowedFileTypes.Add "m4v",""
allowedFileTypes.Add "ogv",""
allowedFileTypes.Add "ogg",""
allowedFileTypes.Add "webm",""
allowedFileTypes.Add "ics",""

'internet related types
allowedFileTypes.Add "htm",""
allowedFileTypes.Add "html",""
allowedFileTypes.Add "css",""
allowedFileTypes.Add "swf",""
allowedFileTypes.Add "js",""
allowedFileTypes.Add "less",""

'allowedFileTypes.Add "vbs",""
'allowedFileTypes.Add "swt",""
'compression types
allowedFileTypes.Add "rar",""
allowedFileTypes.Add "zip",""
allowedFileTypes.Add "tar",""
allowedFileTypes.Add "gz",""
'fonts
allowedFileTypes.Add "woff",""
allowedFileTypes.Add "ttf",""
allowedFileTypes.Add "eot",""

end function
function removeEmptyP(sValue)
if sValue="<p>&#160;</p>" then sValue=""
if sValue="<p>&nbsp;</p>" then sValue=""
if sValue="<div>&nbsp;</div>" then sValue=""
if sValue="<br class=""innova"">" then sValue=""
removeEmptyP=sValue
end function
Sub Format_RFC822_DateAndTime(dtTime, sFormattedDate)
     
      Select Case DatePart("w", dtTime)
         Case vbSunday
            sFormattedDate = "Sun, "
         Case vbMonday
            sFormattedDate = "Mon, "
         Case vbTuesday
            sFormattedDate = "Tue, "
         Case vbWednesday
            sFormattedDate = "Wed, "
         Case vbThursday
            sFormattedDate = "Thu, "
         Case vbFriday
            sFormattedDate = "Fri, "
         Case vbSaturday
            sFormattedDate = "Sat, "
         Case Else
            sFormattedDate = ""
      End Select
      sFormattedDate = sFormattedDate & Right(CStr(DatePart("d", dtTime) + 100), 2) & " "
      
      Select Case DatePart("m", dtTime)
         Case 1
            sFormattedDate = sFormattedDate & "Jan "
         Case 2
            sFormattedDate = sFormattedDate & "Feb "
         Case 3
            sFormattedDate = sFormattedDate & "Mar "
         Case 4
            sFormattedDate = sFormattedDate & "Apr "
         Case 5
            sFormattedDate = sFormattedDate & "May "
         Case 6
            sFormattedDate = sFormattedDate & "Jun "
         Case 7
            sFormattedDate = sFormattedDate & "Jul "
         Case 8
            sFormattedDate = sFormattedDate & "Aug "
         Case 9
            sFormattedDate = sFormattedDate & "Sep "
         Case 10
            sFormattedDate = sFormattedDate & "Oct "
         Case 11
            sFormattedDate = sFormattedDate & "Nov "
         Case 12
            sFormattedDate = sFormattedDate & "Dec "
      End Select
      
'      sFormattedDate = sFormattedDate & Format(dtTime, "mmm ")  ' Systems settings dependent (locale)
      sFormattedDate = sFormattedDate & DatePart("yyyy", dtTime) & " "
      sFormattedDate = sFormattedDate & Right(CStr(DatePart("h", dtTime) + 100), 2) & ":"
      sFormattedDate = sFormattedDate & Right(CStr(DatePart("n", dtTime) + 100), 2) & ":"
      sFormattedDate = sFormattedDate & Right(CStr(DatePart("s", dtTime) + 100), 2) & " "
      sFormattedDate = sFormattedDate & "GMT"   ' Hard coded time zone
      
End Sub
Function UserIP()
' This returns the True IP of the client calling the requested page
' Checks to see if HTTP_X_FORWARDED_FOR has a value then the client is operating via a proxy
UserIP = Server.HTMLEncode(Server.URLEncode(Request.ServerVariables ( "HTTP_X_FORWARDED_FOR" )))
If UserIP = "" Then
UserIP = Request.ServerVariables ( "REMOTE_ADDR" )
End if
End Function
Function devVersion()
devVersion=lcase(Request.ServerVariables ("http_host"))="dev.quickersite.com"
'devVersion=false
end function
function wrapInHTML(body,subject)
if instr(1,body,"<html>",vbTextCompare)=0 then
wrapInHTML="<!DOCTYPE html><html>" & vbcrlf & "<head>" & vbcrlf & "<meta HTTP-EQUIV=""Content-Type"" content=""text/html; charset=" & QS_CHARSET & """>" & vbcrlf & "<title>"& subject &"</title>" & vbcrlf & "</head>" & vbcrlf & "<body "
if not isLeeg(sOWBodyBGColor) then
wrapInHTML=wrapInHTML & " style=""background-color:" & sOWBodyBGColor & """ bgColor=""" & sOWBodyBGColor & """ "
end if
wrapInHTML=wrapInHTML & ">"&vbcrlf&body&vbcrlf&"</body>"&vbcrlf&"</html>"
else
wrapInHTML=body
end if
end function
sub removeApplication()
dim saveCid
saveCId=Application("QS_CMS_iCustomerID")
Application.Contents.removeAll()
Application("QS_CMS_iCustomerID")=saveCid
end sub
function sanitize(sValue)
sanitize=quotRep(sValue)
end function
function run(sScript)
if not isLeeg(sScript) then server.execute(sScript)
end function
function treatParentheses(value)
treatParentheses=replace(value,"(","\(",1,-1,1)
treatParentheses=replace(treatParentheses,")","\)",1,-1,1)
treatParentheses=replace(treatParentheses,"[","\[",1,-1,1)
treatParentheses=replace(treatParentheses,"]","\]",1,-1,1)
end function
Function urlDecode(str)
str = Replace(str, "%3F", "?",1,-1,1)
str = Replace(str, "%2F", "/",1,-1,1)
str = Replace(str, "%7C", "|",1,-1,1)
str = Replace(str, "%5C", "\",1,-1,1)
str = Replace(str, "%21", "!",1,-1,1)
str = Replace(str, "%40", "@",1,-1,1)
str = Replace(str, "%23", "#",1,-1,1)
str = Replace(str, "%24", "$",1,-1,1)
str = Replace(str, "%25", "%",1,-1,1)
str = Replace(str, "%5E", "^",1,-1,1)
str = Replace(str, "%26", "&",1,-1,1)
str = Replace(str, "%2A", "*",1,-1,1)
str = Replace(str, "%28", "(",1,-1,1)
str = Replace(str, "%29", ")",1,-1,1)
str = Replace(str, "%7B", "{",1,-1,1)
str = Replace(str, "%7D", "}",1,-1,1)
str = Replace(str, "%3A", ":",1,-1,1)
str = Replace(str, "%3B", ";",1,-1,1)
str = Replace(str, "%2E", ".",1,-1,1)
str = Replace(str, "%2D", "-",1,-1,1)
str = Replace(str, "%5B", "[",1,-1,1)
str = Replace(str, "%5D", "]",1,-1,1)
str = Replace(str, "%2C", ",",1,-1,1)
str = Replace(str, "%3D", "=",1,-1,1)
str = Replace(str, "%2B", "+",1,-1,1)
str = Replace(str, "%2D", "-",1,-1,1)
str = Replace(str, "%5F", "_",1,-1,1)
str = Replace(str, "%7E", "~",1,-1,1)
str = Replace(str, "%60", "`",1,-1,1) 
str = Replace(str, "%27", "'",1,-1,1)
str = Replace(str, "%22", Chr(34),1,-1,1)
str = Replace(str, "%20", " ",1,-1,1)
str = Replace(str, "%3C", "<",1,-1,1)
str = Replace(str, "%3E", ">",1,-1,1)
urlDecode = str
End Function
function filterURL(sURL)
dim urlArr
urlArr=split(sURL,"/")
dim looper
for looper = lbound(urlArr) to ubound(urlArr)
if looper<3 then
filterURL=filterURL&urlArr(looper)&"/"
else
exit for
end if
next
end function
Function date2sql(date_value)
If isDate(date_value) Then
date2sql = "convert(smalldatetime, '" & year(date_value) & "-" & month(date_value) & "-" & day(date_value) & "', 120)"
Else
date2sql = "null"
End If
End Function
Function date2mysql(date_value)
If isDate(date_value) Then
date2mysql = "date('" & year(date_value) & "-" & month(date_value) & "-" & day(date_value) & "')"
Else
date2mysql = "null"
End If
End Function
function filterQuery(sUrl)
dim re,matches,match
set re=new regexp
re.Global =true
re.IgnoreCase =true
re.Pattern ="[?|&]+(q=|p=)+[^cache:]+[^&]+[\S]"
if re.Test(sUrl) then
set matches=re.Execute (sUrl)
for each match in matches
filterQuery=filterQuery&urlDecode(match.value) & " "
next
set matches=nothing
filterQuery=replace(filterQuery,"&q=","",1,-1,1)
filterQuery=replace(filterQuery,"?q=","",1,-1,1)
filterQuery=replace(filterQuery,"&p=","",1,-1,1)
filterQuery=replace(filterQuery,"?p=","",1,-1,1)
filterQuery=replace(filterQuery,"+"," ",1,-1,1)
filterQuery=replace(filterQuery,"&","",1,-1,1)
exit function
end if
end function
Function IsAlphaNumeric(byVal str)
If IsNull(str) Then str = ""
Dim ianRegEx
Set ianRegEx = New RegExp
ianRegEx.Pattern = "[^a-z0-9\/\_\-\.]"
ianRegEx.Global = True
ianRegEx.IgnoreCase = True
IsAlphaNumeric = (str = ianRegEx.Replace(str,""))
End Function
function tdir
if isEmpty(application(QS_textDirection)) or application(QS_textDirection)="" then
dim rs,sql
sql="select sDirection from tblLanguage where iId=" & customer.language
set rs=db.ExecuteLabels(sql)
application(QS_textDirection)=lcase(convertstr(rs(0)))
end if
tdir=application(QS_textDirection)
end function
function langCode
if isEmpty(application(QS_langcode)) or application(QS_langcode)="" then
dim rs,sql
sql="select sCode from tblLanguage where iId=" & customer.language
set rs=db.ExecuteLabels(sql)
application(QS_langcode)=lcase(convertstr(rs(0)))
end if
langCode=application(QS_langcode)
end function
sub redirectToHP
if isNumeriek(getHomepage.iId) then
if customer.bUserFriendlyURL and not isLeeg(getHomepage.sUserFriendlyURL) then
Response.Redirect (customer.sUrl&"/"&getHomepage.sUserFriendlyURL)
else
Response.Redirect (C_DIRECTORY_QUICKERSITE & "/default.asp?iId="& EnCrypt(getHomepage.iId))
end if
else
'waarschuwingsmail naar siteAmin!
ErrorReport "Geen homepage",err
end if
end sub
function clickEmail(value)
if not isLeeg(value) then
clickEmail="<a  href=" & """" & "mailto:" & value & """" & ">" & value & "</a>"
else
clickEmail=""
end if
end function
function getSecurityWarning
dim gUIP
gUIP=UserIP
if gUIP="::1" then
gUIP="127.0.0.1"
end if
dim sw
sw=l("explsecuritylogin")
sw=replace(sw,"[ip]",gUIP,1,-1,1)
sw=replace(sw,"[number]",QS_number_of_allowed_attempts_to_login,1,-1,1)
getSecurityWarning=sw
end function
function secCode
if isLeeg(session("secCode")) then
session("secCode")=GeneratePassWord
end if
secCode=session("secCode")
end function
function QS_secCodeHidden
QS_secCodeHidden="<input type='hidden' name='QSSEC' value='" & secCode & "' />"
end function
function QS_secCodeURL
QS_secCodeURL="QSSEC="&secCode
end function
function checkCSRF
if convertStr(Request("QSSEC")) <> convertStr(secCode) then
sendCSRFWarning()
end if
end function
function checkCSRF_Upload(qssec_value)
if convertStr(qssec_value) <> convertStr(secCode) then
sendCSRFWarning()
end if
end function
function sendCSRFWarning()
if not isLeeg(C_ADMINEMAIL) and isLeeg(application("CSRFMS" & UserIP)) then
application("CSRFMS" & UserIP)="mailSENT"
dim theMail
set theMail=new cls_mail_message
theMail.receiver=C_ADMINEMAIL
theMail.subject="CSRFProblem on "& customer.siteName
theMail.body="<p>Visitor details:</p>" & getVisitorDetails
'theMail.send
set theMail=nothing
end if
'Response.End
end function
function wrapInShadowCSS(picLink)
wrapInShadowCSS="<table border=""0"" cellpadding='0' cellspacing='0' style=""border-style:none"" id=""shadow-container""><tr><td><table  cellpadding='0' cellspacing='0' class=""shadow1""><tr><td><table  cellpadding='0' cellspacing='0' class=""shadow2""><tr><td><table  cellpadding='0' cellspacing='0' class=""shadow3""><tr><td><table  cellpadding='0' cellspacing='0' class=""container""><tr><td>"
wrapInShadowCSS=wrapInShadowCSS&picLink
wrapInShadowCSS=wrapInShadowCSS&"</td></tr></table></td></tr></table></td></tr></table></td></tr></table></td></tr></table>"
end function
dim smileyJS
smileyJS=""
function shortListOfSmilies(element)
shortListOfSmilies=""
if smileyJS="" then
smileyJS=vbcrlf&vbcrlf&"<script type=""text/javascript"">function wSm(el){" & vbcrlf
dim  smiley
dim cSimileys
set cSimileys=smilies
for each smiley in smilies
smileyJS=smileyJS&"document.write ('<a href=""#"" onclick=""javascript:return addSmiley\('+el+',\' " & (smiley) & " \'\);""><img style=""margin:0px;border-style:none"" border=""0"" src=""" & C_DIRECTORY_QUICKERSITE & "/fixedImages/smileys/" & cSimileys(smiley) & """ alt=""" & sanitize(replace(cSimileys(smiley),".gif","",1,-1,1)) & """ /><\/a> ');" & vbcrlf
next
set cSimileys=nothing
smileyJS=smileyJS & "return false;" & vbcrlf
smileyJS=smileyJS & "}" & vbcrlf
smileyJS=smileyJS & "function addSmiley(el,sm){" & vbcrlf
smileyJS=smileyJS & "el.value+=sm;"	& vbcrlf
smileyJS=smileyJS & "return false;" & vbcrlf
smileyJS=smileyJS & "}" & vbcrlf
smileyJS=smileyJS & "</script>"& vbcrlf& vbcrlf
shortListOfSmilies=smileyJS
end if
shortListOfSmilies=shortListOfSmilies & "<script type=""text/javascript"">wSm('" & element & "');</script>"&vbcrlf&vbcrlf
end function
dim p_smilies
set p_smilies=nothing
function smilies
if p_smilies is nothing then
set p_smilies=server.CreateObject ("scripting.dictionary")
p_smilies.Add "[:)]","happy.gif"
p_smilies.Add "[:laughing:]","laughing.gif"
p_smilies.Add "[:kiss:]","kiss.gif"
p_smilies.Add "[:smartboy:]","smartboy.gif"
p_smilies.Add "[:cheers:]","cheers.gif"
p_smilies.Add "[:conicalhat:]","conicalhat.gif"
p_smilies.Add "[:razz:]","razz.gif"
p_smilies.Add "[:innocent:]","innocent.gif"
p_smilies.Add "[:greenbiggrin:]","greenbiggrin.gif"
p_smilies.Add "[:frown:]","frown.gif"
p_smilies.Add "[:crazy:]","crazy.gif"
p_smilies.Add "[:yell:]","yell.gif"
p_smilies.Add "[:madashell:]","madashell.gif"
p_smilies.Add "[:yuck:]","yuck.gif"
p_smilies.Add "[:drunk:]","drunk.gif"
p_smilies.Add "[:vampire:]","vampire.gif"
p_smilies.Add "[:greencool:]","greencool.gif"
p_smilies.Add "[:ghost:]","ghost.gif"
p_smilies.Add "[:zombie:]","zombie.gif"
p_smilies.Add "[:working:]","working.gif"
end if
set smilies=p_smilies
end function
function addSmilies(sValue)
addSmilies=sValue
dim cSmilies,  smiley
set cSmilies=smilies
for each smiley in cSmilies
addSmilies=replace(addSmilies,smiley,"<img style='margin-top:0px;margin-left:0px;margin-right:0px;margin-bottom:-2px;border-style:none' border='0' src='" & C_DIRECTORY_QUICKERSITE & "/fixedImages/smileys/" & cSmilies(smiley) & "' alt='' />",1,-1,1)
next
set cSmilies=nothing
end function
dim currentContacts
set currentContacts=nothing
function getContact(id)
if currentContacts is nothing then
set currentContacts=server.CreateObject ("scripting.dictionary")
end if
if convertGetal(id)<>0 then 
if currentContacts.Exists (id) then
set getContact=currentContacts(id)
else
on error resume next
set getContact=new cls_contact
getContact.quickpick(id)
currentContacts.Add id,getContact
on error goto 0
end if
else
set getContact=new cls_contact
end if
end function
Function URLDecodeAJAX(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecodeAJAX = ""
       Exit Function
    End If
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ",1,-1,1)
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If
    URLDecodeAJAX = sOutput
End Function
Function PCase(ByVal strInput)' As String
        Dim I 'As Integer
        Dim CurrentChar, PrevChar 'As Char
        Dim strOutput 'As String
        PrevChar = ""
        strOutput = ""
        For I = 1 To Len(strInput)
            CurrentChar = Mid(strInput, I, 1)
            Select Case PrevChar
                Case "", " ", ".", "-", ",", """", "'"
                    strOutput = strOutput & UCase(CurrentChar)
                Case Else
                    strOutput = strOutput & LCase(CurrentChar)
            End Select
            PrevChar = CurrentChar
        Next 'I
        PCase = strOutput
    End Function 
function getArtHeaderPart
if bUseArtLoginTemplate then
getArtHeaderPart=ArtHeaderPartLOGIN
else
getArtHeaderPart=ArtHeaderPart
end if
end function
function getArtFooterPart
if bUseArtLoginTemplate then
getArtFooterPart=ArtFooterPartLOGIN
else
getArtFooterPart=ArtFooterPart
end if
end function
function getArtTopBannerAndMenu
if bUseArtLoginTemplate then
getArtTopBannerAndMenu=ArtTopBannerAndMenuLOGIN
else
getArtTopBannerAndMenu=ArtTopBannerAndMenu
end if
end function
Function SaveBinaryData(FileName, ByteArray)
  Const adTypeBinary = 1
  Const adSaveCreateOverWrite = 2
  
  'Create Stream object
  Dim BinaryStream
  Set BinaryStream = CreateObject("ADODB.Stream")
  
  'Specify stream type - we want To save binary data.
  BinaryStream.Type = adTypeBinary
  
  'Open the stream And write binary data To the object
  BinaryStream.Open
  BinaryStream.Write ByteArray
  
  'Save binary data To disk
  BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
End Function
function getArtLink(url,label,javascript,target,rel)
getArtLink="<span class=""art-button-wrapper"">"
getArtLink=getArtLink & "<span class=""art-button-l""> </span>"
getArtLink=getArtLink & "<span class=""art-button-r""> </span>"
getArtLink=getArtLink & "<a rel=""" & rel & """ onclick=""" & javascript & """ class=""art-button"" href=""" & url & """"
if not isLeeg(target) then
getArtLink=getArtLink & " target=""" & target & """ "
end if
getArtLink=getArtLink & ">" & label & "</a>"
getArtLink=getArtLink & "</span>"
end function
function getBacksiteURL(domain)
if right(domain,len(C_DIRECTORY_QUICKERSITE))=C_DIRECTORY_QUICKERSITE then
getBacksiteURL=domain & "/backsite"
else
getBacksiteURL=domain & C_DIRECTORY_QUICKERSITE & "/backsite"
end if
end function
function checkMobileBrowser
on error resume next
if isLeeg(customer.sMOBBrowsers) then exit function
if customer.sMOBurl="http://" and convertgetal(customer.iDefaultMobileTemplate)=0 then exit function
if isLeeg(customer.sMOBurl) and convertgetal(customer.iDefaultMobileTemplate)=0 then exit function
'if isLeeg(session("checkMobileBrowser")) then
dim user_agent,Regex,match,pattern
user_agent = Request.ServerVariables("HTTP_USER_AGENT") 
if isleeg(user_agent) then exit function
application("QSmostrecentUA")=left("<b>" & now() & "</b>: " & sanitize(user_agent) & vbcrlf & application("QSmostrecentUA"),750)
dim arrMB,i
arrMB=split(customer.sMOBBrowsers,vbcrlf)
for i=lbound(arrMB) to ubound(arrMB)
if not isLeeg(arrMB(i)) then
pattern=pattern & arrMB(i) & "|"
end if
next
pattern=left(pattern,len(pattern)-1)
Set Regex = New RegExp
With Regex
   .Pattern = "(" & pattern & ")"
   .IgnoreCase = True
   .Global = True
End With
 
match = Regex.Test(user_agent)
 
If match Then 
if convertGetal(customer.iDefaultMobileTemplate)<>0 then
session("iTemplateID")=customer.iDefaultMobileTemplate
elseif isleeg(session("checkMobileBrowser")) then
session("checkMobileBrowser")="y"
response.redirect(customer.sMOBurl)
end if
end if
'end if
on error goto 0
end function
function JQDatePicker(DateFieldName)
JQDatePicker=vbcrlf & "<script type=""text/javascript"">$(function() {$(""#" 
JQDatePicker=JQDatePicker & DateFieldName
JQDatePicker=JQDatePicker & """).datepicker({hideIfNoPrevNext:true, yearRange: '1900:" & year(date())+20 & "',changeMonth: true,changeYear: true,dateFormat:"""
select case customer.sDatumFormat
case QS_dateFM_US
JQDatePicker=JQDatePicker & "mm/dd/yy"
case else
JQDatePicker=JQDatePicker & "dd/mm/yy"
end select
JQDatePicker=JQDatePicker & """}).attr('readonly','readonly');});document.getElementById('" & DateFieldName & "').style.width='110px';</script>" & vbcrlf
'add reset
JQDatePicker=JQDatePicker & "<i><a href=""#"" onclick=""javascript:document.getElementById('" & DateFieldName & "').value='';return false;"" >" & lcase(convertStr(l("reset"))) & "</a></i> "
end function
function JQDatePickerFT(fromDate,toDate)
JQDatePickerFT=vbcrlf & "<script type=""text/javascript"">"
JQDatePickerFT=JQDatePickerFT&"$(function() {" & vbcrlf
JQDatePickerFT=JQDatePickerFT&"$( ""#" & fromDate & """ ).datepicker({" & vbcrlf
JQDatePickerFT=JQDatePickerFT&" defaultDate: ""+1w""," & vbcrlf
JQDatePickerFT=JQDatePickerFT&" dateFormat:"""
select case customer.sDatumFormat
case QS_dateFM_US
JQDatePickerFT=JQDatePickerFT & "mm/dd/yy"
case else
JQDatePickerFT=JQDatePickerFT & "dd/mm/yy"
end select
JQDatePickerFT=JQDatePickerFT&""", changeMonth: true," & vbcrlf
JQDatePickerFT=JQDatePickerFT&"numberOfMonths: 3," & vbcrlf
JQDatePickerFT=JQDatePickerFT&"onClose: function( selectedDate ) {" & vbcrlf
JQDatePickerFT=JQDatePickerFT&" $( ""#" & toDate & """ ).datepicker( ""option"", ""minDate"", selectedDate );" & vbcrlf
JQDatePickerFT=JQDatePickerFT&"}" & vbcrlf
JQDatePickerFT=JQDatePickerFT&" });" & vbcrlf
JQDatePickerFT=JQDatePickerFT&"$( ""#" & toDate & """ ).datepicker({" & vbcrlf
JQDatePickerFT=JQDatePickerFT&" defaultDate: ""+1w""," & vbcrlf
JQDatePickerFT=JQDatePickerFT&" dateFormat:"""
select case customer.sDatumFormat
case QS_dateFM_US
JQDatePickerFT=JQDatePickerFT & "mm/dd/yy"
case else
JQDatePickerFT=JQDatePickerFT & "dd/mm/yy"
end select
JQDatePickerFT=JQDatePickerFT&""", changeMonth: true," & vbcrlf
JQDatePickerFT=JQDatePickerFT&" numberOfMonths: 3," & vbcrlf
JQDatePickerFT=JQDatePickerFT&"onClose: function( selectedDate ) {" & vbcrlf
JQDatePickerFT=JQDatePickerFT&"$( ""#" & fromDate & """ ).datepicker( ""option"", ""maxDate"", selectedDate );" & vbcrlf
JQDatePickerFT=JQDatePickerFT&" }" & vbcrlf
JQDatePickerFT=JQDatePickerFT&" });" & vbcrlf
JQDatePickerFT=JQDatePickerFT&" });" & vbcrlf 
JQDatePickerFT=JQDatePickerFT&"</script>"
JQDatePickerFT=JQDatePickerFT & "<i><a href=""#"" onclick=""javascript:document.getElementById('" & fromDate & "').value='';document.getElementById('" & toDate & "').value='';return false;"" >reset dates</a></i> "
end function
function JQColorPicker(ColorFieldName)
JQColorPicker=vbcrlf & "<script type=""text/javascript"">$(""#" & ColorFieldName & """).spectrum({preferredFormat: ""hex"",showInitial: true,showInput: true});</script>" & vbcrlf
end function
function JQColorPickerD(ColorFieldName)
JQColorPickerD=vbcrlf & "<script type=""text/javascript"">$(""#" & ColorFieldName & """).spectrum({preferredFormat: ""hex"",showInitial: true,showInput: true,clickoutFiresChange: true});</script>" & vbcrlf
end function
Function HEXCOL2RGB(byval HexColor,percentage)  
if not isLeeg(HexColor) then
if len(HexColor)=4 then
HexColor=HexColor&right(HexColor,3)
end if
Dim Red,Green,Blue,Color
Color = Replace(HexColor, "#", "",1,-1,1)  
Blue = convertGetal("&H" & Mid(Color, 1, 2)) 
Red = convertGetal("&H" & Mid(Color, 3, 2))  
Green = convertGetal("&H" & Mid(Color, 5, 2)) 
HEXCOL2RGB = RGB(Red, Green, Blue) 
'16777215
'1118481
end if
End Function

function splitby(v,s)
v=removeHTML(v)
if isLeeg(v) then
splitby=""
exit function
end if
while s<len(v)
splitby=splitby & left(v,s) & "<br />"
v=right(v,len(v)-s)
wend
splitby=splitby & v
end function

Function RemoveHTMLComments( strText ) 
Dim nPos1
Dim nPos2

nPos1 = InStr(strText, "<!--") 
Do While nPos1 > 0 
nPos2 = InStr(nPos1 + 4, strText, "-->") 
If nPos2 > 0 Then 
strText = Left(strText, nPos1 - 1) & Mid(strText, nPos2 + 3) 
Else 
Exit Do 
End If 
nPos1 = InStr(strText, "<!--") 
Loop 

RemoveHTMLComments = strText 
End Function 

function removeDots(sValue)
removeDots=replace(sValue,".","",1,-1,1)
end function
%>
