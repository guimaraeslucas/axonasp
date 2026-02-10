<!-- #include file="security.asp"-->
<!-- #INCLUDE FILE='settings.asp' -->
<%

dim showInsertButton,showSelectButton
if request.querystring("showInsertButton")="true" then
session("showInsertButton")=true
showInsertButton=true
end if

if session("showInsertButton")=true then
showInsertButton=true
end if

if request.querystring("showSelectButton")="true" then
	showSelectButton=true
end if

'*** Permission ***
bReadOnly0=false
bReadOnly1=false
bReadOnly2=false
bReadOnly3=false
'*** /Permission ***

set oFSO = server.CreateObject ("Scripting.FileSystemObject")
set oFolder_base0 = oFSO.GetFolder(server.MapPath(arrBaseFolder(0))) 'base folder (Physical)
if arrBaseFolder(1)<>"" then set oFolder_base1 = oFSO.GetFolder(server.MapPath(arrBaseFolder(1))) 'base folder (Physical)
if arrBaseFolder(2)<>"" then set oFolder_base2 = oFSO.GetFolder(server.MapPath(arrBaseFolder(2))) 'base folder (Physical)
if arrBaseFolder(3)<>"" then set oFolder_base3 = oFSO.GetFolder(server.MapPath(arrBaseFolder(3))) 'base folder (Physical)

dim sOptions

%>
<!--#include file="i_upload_object_FSO.asp"-->
<%

'*** Permission ***
bWriteFolderAdmin=false
if Not IsEmpty(oFolder_base0) then
	if InStr(LCase(CStr(currFolder)),LCase(CStr(oFolder_base0.path)))<>0 AND bReadOnly0=true then bWriteFolderAdmin=true
end if
if Not IsEmpty(oFolder_base1) then
	if InStr(LCase(CStr(currFolder)),LCase(CStr(oFolder_base1.path)))<>0 AND bReadOnly1=true then bWriteFolderAdmin=true
end if	
if Not IsEmpty(oFolder_base2) then
	if InStr(LCase(CStr(currFolder)),LCase(CStr(oFolder_base2.path)))<>0 AND bReadOnly2=true then bWriteFolderAdmin=true
end if	
if Not IsEmpty(oFolder_base3) then
	if InStr(LCase(CStr(currFolder)),LCase(CStr(oFolder_base3.path)))<>0 AND bReadOnly3=true then bWriteFolderAdmin=true
end if
sFolderAdmin=""
if bWriteFolderAdmin then sFolderAdmin="style='display:none'"
'*** /Permission ***

function writeFolderSelections()
	response.Write "<select name='selCurrFolder' id='selCurrFolder' onchange='changeFolder()' class='inpSel'>"

	if arrBaseFolder(1)="" then
		response.write recursive(oFolder_base0,oFolder_base0,arrBaseName(0))
	elseif arrBaseFolder(2)="" then
		recursive oFolder_base0,oFolder_base0,arrBaseName(0) 
		response.write recursive(oFolder_base1,oFolder_base1,arrBaseName(1))
	elseif arrBaseFolder(3)="" then
		recursive oFolder_base0,oFolder_base0,arrBaseName(0) 	
		recursive oFolder_base1,oFolder_base1,arrBaseName(1)
		response.write recursive(oFolder_base2,oFolder_base2,arrBaseName(2))
	else
		recursive oFolder_base0,oFolder_base0,arrBaseName(0) 	
		recursive oFolder_base1,oFolder_base1,arrBaseName(1)
		recursive oFolder_base2,oFolder_base2,arrBaseName(2)
		response.write recursive(oFolder_base3,oFolder_base3,arrBaseName(3))
	end if
		
	response.Write "</select>"
	

end function

function recursive(oFolder,oFolder_base,sName)
	set oSubFolders = oFolder.SubFolders

    if InStr(1,oFolder.path,"_vti_cnf")=0 then
			sDisplayed = sName & Replace(Replace(oFolder.path,oFolder_base.path,""),"\","/")'display hanya bagian di dalam base folder, ex: "", "/gallery01", "/gallery02"
			
			if lcase(trim(CStr(currFolder)))=lcase(trim(CStr(oFolder.path))) then
				sOptions = sOptions & "<option value=""" & oFolder.path & """ selected>" & sDisplayed & "</option>" & vbCrLf
			else
				sOptions = sOptions & "<option value=""" & oFolder.path & """>" & sDisplayed & "</option>" & vbCrLf
			end if
    end if
    
	for each item in oSubFolders 
			recursive item,oFolder_base,sName
    next    
    
    sOptions = sOptions & vbCrLf
    recursive = sOptions	
end function

function getExt(sFile)'ffilter
	dim sExt
	sExt=""
	for each Item In split(sFile,".")
		sExt = Item
	next
	getExt=sExt
end function

function writeFileSelections()
'response.write currFolder
'response.end 
	set oFolder = oFSO.GetFolder(currFolder)
	set oFiles = oFolder.files
	nIndex=0
	
	bFileFound=false

	Response.Write "<div style='overflow:auto;height:300px;width:100%;margin-top:3px;margin-bottom:2px;'>" & VbCrLf
	Response.Write "<table border=0 cellpadding=2 cellspacing=0 width=100% height=100% >" & VbCrLf
	'Response.Write "<tr><td colspan=4 width='100%'><b>"&getTxt("Files")&"</b></td></tr>"
	sColor = "#e7e7e7"
	for each oFile in oFiles
		
		'ffilter ~~~~~~~~~~
		bDisplay=false
		sExt=getExt(oFile.path)
		if ffilter="flash" then
			if LCase(sExt)="swf" then bDisplay=true
		elseif ffilter="media" then
			if LCase(sExt)="avi" or LCase(sExt)="wmv" or LCase(sExt)="mpg" or _
			   LCase(sExt)="mpeg" or LCase(sExt)="wav" or LCase(sExt)="wma" or _
			   LCase(sExt)="mid" or LCase(sExt)="mp3" then bDisplay=true
		elseif ffilter="image" then
			if LCase(sExt)="gif" or LCase(sExt)="jpg" or LCase(sExt)="jpeg" or LCase(sExt)="png" then bDisplay=true
		else 'all
			bDisplay=true
		end if
		'~~~~~~~~~~~~~~~~~~
	
		if bDisplay then
		
			bFileFound=true
			
			nIndex=nIndex+1
			
			'server.MapPath(sBaseFolder)
			'oFSO.GetFolder(server.MapPath(sBaseFolder))
			' => sama. yg satu .. yg lain ..
			'response.Write oFile.path & "  -  " & oFSO.GetFolder(server.MapPath(sBaseFolder)) & "<br />"
			if InStr(1,oFile.path,oFSO.GetFolder(server.MapPath(arrBaseFolder(0)))) <> 0 then
				sBaseFolder_PhysicalPath = oFolder_base0.path 'folder terakhir tdk diikuti dgn / shg sFile_VirtualPath_UnderBaseFolder diawali dgn /
				sBaseFolder = arrBaseFolder(0)
			end if
			if arrBaseFolder(1)<>"" then'NEW 2.4
				if InStr(1,oFile.path,oFSO.GetFolder(server.MapPath(arrBaseFolder(1)))) <> 0 then
					sBaseFolder_PhysicalPath = oFolder_base1.path 'folder terakhir tdk diikuti dgn / shg sFile_VirtualPath_UnderBaseFolder diawali dgn /
					sBaseFolder = arrBaseFolder(1)
				end if
			end if
			if arrBaseFolder(2)<>"" then'NEW 2.4
				if InStr(1,oFile.path,oFSO.GetFolder(server.MapPath(arrBaseFolder(2)))) <> 0 then
					sBaseFolder_PhysicalPath = oFolder_base2.path 'folder terakhir tdk diikuti dgn / shg sFile_VirtualPath_UnderBaseFolder diawali dgn /
					sBaseFolder = arrBaseFolder(2)
				end if
			end if
			if arrBaseFolder(3)<>"" then'NEW 2.4
				if InStr(1,oFile.path,oFSO.GetFolder(server.MapPath(arrBaseFolder(3)))) <> 0 then
					sBaseFolder_PhysicalPath = oFolder_base3.path 'folder terakhir tdk diikuti dgn / shg sFile_VirtualPath_UnderBaseFolder diawali dgn /
					sBaseFolder = arrBaseFolder(3)
				end if
			end if

			sFile_PhysicalPath = oFile.path

			response.write "<!--Used for allocating problem:-->" & VbCrLf
			response.write "<!--"&sFile_PhysicalPath&"-->" & VbCrLf
			response.write "<!--"&sBaseFolder_PhysicalPath&"-->" & VbCrLf
			
			sBaseFolder_PhysicalPath=LCase(sBaseFolder_PhysicalPath)'NEW 2.4
			sFile_PhysicalPath=LCase(sFile_PhysicalPath)'NEW 2.4
			
			sFile_VirtualPath_UnderBaseFolder = replace(replace(sFile_PhysicalPath,sBaseFolder_PhysicalPath,""),"\","/") 'physical to virtual, ex: "", "/gallery01", "/gallery02"
			
			response.write "<!--"&sFile_VirtualPath_UnderBaseFolder&"-->" & VbCrLf
			
			if mid(sFile_VirtualPath_UnderBaseFolder,1,1)="/"  then
				sFile_VirtualPath_UnderBaseFolder=mid(sFile_VirtualPath_UnderBaseFolder,2)
			end if		
			
			response.write "<!--"&sFile_VirtualPath_UnderBaseFolder&"-->" & VbCrLf
			
			sFile_VirtualPath = CStr(sBaseFolder & sFile_VirtualPath_UnderBaseFolder)
			response.write "<!--"&sBaseFolder&"-->" & VbCrLf
			response.write "<!--"&sFile_VirtualPath&"-->" & VbCrLf

			if sColor = "#EFEFF5" then
				sColor = ""
			else
				sColor = "#EFEFF5"
			end if

			'icons
			sIcon="ico_unknown.gif"
			If LCase(sExt)="asp" then sIcon="ico_asp.gif"
			If LCase(sExt)="bmp" then sIcon="ico_bmp.gif"
			If LCase(sExt)="css" then sIcon="ico_css.gif"
			If LCase(sExt)="doc" then sIcon="ico_doc.gif"
			If LCase(sExt)="exe" then sIcon="ico_exe.gif"
			If LCase(sExt)="gif" then sIcon="ico_gif.gif"
			If LCase(sExt)="htm" then sIcon="ico_htm.gif"
			If LCase(sExt)="html" then sIcon="ico_htm.gif"
			If LCase(sExt)="jpg" then sIcon="ico_jpg.gif"
			If LCase(sExt)="jpeg" then sIcon="ico_jpg.gif"
			If LCase(sExt)="js"	 then sIcon="ico_js.gif"
			If LCase(sExt)="mdb" then sIcon="ico_mdb.gif"
			If LCase(sExt)="mov" then sIcon="ico_mov.gif"
			If LCase(sExt)="mp3" then sIcon="ico_mp3.gif"
			If LCase(sExt)="pdf" then sIcon="ico_pdf.gif"
			If LCase(sExt)="png" then sIcon="ico_png.gif"
			If LCase(sExt)="ppt" then sIcon="ico_ppt.gif"
			If LCase(sExt)="mid" then sIcon="ico_sound.gif"
			If LCase(sExt)="wav" then sIcon="ico_sound.gif"
			If LCase(sExt)="wma" then sIcon="ico_sound.gif"
			If LCase(sExt)="swf" then sIcon="ico_swf.gif"
			If LCase(sExt)="txt" then sIcon="ico_txt.gif"
			If LCase(sExt)="vbs" then sIcon="ico_vbs.gif"
			If LCase(sExt)="avi" then sIcon="ico_video.gif"
			If LCase(sExt)="wmv" then sIcon="ico_video.gif"	
			If LCase(sExt)="mpeg" then sIcon="ico_video.gif"					
			If LCase(sExt)="mpg" then sIcon="ico_video.gif"
			If LCase(sExt)="xls" then sIcon="ico_xls.gif"
			If LCase(sExt)="zip" then sIcon="ico_zip.gif"

			if(LCase(oFile.name)=LCase(sUploadedFile))then
				sColorResult="yellow"
				iSelected=nIndex
			else
				sColorResult=sColor
			end if

			Response.Write "<tr style='background:" & sColorResult & "'>" & VbCrLf & _
				"<td><img src='images/"&sIcon&"'></td>" & VbCrLf & _
				"<td valign=top width=100% ><u id=""idFile"&nIndex&""" style='cursor:pointer;' onclick=""selectFile(" & nIndex & ")"">" & oFile.name & "</u>&nbsp;&nbsp;<img style='cursor:pointer;' onclick=""downloadFile(" & nIndex & ")"" src='download.gif'></td>" & VbCrLf & _
				"<input type=hidden name=inpFile" & nIndex & " id=inpFile" & nIndex & " value=""" & sFile_VirtualPath & """>" & VbCrLf & _
				"<td valign=top align=right nowrap>" & FormatNumber(oFile.size/1000,1) & " kb&nbsp;</td>" & VbCrLf & _			
				"<td valign=top nowrap><a onclick=""deleteFile(" & nIndex & ")"" style=""font-size:10px;cursor:pointer;color:crimson;" & sFolderAdmin & """>"
			
			if not bWriteFolderAdmin then
				Response.Write "del" & VbCrLf
			end if
			
			Response.Write "</a></td></tr>" & VbCrLf	
		end if
	next
	if bFileFound=false then
		Response.Write "<tr><td colspan=4 height=100% align=center><script>document.write(getTxt('Empty...'))</script></td></tr></table></div>"
	else
		Response.Write "<tr><td colspan=4 height=100% ></td></tr></table></div>"
	end if

	Response.Write "<input type=hidden name=inpUploadedFile id=inpUploadedFile value='" & iSelected & "'>"
	Response.Write "<input type=hidden name=inpNumOfFiles id=inpNumOfFiles value='" & nIndex & "'>"
end function
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text-html; charset=Windows-1252">
<link href="style.css" rel="stylesheet" type="text/css">
<%
sLang="english"
if Len(Cstr(Request.QueryString("lang")))<>0 then
	sLang=Request.QueryString("lang")
end if
%>
<script type="text/javascript">
	var sLang="<%=sLang%>";
	document.write("<scr"+"ipt src='language/"+sLang+"/asset.js'></scr"+"ipt>");
</script>
<script type="text/javascript">writeTitle()</script>
<script type="text/javascript">
/*Used for allocating problem:*/
/*<%response.write arrBaseFolder(0)%>*/
/*<%response.write arrBaseFolder(1)%>*/
/*<%response.write arrBaseFolder(2)%>*/
/*<%response.write arrBaseFolder(3)%>*/
var bReturnAbsolute=<%if bReturnAbsolute then response.write "true" else response.write "false"%>;
var activeModalWin;

function getAction(isUpload)//NEW 2.4
	{
	//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	//Clean previous ffilter=...
	sQueryString=window.location.search.substring(1)
	sQueryString=sQueryString.replace(/&upload=Y/,"")//NEW 2.4

	sQueryString=sQueryString.replace(/ffilter=media/,"")
	sQueryString=sQueryString.replace(/ffilter=image/,"")
	sQueryString=sQueryString.replace(/ffilter=flash/,"")
	sQueryString=sQueryString.replace(/ffilter=/,"")
	if(sQueryString.substring(sQueryString.length-1)=="&")
		sQueryString=sQueryString.substring(0,sQueryString.length-1)

	if(sQueryString.indexOf("=")==-1) 
		{//no querystring
		sAction="assetmanagerIF.asp?ffilter="+document.getElementById("selFilter").value;
		}
	else 
		{
		sAction="assetmanagerIF.asp?"+sQueryString+"&ffilter="+document.getElementById("selFilter").value
		}
	
	if(isUpload) sAction+="&upload=Y";//NEW 2.4
	//~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	return sAction;
	}

function applyFilter()//ffilter
	{
	var Form1 = document.forms.Form1;
	
	Form1.elements.inpCurrFolder.value=document.getElementById("selCurrFolder").value;
	Form1.elements.inpFileToDelete.value="";

	Form1.action=getAction()
	Form1.submit()
	}
function refreshAfterDelete(sDestination)
	{
	var Form1 = document.forms.Form1;
	
	Form1.elements.inpCurrFolder.value=sDestination;
	Form1.elements.inpFileToDelete.value="";
	
	Form1.action=getAction()
	Form1.submit();
	}
function changeFolder()
	{
	var Form1 = document.forms.Form1;
	
	Form1.elements.inpCurrFolder.value=document.getElementById("selCurrFolder").value;
	Form1.elements.inpFileToDelete.value="";
	
	Form1.action=getAction()
	Form1.submit();
	}
function upload()
	{
	var Form2 = document.forms.Form2;
	
	if(Form2.elements.File1.value == "")return;

	var sFile=Form2.elements.File1.value.substring(Form2.elements.File1.value.lastIndexOf("\\")+1);
	for(var i=0;i<document.getElementById("inpNumOfFiles").value;i++)
		{
		if(sFile==document.getElementById("idFile"+(i*1+1)).innerHTML)
			{
			if(confirm(getTxt("File already exists. Do you want to replace it?"))!=true)return;
			}
		}

	Form2.elements.inpCurrFolder2.value=document.getElementById("selCurrFolder").value;
	document.getElementById("idUploadStatus").innerHTML=getTxt("Uploading...")
		
	Form2.action=getAction(true);//NEW 2.4
	Form2.submit();
	}
function modalDialogShow(url,width,height)//moz
    {
    var left = screen.availWidth/2 - width/2;
    var top = screen.availHeight/2 - height/2;
    activeModalWin = window.open(url, "", "width="+width+"px,height="+height+",left="+left+",top="+top);
    window.onfocus = function(){if (activeModalWin.closed == false){activeModalWin.focus();};};
    }
function newFolder()
	{
	if(navigator.appName.indexOf('Microsoft')!=-1)
		window.showModalDialog("foldernew.asp",window,"dialogWidth:250px;dialogHeight:192px;edge:Raised;center:Yes;help:No;resizable:No;");
	else
		modalDialogShow("foldernew.asp", 250, 150);
	}
function deleteFolder()
	{
	var selCurrFolder = document.getElementById("selCurrFolder");

	if(selCurrFolder.value.toLowerCase()==document.getElementById("inpAssetBaseFolder0").value.toLowerCase() ||
	selCurrFolder.value.toLowerCase()==document.getElementById("inpAssetBaseFolder1").value.toLowerCase() ||
	selCurrFolder.value.toLowerCase()==document.getElementById("inpAssetBaseFolder2").value.toLowerCase() ||
	selCurrFolder.value.toLowerCase()==document.getElementById("inpAssetBaseFolder3").value.toLowerCase())
		{
		alert(getTxt("Cannot delete Asset Base Folder."));
		return;
		}
	
	if(navigator.appName.indexOf('Microsoft')!=-1)
		window.showModalDialog("folderdel.asp",window,"dialogWidth:250px;dialogHeight:192px;edge:Raised;center:Yes;help:No;resizable:No;");
	else
		modalDialogShow("folderdel.asp", 250, 150);
	}
function downloadFile(index)
	{
	sFile_RelativePath = document.getElementById("inpFile"+index).value;
	window.open(sFile_RelativePath)
	}
function selectFile(index)
	{
	sFile_RelativePath = document.getElementById("inpFile"+index).value;

	//This will make an Absolute Path
	if(bReturnAbsolute)
		{
		sFile_RelativePath = window.location.protocol + "//" + window.location.host.replace(/:80/,"") + sFile_RelativePath
		//Ini input dr yg pernah pake port:
		//sFile_RelativePath = window.location.protocol + "//" + window.location.host.replace(/:80/,"") + "/" + sFile_RelativePath.replace(/\.\.\//g,"")
		}
	
	document.getElementById("inpSource").value=sFile_RelativePath;
	
	var arrTmp = sFile_RelativePath.split(".");
	var sFile_Extension = arrTmp[arrTmp.length-1]	
	var sHTML="";
	
	//Image
	if(sFile_Extension.toUpperCase()=="JPG" || sFile_Extension.toUpperCase()=="JPEG")
		{
		<%if cBool(Application("QS_ASPX")) then%>
		sHTML = "<img src=\"" + "<%=Application("QS_CMS_C_DIRECTORY_QUICKERSITE")%>/showThumb.aspx?FSR=0&amp;maxSize=287&amp;img=" + sFile_RelativePath + "\" >"
		<%else%>
		sHTML = "<img src=\"" + sFile_RelativePath + "\" >"
		<%end if%>
		}
	//GIF or PNG
	else if(sFile_Extension.toUpperCase()=="PNG" || sFile_Extension.toUpperCase()=="GIF")
		sHTML ="<img style='width:95%' src=\"" + sFile_RelativePath + "\" >"
	
	//SWF
	else if(sFile_Extension.toUpperCase()=="SWF")
		{
		sHTML = "<object "+
			"classid='clsid:D27CDB6E-AE6D-11cf-96B8-444553540000' " +
			"width='100%' "+
			"height='100%' " +
			"codebase='http://active.macromedia.com/flash6/cabs/swflash.cab#version=6.0.0.0'>"+
			"	<param name=movie value='"+sFile_RelativePath+"'>" +
			"	<param name=quality value='high'>" +
			"	<embed src='"+sFile_RelativePath+"' " +
			"		width='100%' " +
			"		height='100%' " +
			"		quality='high' " +
			"		pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash'>" +
			"	</embed>"+
			"</object>";
		}
	//Video
	else if(sFile_Extension.toUpperCase()=="WMV"||sFile_Extension.toUpperCase()=="AVI"||sFile_Extension.toUpperCase()=="MPG")
		{
		sHTML = "<embed src='"+sFile_RelativePath+"' hidden=false autostart='true' type='video/avi' loop='true'></embed>";
		}
	//Sound
	else if(sFile_Extension.toUpperCase()=="WMA"||sFile_Extension.toUpperCase()=="WAV"||sFile_Extension.toUpperCase()=="MID")
		{
		sHTML = "<embed src='"+sFile_RelativePath+"' hidden=false autostart='true' type='audio/wav' loop='true'></embed>";
		}

    <%
        if not cBool(Session(Application("QS_CMS_iCustomerID") & "isAUTHENTICATEDSecondAdmin")) and request.querystring("CKEditorFuncNum")=""  then     
    %>

    else if(sFile_Extension.toUpperCase()=="HTM"||sFile_Extension.toUpperCase()=="HTML"||sFile_Extension.toUpperCase()=="CSS"||sFile_Extension.toUpperCase()=="TXT"||sFile_Extension.toUpperCase()=="JS")
        {
    window.open('../bs_editcode.asp?path=' + sFile_RelativePath, '', 'resizable=yes,status=no,location=no,toolbar=no,menubar=no,fullscreen=no,scrollbars=yes,dependent=no,width=950,left=10,height=620,top=10'); return false;
        }
    <%
        end if
    %>
	//Files (Hyperlinks)
	else
		{	
		sHTML = "<br /><br /><br /><br /><br /><br />Not Available"
		}
		
	document.getElementById("idPreview").innerHTML = sHTML;
	}
function deleteFile(index)
	{
	if (confirm(getTxt("Delete this file ?")) == true) 
		{	
		sFile_RelativePath = document.getElementById("inpFile"+index).value;

		var Form1 = document.getElementById("Form1");
		Form1.elements.inpCurrFolder.value=document.getElementById("selCurrFolder").value;
		Form1.elements.inpFileToDelete.value=sFile_RelativePath;

		Form1.action=getAction()
		Form1.submit();
		}
	}
bOk=false;
function doOk()
	{
	if(navigator.appName.indexOf('Microsoft')!=-1)
		window.returnValue=inpSource.value;
	else
		window.opener.setAssetValue(document.getElementById("inpSource").value);
	bOk=true;
	self.close();
	}
function doUnload()
	{
	if(navigator.appName.indexOf('Microsoft')!=-1)
		if(!bOk)window.returnValue="";
	else
		if(!bOk)window.opener.setAssetValue("");
	}
</script>

    <%
        dim CKEditorFuncNum
        if Request.querystring("CKEditorFuncNum")<>"" then
            CKEditorFuncNum=Request.querystring("CKEditorFuncNum")
        else
            CKEditorFuncNum="0"
        end if

    %>

<script type="text/javascript">
			function setU(surl) {window.opener.CKEDITOR.tools.callFunction(<%=CKEditorFuncNum%>,surl)};
			</script>	
</head>
<body onunload="doUnload()" onload="loadTxt();this.focus();if(document.getElementById('inpUploadedFile').value!='')selectFile(document.getElementById('inpUploadedFile').value);" style="overflow:hidden;margin:0px;">
<%
'response.write "CKEditorFuncNum: " & 
%>
<input type="hidden" name="inpAssetBaseFolder0" id="inpAssetBaseFolder0" value="<%=oFolder_base0.path%>">
<input type="hidden" name="inpAssetBaseFolder1" id="inpAssetBaseFolder1" value="<%if arrBaseFolder(1)<>"" then response.Write(oFolder_base1.path)%>">
<input type="hidden" name="inpAssetBaseFolder2" id="inpAssetBaseFolder2" value="<%if arrBaseFolder(2)<>"" then response.Write(oFolder_base2.path)%>">
<input type="hidden" name="inpAssetBaseFolder3" id="inpAssetBaseFolder3" value="<%if arrBaseFolder(3)<>"" then response.Write(oFolder_base3.path)%>">

<table width="100%" height="100%" align=center style="" cellpadding=0 cellspacing=0 border=0 >
<tr>
<td valign=top style="background:url('bg.gif') no-repeat right bottom;padding-top:5px;padding-left:5px;padding-right:5px;padding-bottom:0px;">
		<!--ffilter-->
		<form method=post name="Form1" id="Form1">
				<input type="hidden" name="inpFileToDelete" ID="Hidden1">
				<input type="hidden" name="inpCurrFolder" ID="Hidden2">
		</form>

		<table width=100% border="0">
		<tr>
		<td>
				<table cellpadding="2" cellspacing="2" border="0">
				<tr>
				<td valign=center nowrap><%writeFolderSelections()%>&nbsp;</td>
				<td nowrap>
					<span onclick="newFolder()" style="cursor:pointer;"><u><span name="txtLang" id="txtLang" <%=sFolderAdmin%>>New&nbsp;Folder</span></u></span>&nbsp;
					<span onclick="deleteFolder()" style="cursor:pointer;"><u><span name="txtLang" id="txtLang" <%=sFolderAdmin%>>Del&nbsp;Folder</span></u></span>
				</td>
				<td  width=100% align="right">
				<%				
				'ffilter~~~~~~~~~
					sHTMLFilter = "<select name=selFilter id=selFilter onchange='applyFilter()' class='inpSel'>"'ffilter
					sAll=""
					sMedia=""
					sImage=""
					sFlash=""	
					if ffilter="" then sAll="selected"
					if ffilter="media" then sMedia="selected"
					if ffilter="image" then sImage="selected"
					if ffilter="flash" then sFlash="selected"
					sHTMLFilter = sHTMLFilter & "	<option name=optLang id=optLang value='' "&sAll&"></option>"
					sHTMLFilter = sHTMLFilter & "	<option name=optLang id=optLang value='media' "&sMedia&"></option>"
					sHTMLFilter = sHTMLFilter & "	<option name=optLang id=optLang value='image' "&sImage&"></option>"
					sHTMLFilter = sHTMLFilter & "	<option name=optLang id=optLang value='flash' "&sFlash&"></option>"
					sHTMLFilter = sHTMLFilter & "</select>"
					Response.Write sHTMLFilter
				'~~~~~~~~~
				%>
				</td>
				</tr>
				</table>
		</td>
		</tr>		
		<tr>
		<td valign=top align="center">
		
				<table width=100% cellpadding=0 cellspacing=0>
				<tr>
				<td>
					<div id="idPreview" style="text-align:center;overflow:auto;width:297;height:267;border:#d7d7d7 5px solid;border-bottom:#d7d7d7 3px solid;background:#ffffff;margin-right:2;"></div>
					<div align=center><form name="form3"><input type="text" id="inpSource" name="inpSource" style="border:#cfcfcf 1px solid;width:295" class="inpTxt"></form></div>
				</td>
				<td valign=top width=100%>			
					<%writeFileSelections()%>
				</td>
				</tr>
				<%
				if showInsertButton then
				%>
				<tr><td><input type=button  onclick="javascript:setU(form3.inpSource.value);window.close()" value=Insert /></td></tr>
				<%
				end if
				if showSelectButton then
				%>
				<tr><td><input type=button  onclick="javascript:window.parent.document.mainform.sUrlRRSImage.value=form3.inpSource.value;window.parent.$.fn.colorbox.close();" value=Select /></td></tr>
				<%
				end if
				%>
				</table>		
					
		</td>
		</tr>
		<tr>
		<td>
				<div <%=sFolderAdmin%>>
					<div style="height:12">
					<font color=red><%=sMsg%></font>
					<span style="font-weight:bold" id=idUploadStatus></span>
					</div>					
					<form enctype="multipart/form-data" method="post" name="Form2" id="Form2">
					<input type="hidden" name="inpCurrFolder2" ID="inpCurrFolder2">
					<!--ffilter-->
					<input type="hidden" name="inpFilter" ID="inpFilter" value="<%=ffilter%>">
					<span name="txtLang" id="txtLang">Upload File</span>: <input type="file" id="File1" name="File1" class="inpTxt">&nbsp;
					<input name="btnUpload" id="btnUpload" type="button" value="upload" onclick="upload()" class="inpBtn" onmouseover="this.className='inpBtnOver';" onmouseout="this.className='inpBtnOut'">
					<br>
					<%
					dim fff,sss
					sss=lcase(server.mappath(Application("QS_CMS_C_VIRT_DIR") & Application("QS_CMS_userfiles")))
					fff=lcase(currfolder)
					
					dim prepareLink
					prepareLink=replace(fff,sss,"",1,-1,1)
					prepareLink=replace(prepareLink,"\","/",1,-1,1)
					
					'prepareLink=prepareLink
					'if prepareLink<>"" then
					if trim(Request.querystring("CKEditorFuncNum"))="" then
					%>
					
					<%
					end if
					%>
					</form>
				</div>
		</td>
		</tr>
		</table>

</td>
</tr>


<tr>
<td class="dialogFooter" style="height:10px;padding-right:15px;" align=right valign=middle>
	<table cellpadding=0 cellspacing=0 ID="Table2">
	<tr>
	<td>
	<input name="btnOk" id="btnOk" type="button" value=" ok " onclick="doOk()" class="inpBtn" style="visibility:hidden" onmouseover="this.className='inpBtnOver';" onmouseout="this.className='inpBtnOut'">
	</td>
	</tr>
	</table>
</td>
</tr>
</table>

</body>
</html>
