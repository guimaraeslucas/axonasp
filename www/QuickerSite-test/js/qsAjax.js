var xmlhttp;
var qs_div;
var mode;

function getVote(int,poll)
{
xmlhttp=GetXmlHttpObject();
var url="default.asp?pageAction=";
if (mode=='view')
{
url=url+"voteshowresults";
}
else if (mode=='vote')
{
url=url+"voteshowballot";
}
else
{
url=url+"vote";
url=url+"&vote="+int;
}
url=url+"&iPolliD="+poll;
url=url+"&sid="+Math.random();
xmlhttp.onreadystatechange=stateChanged;
xmlhttp.open("GET",url,true);
xmlhttp.send(null);
}

function getSub(sName,sEmail,iCatID,secCode)
{
xmlhttp=GetXmlHttpObject();
var url="default.asp?pageAction=";
if (mode=='fb')
{
url=url+"nlcatshowfb";
url=url+"&sName="+encodeURIComponent(sName);
url=url+"&sEmail="+encodeURIComponent(sEmail);
}
else
{
url=url+"nlcatshowform";
}
url=url+"&iNewsletterCategoryID="+iCatID;
url=url+"&cQSSECNL="+secCode;
url=url+"&sid="+Math.random();
xmlhttp.onreadystatechange=stateChanged;
xmlhttp.open("GET",url,true);
xmlhttp.send(null);
}


function stateChanged()
{
  if (xmlhttp.readyState==4)
  {
  document.getElementById('div'+qs_div).innerHTML=xmlhttp.responseText;
  }
}

function GetXmlHttpObject()
{
var objXMLHttp=null;
if (window.XMLHttpRequest)
  {
  objXMLHttp=new XMLHttpRequest();
  }
else if (window.ActiveXObject)
  {
  objXMLHttp=new ActiveXObject("Microsoft.XMLHTTP");
  }
return objXMLHttp;
}
