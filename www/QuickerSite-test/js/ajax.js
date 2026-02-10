var xmlhttp;
var qs_div;


function getPage(sCode)
{
xmlhttp=GetXmlHttpObject();
if (xmlhttp==null)
  {
  alert ("Browser does not support HTTP Request");
  return;
  }
var url="test.asp";
url=url+"?sCode="+sCode;
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