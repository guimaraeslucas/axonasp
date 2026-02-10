						
  var expDays= 365;
  var CookieInfoStr='';

  var expdate = new Date();
  var olddate = new Date();
  expdate.setTime (expdate.getTime() + (expDays*24*60*60*1000));
  olddate.setTime (expdate.getTime());
		
function StoreDocumentLink(){
    document.write ('<a href="javascript:addCookieArray(\'' + document.title + '\',\'' + location.href + '\');"><img src="fixedImages/public/bookmark.gif" border="0"></a>');
}	
function ClearStoredDocumentsLink(){
    document.write ('+ <a href="javascript:del();">Wis favorieten</a><br>');
}			
function ShowStoredDocumentsLink(){
  var i = 0;
  //document.write ('<ul>');
  while (getCookie('names' + i) != null) {
   document.write ('<img src="fixedImages/public/redArrow.gif" width="7" border="0" align="absmiddle" hspace="2">');  
   document.write ('<a href="' + getCookie('urls' + i) + '">' + getCookie('names' + i) + '</span></a>' );
   document.write ('<br>');
   i++; 
  }  
  if (i==0){
   document.write ('<img src="fixedImages/public/redArrow.gif" width="7" border="0" align="absmiddle" hspace="2">');  
   document.write ('Er zijn momenteel geen favorieten aanwezig');
   document.write ('<br>');
  }
  //document.write ('</ul>');

}


function getCookie (name) {
var dcookie = document.cookie; 
var cname = name + "=";
var clen = dcookie.length;
var cbegin = 0;
  while (cbegin < clen) {
  var vbegin = cbegin + cname.length;
  if (dcookie.substring(cbegin, vbegin) == cname) { 
    var vend = dcookie.indexOf (";", vbegin);
    if (vend == -1) vend = clen;
    return unescape(dcookie.substring(vbegin, vend));
  }
  cbegin = dcookie.indexOf(" ", cbegin) + 1;
  if (cbegin == 0) break;
  }
return null;
}
function setCookie (name, value, expires) {
if (!expires) expires = new Date();
document.cookie=name+"="+escape (value)+"; expires="+expires.toGMTString()+"; path=/";
}
function setCookieArray(name){
this.length = setCookieArray.arguments.length - 1;
  for (var i = 0; i < this.length; i++) {
  this[i + 1] = setCookieArray.arguments[i + 1];
  setCookie (name + i, this[i + 1], expdate);
  }  
}
function addCookieArray(name, val){
  var i = 0;
  var found=false;
    
  while (getCookie('names' + i) != null) { 
    if (getCookie('names' + i) == name) found=true;
    i++;	
  }    
  
  if (!found) {
  setCookie ('names' + i, name, expdate);
  setCookie ('urls' + i, val, expdate);
  location.reload();
  }
  else {
  //alert('This page has been added to your stored documents. ');
  }
  
}

function getCookieArray(name){
var i = 0;
  while (getCookie(name + i) != null) {
  this[i + 1] = getCookie(name + i);
  i++; this.length = i;
  }		
}
function del() {
var i = 0;
if (window.confirm('Alle favorieten wissen?  ')) {

  while (getCookie('names' + i) != null) {
  setCookie ('names' + i, '', olddate);
  setCookie ('urls' + i, '', olddate);
  i++; 
  }  
  i++; 
  setCookie ('names' + i, '', olddate);
  setCookie ('urls' + i, '', olddate);   
  //alert('Stored documents have been cleared. ');
  location.reload();
}
}
