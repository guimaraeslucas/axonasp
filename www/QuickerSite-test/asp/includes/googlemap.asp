
<%'################################################################################################
'######
'###### (c) Pieter Cooreman - Free Google Map VBScript Class
'###### You are free to use this script - but you have to leave the copyright statement intact
'###### Before you plan to use this script in any commercial application, make sure to read
'###### the Google Map Terms on http://code.google.com/intl/nl-BE/apis/maps/terms.html
'######
'###### Submit bugs to info@quickersite.com. Thx!
'######
'################################################################################################
'###### 
'###### With this VBSript class you can show locations in a Google Map
'###### 
'###### You can:
'###### -set a custom width, height, title-tag, zoomlevel for the GoogleMap
'###### -set the default maptype - Most common: "G_NORMAL_MAP", "G_SATELLITE_MAP", "G_HYBRID_MAP"
'###### -add multipe addresses to one map
'###### -use a custom icon for each marker
'###### -use JavaScript/HTML to show when you click a marker
'###### -set some custom errorMessages in case something goes wrong
'###### -Check out both classes cls_googleMap and cls_address to read the details
'###### 
'################################################################################################
class cls_googleMap
public key, width, height, sTitle, addresses, zoomlevel, divID, errorMessage, noAddressSet, useDelay, maptype
public enableGoogleBar,GLargeMapControl,GMapTypeControl,enableScrollWheelZoom,GOverviewMapControl
Private Sub Class_Initialize
sTitle="Google Maps made easy for ASP/VBScript Developers - By Pieter Cooreman"
zoomlevel=11
width=500
height=400
divID="map_canvas"
useDelay=0 'number of ms to delay each address-search - 200 ms is very useful in case you have many addresses (>100)
errorMessage="This address was not found in Google Maps" 'leave empty in case you don't need this
noAddressSet="No address was set!" 'leave empty in case you don't need this
maptype="G_NORMAL_MAP" 'Most common: "G_NORMAL_MAP", "G_SATELLITE_MAP", "G_HYBRID_MAP"
enableGoogleBar=true 'show the Google Search Box bottom-left (true or false)
GLargeMapControl=true 'show the Large Map Control  (true or false)
GMapTypeControl=true 'show the 3 maptypes (true or false)
enableScrollWheelZoom=true
GOverviewMapControl=true 'show the zoom window bottom right
set addresses=server.CreateObject ("scripting.dictionary")
End Sub
Private Sub Class_Terminate
set addresses=nothing
End Sub
Public function addAddress(addressObject)
if trim(addressObject.searchString)<>"" then 'no empty addresses - Google doesn't like that
if not addresses.Exists (addressObject.searchString) then 'no double addresses - - Google doesn't like that either
addresses.Add addressObject.searchString,addressObject 'add the address to the addresses collection
end if
end if
end function
'the function showMap, generates the full html 
public Function showMap
showMap=showMap &"<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Strict//EN""" & vbcrlf & """http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd"">" & vbcrlf
showMap=showMap &"<html xmlns=""http://www.w3.org/1999/xhtml"" xmlns:v=""urn:schemas-microsoft-com:vml"">"& vbcrlf
showMap=showMap &"<head>"& vbcrlf
showMap=showMap &"<meta http-equiv=""content-type"" content=""text/html; charset=UTF-8"" />"& vbcrlf
showMap=showMap &"<title>" & sTitle & "</title>"& vbcrlf
showMap=showMap &"<script src=""http://maps.google.com/maps?file=api&amp;v=2&amp;key=" & key & """ type=""text/javascript""></script>"& vbcrlf
showMap=showMap &"</head>"& vbcrlf
showMap=showMap &"<body style='margin:0' onload=""initialize()"" onunload=""GUnload()"">"& vbcrlf
showMap=showMap &"<script type=""text/javascript"">"& vbcrlf
showMap=showMap &"var map = null;"& vbcrlf
showMap=showMap &"var geocoder = null;"& vbcrlf
showMap=showMap &"geocoder = new GClientGeocoder();"& vbcrlf
showMap=showMap &"function initialize() {"& vbcrlf
showMap=showMap &"if (GBrowserIsCompatible()) {"& vbcrlf
showMap=showMap &"map = new GMap2(document.getElementById(""map_canvas""),{size:new GSize(" & width & "," & height & ")});"& vbcrlf
if enableGoogleBar then
showMap=showMap &"map.enableGoogleBar();"& vbcrlf
end if
if enableScrollWheelZoom then
showMap=showMap &"map.enableScrollWheelZoom();"& vbcrlf
end if
if GLargeMapControl then
showMap=showMap &"map.addControl(new GLargeMapControl());"& vbcrlf
end if
if GMapTypeControl then
showMap=showMap &"map.addControl(new GMapTypeControl());"& vbcrlf
end if
if GOverviewMapControl then
showMap=showMap &"var mini = new GOverviewMapControl(new GSize(100,100));" & vbcrlf
showMap=showMap &"map.addControl(new GMapTypeControl());" & vbcrlf
showMap=showMap &"map.addControl(mini);" & vbcrlf
end if
showMap=showMap &showAddress& vbcrlf
showMap=showMap &createMarker& vbcrlf
showMap=showMap &getJSArray& vbcrlf
showMap=showMap &"}" & vbcrlf
showMap=showMap &"}" & vbcrlf
showMap=showMap &"function isEmpty(inputStr){return !(inputStr&&inputStr.length)}"& vbcrlf
showMap=showMap &"</script>"& vbcrlf
showMap=showMap &"<div id=""" & divID & """ style=""width: " & width & "px; height: " & height & "px""></div>"& vbcrlf
showMap=showMap &"</body>"& vbcrlf
showMap=showMap &"</html>"	& vbcrlf
end function
private function getJSArray
if addresses.Count=0 then
if noAddressSet<>"" then getJSArray="alert('" & quotrepJS(noAddressSet) & "');" & vbcrlf
else
getJSArray="var AddressesArray = new Array (" & addresses.Count  & ");"& vbcrlf
'loop through addresses in VBScript
dim address,counter
counter=0
for each address in addresses
getJSArray=getJSArray & "AddressesArray [" & counter & "] = new Array(6);"
getJSArray=getJSArray & "AddressesArray [" & counter & "] [0]='"& quotRepJS(addresses(address).searchstring) &"';" & vbcrlf
if addresses(address).html<>"" then
if isNumeric(addresses(address).infoWindowWidth) then
getJSArray=getJSArray & "AddressesArray [" & counter & "] [1]='<div style=""padding:3;width:" & addresses(address).infoWindowWidth & "px"">"& quotRepJS(addresses(address).html) & "</div>';" & vbcrlf
else
getJSArray=getJSArray & "AddressesArray [" & counter & "] [1]='"& quotRepJS(addresses(address).html) & "';" & vbcrlf
end if
else
getJSArray=getJSArray & "AddressesArray [" & counter & "] [1]='';" & vbcrlf
end if
getJSArray=getJSArray & "AddressesArray [" & counter & "] [2]='"& quotRepJS(addresses(address).iconURL) & "';" & vbcrlf
getJSArray=getJSArray & "AddressesArray [" & counter & "] [3]='"& cBOOLcustom(addresses(address).openInfoWindowHtml) & "';" & vbcrlf
getJSArray=getJSArray & "AddressesArray [" & counter & "] [4]='"& quotRepJS(addresses(address).latitude) & "';" & vbcrlf
getJSArray=getJSArray & "AddressesArray [" & counter & "] [5]='"& quotRepJS(addresses(address).longitude) & "';" & vbcrlf
counter=counter+1
next
'loop through addresses in JavaScript, and call the JS function showAddress
getJSArray=getJSArray & "for (i=0;i<AddressesArray.length;i++){"& vbcrlf
getJSArray=getJSArray & "showAddress(AddressesArray[i][0],AddressesArray[i][1],AddressesArray[i][2],AddressesArray[i][3],AddressesArray[i][4],AddressesArray[i][5]);"& vbcrlf
getJSArray=getJSArray & "};" & vbcrlf
end if
end function
private function createMarker()
 createMarker="function createMarker(latlng,html,iconUrl,openHTML) {" & vbcrlf
 createMarker=createMarker&"if(!isEmpty(iconUrl)){" & vbcrlf
 createMarker=createMarker&"var myIcon = new GIcon(G_DEFAULT_ICON);" & vbcrlf
 createMarker=createMarker&"myIcon.image = iconUrl;" & vbcrlf
 createMarker=createMarker&"myIcon.iconSize = new GSize(32, 32);"& vbcrlf
 createMarker=createMarker&"var marker = new GMarker(latlng,{icon:myIcon});" & vbcrlf
 createMarker=createMarker&"}else{var marker = new GMarker(latlng);}" & vbcrlf
 createMarker=createMarker&"if(!isEmpty(html)){" & vbcrlf
 createMarker=createMarker&"marker.value = html;"& vbcrlf
 createMarker=createMarker&"GEvent.addListener(marker,""click"", function() {"& vbcrlf
 createMarker=createMarker&"map.openInfoWindowHtml(latlng, marker.value);"& vbcrlf
 createMarker=createMarker&"});}"& vbcrlf
 createMarker=createMarker&"return marker;}"& vbcrlf
 
end function
private function fixedSA
fixedSA=""
if useDelay>0 then
fixedSA=fixedSA & "var date = new Date();"& vbcrlf
fixedSA=fixedSA & "var curDate = null;"& vbcrlf
fixedSA=fixedSA & "do {curDate = new Date();}"& vbcrlf
fixedSA=fixedSA & "while(curDate-date < " & useDelay & ");"& vbcrlf
end if
fixedSA=fixedSA & "map.setCenter(point," & zoomlevel &"," & maptype & ");"& vbcrlf
fixedSA=fixedSA & "var marker=createMarker(point,html,iconUrl,openHTML);"& vbcrlf
            fixedSA=fixedSA & "map.addOverlay(marker);"& vbcrlf
            fixedSA=fixedSA & "if (openHTML=='True'){marker.openInfoWindowHtml(html)};" & vbcrlf
end function
private function showAddress
showAddress="function showAddress(searchstring,html,iconUrl,openHTML,lat,lon) {" & vbcrlf
'showAddress=showAddress & "return true;" & vbcrlf
showAddress=showAddress & "if(!isEmpty(lat)&&!isEmpty(lon)){var point=new GLatLng(lat,lon);"& vbcrlf
showAddress=showAddress & fixedSA            
showAddress=showAddress & "return true;};" & vbcrlf
showAddress=showAddress & "if (geocoder) {"& vbcrlf
showAddress=showAddress & "geocoder.getLatLng(searchstring,function(point) {" & vbcrlf
showAddress=showAddress & "if (!point) {"& vbcrlf
if errorMessage<>"" then showAddress=showAddress & "alert('" & quotRepJS(errorMessage) & "')"& vbcrlf
showAddress=showAddress & "} else {"& vbcrlf
showAddress=showAddress & fixedSA   
            showAddress=showAddress & "}});}}" & vbcrlf
            
end function
private function quotRepJS(sValue)
quotRepJS=replace(sValue,"\","\\",1,-1,1)
quotRepJS=replace(quotRepJS,"'","\'",1,-1,1)
quotRepJS=replace(quotRepJS,vbcrlf," ",1,-1,1)
end function
private function cBOOLcustom(value)
if isNull(value) then
cBOOLcustom="False"
elseif value=False then
cBOOLcustom="False"
else
cBOOLcustom="True"
end if
end function
'this function retrieves the coordinates of a location
public function getCoordinates	(q,byref lat,byref lon)
getCoordinates=false
if isNull(q) then q=""
if trim(q)<>"" then
on error resume next 
'get the codes from gooogle
dim oXMLHTTP
Set oXMLHTTP = server.createobject("MSXML2.ServerXMLHTTP")
oXMLHTTP.open "GET", "http://maps.google.com/maps/geo?q="& server.URLEncode(q) & "&output=CSV&key=" & key,false
oXMLHTTP.send
dim arrLL
arrLL=split(oXMLHTTP.responseText,",")
lat=arrLL(2)
lon=arrLL(3)
set oXMLHTTP=nothing
if clng(lat)<>0 and clng(lon)<>0 then getCoordinates=true
on error goto 0
end if
end function
end class
class cls_address
public street,zip,city,state,countrycode,country,html,iconURL,openInfoWindowHtml,Latitude,Longitude,infoWindowWidth
Private Sub Class_Initialize
openInfoWindowHtml=false
infoWindowWidth=270
Latitude=""
Longitude=""
end sub
public property get searchString
if isNull(street) then street=""
if isNull(zip) then zip=""
if isNull(city) then city=""
if isNull(state) then state=""
if isNull(countrycode) then countrycode=""
if isNull(country) then country=""
if isNull(html) then html=""
if isNull(iconURL) then iconURL=""
if isNull(Latitude) then Latitude=""
if isNull(Longitude) then Latitude=""
if cstr(Latitude)="0" then Latitude=""
if cstr(Longitude)="0" then Longitude=""
'When  Latitude and Longitude are set, they will be used, and not the regular address elements...
if trim(Latitude & Longitude)="" then
searchString=trim(trim(street) & " " & trim(zip) & " " & trim(city) & " " & trim(state) & " " & trim(countrycode) & " " & trim(country))
else
searchString=Longitude&Latitude
end if
end property
end class%>
