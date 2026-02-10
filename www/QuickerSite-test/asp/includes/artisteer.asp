
<%dim qs_artisteer_parentlink
qs_artisteer_parentlink="<span class=""l""></span><span class=""r""></span><span class=""t"">{TITLE}</span>"
function getParentLinkForArtisteer(sString)
getParentLinkForArtisteer=replace(qs_artisteer_parentlink,"{TITLE}",sString,1,-1,1)
end function%>
