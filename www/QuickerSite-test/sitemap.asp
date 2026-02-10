<!-- #include file="asp/begin.asp"-->


<%'this script flushes a XML sitemap. You can validate/submit this sitemap on 
'sites like http://www.xml-sitemaps.com/validate-xml-sitemap.html
'All you have to do is provide the link http://www.yourQsite.com/sitemap.asp
'Next, notify the different Search Engines of your sitemap.
dim sitemap
sitemap="<?xml version=""1.0"" encoding=""UTF-8""?>" & vbcrlf
sitemap=sitemap & "<urlset" & vbcrlf
sitemap=sitemap & "xmlns=""http://www.sitemaps.org/schemas/sitemap/0.9""" & vbcrlf
sitemap=sitemap & "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""" & vbcrlf
sitemap=sitemap & "xsi:schemaLocation=""http://www.sitemaps.org/schemas/sitemap/0.9" & vbcrlf
sitemap=sitemap & "http://www.sitemaps.org/schemas/sitemap/0.9/sitemap.xsd"">"& vbcrlf
'reuse the search feature to get a list of all pages
dim getSiteMap
set getSiteMap=new cls_search
getSiteMap.includeURL=false
getSiteMap.allowEmptyString=true
getSiteMap.includeListItems=true
getSiteMap.value=""
dim getResults
set getResults=getSiteMap.results
'loop through the resultset
dim pageKey
for each pageKey in getResults

'theme?
if convertGetal(getResults(pageKey).iThemeID)<>0 then

dim posts
set posts=db.execute("select iId from tblPost where (iPostID is null or iPostID=0) and iThemeID="& getResults(pageKey).iThemeID)

while not posts.eof
sitemap=sitemap & "<url>"& vbcrlf &"<loc>"& customer.sQSUrl & "/default.asp?iId="& encrypt(pageKey) &"&amp;iPostID="& encrypt(posts("iId")) & "</loc>" &  vbcrlf & "</url>"& vbcrlf
posts.movenext
wend 

end if


sitemap=sitemap & "<url>"& vbcrlf
'listitem?
if convertGetal(getResults(pageKey).iListPageID)<>0 then
sitemap=sitemap & "<loc>"& customer.sQSUrl & "/default.asp?iId="& encrypt(getResults(pageKey).iListPageID) &"&amp;item="& encrypt(pagekey) &"</loc>"& vbcrlf
'user friendly url?
elseif customer.bUserFriendlyURL and not isLeeg(getResults(pageKey).sUserFriendlyURL) then
sitemap=sitemap & "<loc>"& customer.sQSUrl & "/"& getResults(pageKey).sUserFriendlyURL &"</loc>"& vbcrlf
'any other page
else
sitemap=sitemap & "<loc>"& customer.sQSUrl & "/default.asp?iId="& encrypt(pageKey) &"</loc>"& vbcrlf
end if
sitemap=sitemap & "</url>"& vbcrlf
next
sitemap=sitemap & "</urlset>"& vbcrlf
set getResults=nothing
set getSiteMap=nothing
'flush the sitemap
Response.Write sitemap
cleanUPASP()%>
