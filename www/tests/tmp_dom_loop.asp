<%
Response.ContentType = "text/plain"
On Error Resume Next
Dim xmlDOM, feeditems, Item, child
Set xmlDOM = Server.CreateObject("MSXML2.DOMDocument")
xmlDOM.async = False
xmlDOM.setProperty "ServerHTTPRequest", True
xmlDOM.Load("https://rss.nytimes.com/services/xml/rss/nyt/World.xml")
Set feeditems = xmlDOM.getElementsByTagName("item")
Response.Write "LEN=" & feeditems.length & vbCrLf
Set Item = feeditems(0)
Response.Write "AFTER_ITEM=" & Err.Number & "|" & Err.Description & vbCrLf
Err.Clear
For Each child In Item.childNodes
    Response.Write "NODE=" & child.nodeName & vbCrLf
    Response.Write "ERR=" & Err.Number & "|" & Err.Description & vbCrLf
    Err.Clear
Next
Response.Write "DONE" & vbCrLf
On Error Goto 0
%>
