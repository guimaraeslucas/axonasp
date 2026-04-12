<!-- #include file="../asplite-test/asplite/asplite.asp"-->
<%
Response.ContentType = "text/plain"
On Error Resume Next
Dim rss : Set rss = aspL.plugin("rss")
Response.Write "P1_ERR=" & Err.Number & "|" & Err.Description & vbCrLf
Err.Clear
rss.maxitems = 1
Response.Write "P2_ERR=" & Err.Number & "|" & Err.Description & vbCrLf
Err.Clear
Dim txt
txt = rss.read("https://rss.nytimes.com/services/xml/rss/nyt/World.xml")
Response.Write "P3_ERR=" & Err.Number & "|" & Err.Description & vbCrLf
Err.Clear
Response.Write "TXT_LEN=" & Len(txt) & vbCrLf
Response.Write txt
On Error Goto 0
%>
