
<%dim showPost
set showPost=new cls_post
pageBody=showPost.buildShortPost()
pageTitle=quotrep(replace(showPost.sSubject,"[","|R|R|R|",1,-1,1))
selectedPage.sSEOTitle=pageTitle%>
