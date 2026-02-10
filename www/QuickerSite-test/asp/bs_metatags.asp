
	<%if secondAdmin.bPageTitleTag then%><tr><td class=QSlabel>TITLE-tag:</td><td><input name="sSEOtitle" value="<%=sanitize(page.sSEOtitle)%>" type=text size=70 maxlength=200 /></td></tr><%end if
if secondAdmin.bPageKeywords then%><tr><td class=QSlabel><%=l("keywords")%>:</td><td><input name="sKeywords" value="<%=sanitize(page.sKeywords)%>" type=text size=45 maxlength=2048 />&nbsp;<i>(META-tag keywords)</i></td></tr><%end if
if secondAdmin.bPageDescription then%><tr><td class=QSlabel><%=l("description")%>:</td><td><textarea name="sDescription" cols=50 rows=2><%=sanitize(page.sDescription)%></textarea>&nbsp;<i>(META-tag description)</i></td></tr><%end if
if secondAdmin.bPageAdditionalHeader then%><tr><td class=QSlabel><%=l("header")%>:</td><td><textarea name="sHeader" cols=50 rows=2><%=quotRep(page.sHeader)%></textarea></td></tr><%end if%>
