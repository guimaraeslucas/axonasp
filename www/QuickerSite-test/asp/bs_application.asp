
<%if customer.bApplication and secondAdmin.bApplicationpath then%><tr><td class=QSlabel><%=l("applicationpath")%>:</td><td><input type=text maxlength=250 size=45 name=sApplication value="<%=quotRep(page.sApplication)%>" />&nbsp;<i>sCode:</i> <input type=text maxlength=48 size=10 name=sCode value="<%=quotRep(page.sCode)%>" /></td></tr><%end if%>
