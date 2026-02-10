
<%if not devVersion() and C_ADMINPASSWORD=QS_defaultPW then Response.Redirect ("noaccess.htm")
if not convertBool(Session(cId & "isAUTHENTICATEDasADMIN")) then Response.Redirect ("noaccess.htm")%>
