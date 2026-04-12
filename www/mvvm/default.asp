<%
@Language = "VBSCRIPT" CodePage = "65001"
%>
<%
Option Explicit
%>
<!--#include file="model/site_model.asp" -->
<!--#include file="viewmodel/site_viewmodel.asp" -->
<%
Dim route
Dim htmlOutput

route = LCase(Trim(Request.QueryString("page")))

If route = "" Then
    route = "home"
End If

Select Case route
    Case "home"
        htmlOutput = RenderMvvmHome()
    Case "example"
        htmlOutput = RenderMvvmExample()
    Case Else
        htmlOutput = RenderMvvmHome()
End Select

Response.Write htmlOutput
%>
