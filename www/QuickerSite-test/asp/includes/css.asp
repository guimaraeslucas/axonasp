
<%function css()
'exit function
css="<style type=""text/css"">"& vbcrlf
css=css & ".QSlabel"
css=css & "{"
css=css & " FONT-STYLE: italic;"
if bltr then
css=css & " TEXT-ALIGN: right"
else
css=css & " TEXT-ALIGN: left"
end if
css=css & "}"& vbcrlf
css=css & ".header"
css=css & "{"
css=css & " FONT-WEIGHT: bolder"
css=css & "}"& vbcrlf
css=css & ".offLine"
css=css & "{"
css=css & " COLOR: #b0b0b0"
css=css & "}"& vbcrlf
css=css & "." & convertTF(false)
css=css & "{"
css=css & "	BACKGROUND-COLOR: #df0000"
css=css & "}"& vbcrlf
css=css & ".main ." & convertTF(true)
css=css & "{"
css=css & "	BACKGROUND-COLOR: #00df00"
css=css & "}"& vbcrlf
css=css & ".cs_silent {BACKGROUND-COLOR: #FFFFFF;color:#000000}"& vbcrlf
css=css & ".cs_profile {BACKGROUND-COLOR: #cccc99;color:#cc6600}"& vbcrlf
css=css & ".cs_read {BACKGROUND-COLOR: #ffffcc;color:#66cc33}"& vbcrlf
css=css & ".cs_write {BACKGROUND-COLOR: #ccccff;color:#9900ff}"& vbcrlf
css=css & ".cs_readwrite {BACKGROUND-COLOR: #777777;color:#FFFFFF}"& vbcrlf
css=css & ".cs_readwrite A {color:#FFFFFF}"& vbcrlf
css=css & ".iM {vertical-align:middle;border:0px;margin:0px 0px 3px 2px;text-decoration:none}"& vbcrlf
css=css & ".pI {PADDING-RIGHT: 0px;PADDING-LEFT: 0px;PADDING-BOTTOM: 0px;PADDING-TOP: 0px;MARGIN: 0px;background-color:" & customer.publicIconColor	& "}"& vbcrlf
css=css & ".piT {PADDING-RIGHT: 0px;PADDING-LEFT: 0px;PADDING-BOTTOM: 0px;PADDING-TOP: 0px;width:16px;height:16px;MARGIN: 0px;}"& vbcrlf
css=css & "</style>"& vbcrlf
end function
dim  menulist, nmbrParentMenus, bLTR, bRTL, sMenuLinkColor, sMenuLinkHoverColor, sIntranetMenuLinkHoverColor
bLTR=tdir=QS_ltr
bRTL=tdir=QS_rtl
function getVMenu(intranet)
end function
function getHMenu(intranet)
end function
sub getValuesForMenu(intranet)
end sub
function fixedCSS
fixedCSS=fixedCSS & vbcrlf& vbcrlf & "/* START FIXED CSS */"
fixedCSS=fixedCSS & vbcrlf& vbcrlf & "/* END FIXED CSS */"
end function
function getQSMenuCSS
end function%>
