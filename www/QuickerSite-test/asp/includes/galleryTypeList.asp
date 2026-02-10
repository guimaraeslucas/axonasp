
<%const QS_gallery_LB="10"
const QS_gallery_SC="15"
const QS_gallery_SS="20"
const QS_gallery_NS="25"

Class cls_galleryTypeList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add QS_gallery_LB,"ColorBox"
list.Add QS_gallery_SC,"JQuery Cycle"
list.Add QS_gallery_SS,"SlideShow"
list.Add QS_gallery_NS,"Nivo Slider"

End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub
Public Function showSelected(mode, selected)
showSelected = ""
 selected=convertGetal(selected)
Select Case mode
Case "single"
showSelected = list(selected)
Case "option"
Dim key
For each key in list
showSelected = showSelected & "<option value='" & key &"'"
If convertStr(selected) = convertStr(key) Then
showSelected = showSelected & " selected"
End If
showSelected = showSelected & ">" & list(key) & "</option>"
Next
End Select
End Function
End class
Class cls_galleryCycleList
Public list
Private Sub Class_Initialize
Set list = Server.CreateObject("scripting.dictionary")
list.Add "blindX",""
list.Add "blindY",""
list.Add "blindZ",""
list.Add "cover",""
list.Add "curtainX",""
list.Add "curtainY",""
list.Add "fade",""
list.Add "fadeZoom",""
list.Add "growX",""
list.Add "growY",""
list.Add "none",""
list.Add "scrollUp",""
list.Add "scrollDown",""
list.Add "scrollLeft",""
list.Add "scrollRight",""
list.Add "scrollHorz",""
list.Add "scrollVert",""
list.Add "shuffle",""
list.Add "slideX",""
list.Add "slideY",""
list.Add "toss",""
list.Add "turnUp",""
list.Add "turnDown",""
list.Add "turnLeft",""
list.Add "turnRight",""
list.Add "uncover",""
list.Add "wipe",""
list.Add "zoom",""
End Sub
Private Sub Class_Terminate
Set list = nothing
End Sub

Public Function showSelected(mode, selected)
	showSelected = ""
	 selected=convertStr(selected)
	Select Case mode
		Case "option"
		Dim key
		For each key in list
			showSelected = showSelected & "<option value='" & key &"'"
			If convertStr(selected) = convertStr(key) Then
			showSelected = showSelected & " selected"
			End If
			showSelected = showSelected & ">" & key & "</option>"
		Next
	End Select
End Function

Public Function showSelectedNS(mode, selected)
	
	showSelectedNS = ""
	selected=convertStr(selected)
	
	dim NSlist
	set NSlist=server.createobject("scripting.dictionary")
	NSlist.add "default",""
	NSlist.add "light",""
	NSlist.add "dark",""
	NSlist.add "bar",""
	
	Select Case mode
		Case "option"
			Dim key
			For each key in NSlist
				showSelectedNS = showSelectedNS & "<option value='" & key &"'"
				If convertStr(selected) = convertStr(key) Then
					showSelectedNS = showSelectedNS & " selected"
				End If
				showSelectedNS = showSelectedNS & ">" & key & "</option>"
			Next
	End Select
	
	set NSlist=nothing
	
End Function


Public Function showSelectedNS2(mode, selected)
	
	showSelectedNS2 = ""
	selected=convertStr(selected)
	
	dim NSlist
	set NSlist=server.createobject("scripting.dictionary")
	NSlist.add "random",""
	NSlist.add "sliceDown",""
	NSlist.add "sliceDownLeft",""	
	NSlist.add "sliceUp",""
	NSlist.add "sliceUpLeft",""
	NSlist.add "sliceUpDown",""
	NSlist.add "fold",""
	NSlist.add "fade",""
	NSlist.add "slideInRight",""
	NSlist.add "slideInLeft",""
	NSlist.add "boxRandom",""
	NSlist.add "boxRain",""
	NSlist.add "boxRainReverse",""
	NSlist.add "boxRainGrow",""
	NSlist.add "boxRainGrowReverse",""
	
	
	Select Case mode
		Case "option"
			Dim key
			For each key in NSlist
				showSelectedNS2 = showSelectedNS2 & "<option value='" & key &"'"
				If convertStr(selected) = convertStr(key) Then
					showSelectedNS2 = showSelectedNS2 & " selected"
				End If
				showSelectedNS2 = showSelectedNS2 & ">" & key & "</option>"
			Next
	End Select
	
	set NSlist=nothing
	
End Function


End class%>
