<%
	'==============================================================
	' RSS/RDF Syndicate Writer v0.95
	' http://www.kattanweb.com/webdev
	'--------------------------------------------------------------
	' Copyright(c) 2002, KattanWeb.com
	'
	' Change Log:
	'--------------------------------------------------------------
	'==============================================================

class kwRSS_writer

	Private Items(500, 8)
	Private CurrentItem
	Public ChannelRSSURI, ChannelURL, ChannelTitle, ChannelDesc, ChannelLanguage
	Public ImageTitle, ImageLink, ImageURL
	Public TextInputURL, TextInputTitle, TextInputDesc, TextInputName
	
	Private myXML

    '>>>>>>>> Setup Initialize event, called automtially when creating an instant of this class using
	'	      Set MyXML = new kwRSS_writer
	Private Sub Class_Initialize
		CurrentItem = -1
	End Sub

    '>>>>>>>> Setup Terminate event, called automtially when killing an instant of this class using
	'	      Set MyXML = nothing
	Private Sub Class_Terminate
		Erase Items
	End Sub

	Public Function SetTitle(ItemTitle)
		Items(CurrentItem, 0) = ItemTitle
	End Function

	Public Function SetLink(ItemLink)
		Items(CurrentItem, 1) = ItemLink
	End Function

	Public Function SetDesc(ItemDesc)
		Items(CurrentItem, 2) = ItemDesc
	End Function
	
	Public Function SetPubDate(ItemDate)
		Items(CurrentItem, 3) = ItemDate
	End Function	
	
	Public Function SetAuthor(ItemAuthor)
		Items(CurrentItem, 4) = ItemAuthor
	End Function
	
	Public Function SetGuid(ItemGUID)
		Items(CurrentItem, 5) = ItemGUID
	End Function
	
	Public Function SetEnclosure(ItemEnclosure)
		Items(CurrentItem, 7) = ItemEnclosure
	End Function	
	
	Public Function setComments(ItemComments)
		if not isLeeg(ItemComments) then Items(CurrentItem, 6) = ItemComments
	End Function
	
	Public Function AddNew
		CurrentItem = CurrentItem + 1
	End Function

	Public Function GetRSS
		set myXML = new aspXML
			myXML.OpenTag("rss")
			myXML.AddAttribute "version", "2.0"
			myXML.OpenTag("channel")
				myXML.QuickTag "title", ChannelTitle
				myXML.QuickTag "link", ChannelURL
				myXML.QuickTag "description", ChannelDesc
				myXML.QuickTag "language", ChannelLanguage
	
			if ImageURL <> "" then
				myXML.OpenTag("image")
					myXML.QuickTag "title", ImageTitle
					myXML.QuickTag "link", ImageLink
					myXML.QuickTag "url", ImageURL
				myXML.CloseTag
			end if
			
			dim i, ItemTitle, ItemLink, ItemDesc, ItemPubDate, ItemAuthor, ItemGUID,ItemComments,ItemEnclosure
			for i = 0 to CurrentItem
				ItemTitle		= Items(i, 0)
				ItemLink		= Items(i, 1)
				ItemDesc		= Items(i, 2)
				ItemPubDate		= Items(i, 3)
				ItemAuthor		= Items(i, 4)
				ItemGUID		= Items(i, 5)
				ItemComments	= Items(i, 6)
				ItemEnclosure	= Items(i, 7)				

				myXML.OpenTag "item"
					myXML.OpenTag "title"
						myXML.AddData ItemTitle
					myXML.CloseTag
					myXML.OpenTag "link"
						myXML.AddData ItemLink
					myXML.CloseTag
					myXML.OpenTag "pubDate"
						myXML.AddData ItemPubDate
					myXML.CloseTag
					
					if not isLeeg(ItemEnclosure) then
						myXML.OpenTag "enclosure"
						myXML.AddAttribute "url", ItemEnclosure
						myXML.AddAttribute "length", "0"
						myXML.AddAttribute "type", "image/jpeg"						
						myXML.CloseTag
					end if
					
					myXML.OpenTag "author"
						myXML.AddData ItemAuthor
					myXML.CloseTag
					myXML.OpenTag "guid"
						myXML.AddData ItemGUID
					myXML.CloseTag
					
					if not isLeeg(ItemComments) then 
						myXML.OpenTag "comments"
							myXML.AddData ItemComments
						myXML.CloseTag
					end if
									
					if ItemDesc <> "" then
						myXML.OpenTag "description"
							myXML.AddData ItemDesc
						myXML.CloseTag
					end if
				
				myXML.CloseTag
				
			next

			if TextInputTitle <> "" then
				myXML.OpenTag "textinput"
					myXML.QuickTag "title", TextInputTitle
					myXML.QuickTag "description", TextInputDesc
					myXML.QuickTag "name", TextInputName
					myXML.QuickTag "link",  TextInputURL
				myXML.CloseTag
			end if		
			

			myXML.CloseAllTags
			GetRSS = myXML.GetXML
			
		Set myXML = nothing
		
	end function

	
end class


' ---------------------------------------------------
'                    aspXML v1.0
' ---------------------------------------------------
' Author: Rami Kattan
' Web Site: http://www.kattanweb.com/webdev
' Email:  aspXML@kattanweb.8k.com
' Date:   July 3, 2002
'
' This class with make easy the construction of XML
' files using simple ASP, without any components.
'
' Features:
'  - Keep track of opened tags, and closing will close
'    last open one.
'  - Can open tags with attributes passed as string
'  - Automatic format for tag names with special characters.
'  - Automatic check if data inside the tag need CData or no.
'  - Can add Date using XSL date format.
' ---------------------------------------------------

class aspXML
	Private top
	Private TagArray()
	Private XML

    '>>>>>>>> Setup Initialize event, called automtially when creating an instant of this class using
	'	      Set MyXML = new aspXML
	Private Sub Class_Initialize
		Redim TagArray(10)
		top = -1
		XML = "<?xml version=""1.0"" encoding=" & """" & QS_CHARSET & """" & "?>" & vbCrLf
	End Sub

    '>>>>>>>> Setup Terminate event, called automtially when killing an instant of this class using
	'	      Set MyXML = nothing
	Private Sub Class_Terminate
		top = null
		XML = null
		Erase TagArray
	End Sub
	
    '>>>>>>>> Reset the class, as if it was just created, Use with care
	Public Function Reset
		call Class_Terminate
		call Class_Initialize
	End Function

    '>>>>>>>> Open a new element tag
	Public Function OpenTag(tagName)
		tagName = tagName
		top = top + 1
		if top > ubound(TagArray) then
			ReDim Preserve TagArray(ubound(TagArray) + 10)
		end if
		TagArray(top) = tagName
		XML = XML & "<" & tagName & ">"
		if top = 0 then	XML = XML & vbCrLf 'Code format, root tag is on separate line
	end function

    '>>>>>>>> Opens a new tag, add the data, and close the tag
	Public Function QuickTag(tagName, Data)
		tagName = tagName
		XML = XML & "<" & tagName & ">" & CheckString(Data) & "</" & tagName & ">" & vbCrLf
	end function

    '>>>>>>>> Add an attribute to the last open tag (can be used before or after adding data)
	Public Function AddAttribute(attribName, attribValue)
		dim lastTag, TextRemoved
		lastTag = inStrRev(XML, ">")
		TextRemoved = Right(XML, len(XML) - lastTag)
		XML = Left(XML, lastTag - 1)
		XML = XML & " " & attribName & "=""" & attribValue & """>"
		XML = XML & TextRemoved
	End function

    '>>>>>>>> Add data to current open tag (automatic check if need CDATA or no)
	Public Function AddData(Data)
		XML = XML & CheckString(Data)
	end function

    '>>>>>>>> Add Comment in the current location
	Public Function AddComment(Data)
		XML = XML & "<!--" & Data & "-->"
	end function

	'>>>>>>>> Close last open tag
	Public Function CloseTag()
		dim tagName
		tagName = TagArray(top)
		XML = XML & "</" & tagName & ">" & vbCrLf
		top = top - 1
	end function

    '>>>>>>>> Close all open tags, including main root tag
	'after calling this function, it is not recomended opening new
	'tags as XML can only have 1 root element
	Public Function CloseAllTags()
		dim tagName
		while (top >= 0)
			tagName = TagArray(top)
			XML = XML & "</" & tagName & ">" & vbCrLf
			top = top - 1
		wend
	end function

    '>>>>>>>> Returns the XML final code
	Public Function GetXML()
		GetXML = XML
	end function

'---------------------------------------------------------------
' Special internal functions
'---------------------------------------------------------------

    '>>>>>>>> Format the data with or without CData
	Private function CheckString(data)
		dim need
		need = false
		if instr(data, "<") then need = true
		if instr(data, "&") then need = true
		if need then
			CheckString = "<![CDATA[" & data & "]]>"
		else
			CheckString = data
		end if
	end function

end class
%>