/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
 * Contact: https://g3pix.com.br
 * Project URL: https://g3pix.com.br/axonasp
 *
 * This Source Code Form is subject to the terms of the Mozilla Public
 * License, v. 2.0. If a copy of the MPL was not distributed with this
 * file, You can obtain one at https://mozilla.org/MPL/2.0/.
 *
 * Attribution Notice:
 * If this software is used in other projects, the name "AxonASP Server"
 * must be cited in the documentation or "About" section.
 *
 * Contribution Policy:
 * Modifications to the core source code of AxonASP Server must be
 * made available under this same license terms.
 */
package axonvm

import (
	"bytes"
	"testing"
)

// TestASPClassResetCanCallLifecycleMembers verifies Call-based same-class lifecycle dispatch works.
func TestASPClassResetCanCallLifecycleMembers(t *testing.T) {
	source := `<%
Class Probe
	Private items()
	Private marker

	Private Sub Class_Initialize
		ReDim items(0)
		items(0) = "ok"
		marker = "init"
	End Sub

	Private Sub Class_Terminate
		Erase items
		marker = "term"
	End Sub

	Public Function Reset
		Call Class_Terminate
		Call Class_Initialize
		Reset = marker & ":" & items(0)
	End Function
End Class

Dim p
Set p = New Probe
Response.Write p.Reset()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "init:ok" {
		t.Fatalf("expected init:ok output, got %q", output.String())
	}
}

// TestASPCallClassMethodWithoutErase isolates Call-based same-class dispatch without array cleanup.
func TestASPCallClassMethodWithoutErase(t *testing.T) {
	source := `<%
Class Probe
	Private marker

	Private Sub Class_Initialize
		marker = "init"
	End Sub

	Private Sub Touch
		marker = "touch"
	End Sub

	Public Function RunTouch
		Call Touch
		RunTouch = marker
	End Function
End Class

Dim p
Set p = New Probe
Response.Write p.RunTouch()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "touch" {
		t.Fatalf("expected touch output, got %q", output.String())
	}
}

// TestASPEraseStatementClearsClassArray isolates the Erase statement on one class field.
func TestASPEraseStatementClearsClassArray(t *testing.T) {
	source := `<%
Class Probe
	Private items()

	Private Sub Class_Initialize
		ReDim items(1)
		items(0) = "a"
		items(1) = "b"
	End Sub

	Public Function ResetItems
		Erase items
		ReDim items(0)
		items(0) = "z"
		ResetItems = items(0)
	End Function
End Class

Dim p
Set p = New Probe
Response.Write p.ResetItems()
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "z" {
		t.Fatalf("expected z output, got %q", output.String())
	}
}

// TestASPReducedRSSWriterFlow verifies the helper pattern used by rss.asp executes end-to-end.
func TestASPReducedRSSWriterFlow(t *testing.T) {
	source := `<%
Class kwRSS_writer
	Private Items(10, 8)
	Private CurrentItem
	Public ChannelURL, ChannelTitle, ChannelDesc, ChannelLanguage
	Private myXML

	Private Sub Class_Initialize
		CurrentItem = -1
	End Sub

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

	Public Function AddNew
		CurrentItem = CurrentItem + 1
	End Function

	Public Function GetRSS
		Set myXML = New aspXML
		myXML.OpenTag "rss"
		myXML.AddAttribute "version", "2.0"
		myXML.OpenTag "channel"
		myXML.QuickTag "title", ChannelTitle
		myXML.QuickTag "link", ChannelURL
		myXML.QuickTag "description", ChannelDesc
		myXML.QuickTag "language", ChannelLanguage
		myXML.OpenTag "item"
		myXML.OpenTag "title"
		myXML.AddData Items(0, 0)
		myXML.CloseTag
		myXML.OpenTag "link"
		myXML.AddData Items(0, 1)
		myXML.CloseTag
		myXML.OpenTag "pubDate"
		myXML.AddData Items(0, 3)
		myXML.CloseTag
		myXML.OpenTag "author"
		myXML.AddData Items(0, 4)
		myXML.CloseTag
		myXML.OpenTag "guid"
		myXML.AddData Items(0, 5)
		myXML.CloseTag
		myXML.OpenTag "description"
		myXML.AddData Items(0, 2)
		myXML.CloseTag
		myXML.CloseTag
		myXML.CloseAllTags
		GetRSS = myXML.GetXML
		Set myXML = Nothing
	End Function
End Class

Class aspXML
	Private top
	Private TagArray()
	Private XML

	Private Sub Class_Initialize
		ReDim TagArray(10)
		top = -1
		XML = "<?xml version=""1.0"" encoding=""UTF-8""?>"
	End Sub

	Private Sub Class_Terminate
		top = Null
		XML = Null
		Erase TagArray
	End Sub

	Public Function Reset
		Call Class_Terminate
		Call Class_Initialize
	End Function

	Public Function OpenTag(tagName)
		If top > UBound(TagArray) Then
			ReDim Preserve TagArray(UBound(TagArray) + 10)
		End If
		top = top + 1
		TagArray(top) = tagName
		XML = XML & "<" & tagName & ">"
	End Function

	Public Function QuickTag(tagName, Data)
		XML = XML & "<" & tagName & ">" & CheckString(Data) & "</" & tagName & ">"
	End Function

	Public Function AddAttribute(attribName, attribValue)
		Dim lastTag, textRemoved
		lastTag = InStrRev(XML, ">")
		textRemoved = Right(XML, Len(XML) - lastTag)
		XML = Left(XML, lastTag - 1)
		XML = XML & " " & attribName & "=""" & attribValue & """>"
		XML = XML & textRemoved
	End Function

	Public Function AddData(Data)
		XML = XML & CheckString(Data)
	End Function

	Public Function CloseTag()
		Dim tagName
		tagName = TagArray(top)
		XML = XML & "</" & tagName & ">"
		top = top - 1
	End Function

	Public Function CloseAllTags()
		Dim tagName
		While (top >= 0)
			tagName = TagArray(top)
			XML = XML & "</" & tagName & ">"
			top = top - 1
		Wend
	End Function

	Public Function GetXML()
		GetXML = XML
	End Function

	Private Function CheckString(data)
		CheckString = data
	End Function
End Class

Dim rss
Set rss = New kwRSS_writer
rss.ChannelTitle = "Site"
rss.ChannelURL = "https://example.test"
rss.ChannelDesc = "Desc"
rss.ChannelLanguage = "en"
rss.AddNew
rss.SetTitle "Title"
rss.SetLink "https://example.test/item"
rss.SetDesc "Body"
rss.SetPubDate "Mon, 01 Jan 2024 00:00:00 GMT"
rss.SetAuthor "Author"
rss.SetGuid "guid-1"
Response.Write Left(rss.GetRSS, 5)
%>`

	compiler := NewASPCompiler(source)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVM(compiler.Bytecode(), compiler.Constants(), compiler.GlobalsCount())
	host := NewMockHost()
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "<?xml" {
		t.Fatalf("expected XML prefix output, got %q", output.String())
	}
}
