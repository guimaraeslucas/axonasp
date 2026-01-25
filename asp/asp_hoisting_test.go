/*
 * AxonASP Server - Version 1.0
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas Guimarães - G3pix Ltda
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
package asp

import (
	"strings"
	"testing"
)

func TestFunctionHoisting(t *testing.T) {
	code := `<%
	Option Explicit
	Dim url
	url = "test.jpg"
	
	'
	' 
	select case trim(lcase(GetFileExtension(url)))
		case "jpg"
			Response.Write "JPEG"
		case else
			Response.Write "Other"
	end select

	Function GetFileExtension(path)
		GetFileExtension = "jpg"
	End Function
	%>`

	parser := NewASPParser(code)
	result, err := parser.Parse()

	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}

	if result.CombinedProgram == nil {
		t.Errorf("Expected CombinedProgram to be non-nil")
	}
}

func TestFunctionHoistingComplex(t *testing.T) {
	// Replicating the user's specific context more closely if possible
	code := `<%
	Dim blockURL
	blockURL = "something.png"

	'
	'
	select case trim(lcase(GetFileExtension(blockURL)))
	case "png", "jpg", "jpeg", "gif"
		' do something
	case "mp4"
		' do something else
	end select

	function GetFileExtension(p)
		GetFileExtension = "png"
	end function
	%>`

	parser := NewASPParser(code)
	result, err := parser.Parse()

	if err != nil {
		t.Fatalf("Parse error: %v", err)
	}

	// Ensure we can extract the VB code and it looks sane
	if !strings.Contains(result.CombinedVBCode, "GetFileExtension") {
		t.Errorf("Combined VB code should contain function name")
	}
}

func TestFunctionHoistingMultiBlock(t *testing.T) {
	code := `<%
	Dim url
	url = "test.jpg"
	%>
	<html>
	<!-- Some HTML comment -->
	</html>
	<%
	'
	'
	select case trim(lcase(GetFileExtension(url)))
		case "jpg"
			Response.Write "JPEG"
	end select
	%>
	<%
	Function GetFileExtension(path)
		GetFileExtension = "jpg"
	End Function
	%>`

	parser := NewASPParser(code)
	result, err := parser.Parse()

	if err != nil {
		t.Fatalf("Parse error: %v\nCombined Code:\n%s", err, result.CombinedVBCode)
	}
}

func TestFunctionHoistingClass(t *testing.T) {
	code := `<%
	dim aspL
	set aspL=new cls_asplite

	class cls_asplite
		Private Sub Class_Initialize()
			on error resume next

			dim blockURL
			blockURL = "test.png"
			if instr(blockURL,"?")>0 then blockURL=left(blockURL,instr(blockURL,"?")-1) : end if
			
			'
			'
			select case trim(lcase(GetFileExtension(blockURL)))
				case "png"
					Response.Write "PNG"
			end select
		End Sub

		public function GetFileExtension(str)
			GetFileExtension = "png"
		end function
	end class
	%>`

	parser := NewASPParser(code)
	result, err := parser.Parse()

	if err != nil {
		t.Fatalf("Parse error: %v\nCombined Code:\n%s", err, result.CombinedVBCode)
	}
}
