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
