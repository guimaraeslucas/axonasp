package server

import (
	"net/http/httptest"
	"path/filepath"
	"strings"
	"testing"
)

func TestForEachVBArray(t *testing.T) {
	code := `
		<%
		Option Explicit
		
		' Test 1: For Each over ReDim array
		Dim arr1 : arr1 = Array()
		ReDim arr1(2)
		arr1(0) = "item1"
		arr1(1) = "item2"
		arr1(2) = "item3"
		
		Dim item, result1
		result1 = ""
		For Each item In arr1
			result1 = result1 & item & ";"
		Next
		Response.Write "Test1: " & result1 & vbCrLf
		
		' Test 2: For Each over ReDim array with dictionaries
		Dim arr2 : arr2 = Array()
		ReDim arr2(1)
		
		Dim dict1, dict2
		Set dict1 = Server.CreateObject("Scripting.Dictionary")
		dict1.Add "name", "field1"
		dict1.Add "type", "text"
		Set arr2(0) = dict1
		
		Set dict2 = Server.CreateObject("Scripting.Dictionary")
		dict2.Add "name", "field2"
		dict2.Add "type", "submit"
		Set arr2(1) = dict2
		
		Dim dict, result2
		result2 = ""
		For Each dict In arr2
			Response.Write "TypeName: " & TypeName(dict) & vbCrLf
			If IsObject(dict) Then
				Response.Write "Is Object: True" & vbCrLf
				Response.Write "Count: " & dict.Count & vbCrLf
				If dict.Count > 0 Then
					result2 = result2 & dict("name") & ":" & dict("type") & ";"
				Else
					result2 = result2 & "EMPTY:"
				End If
			End If
		Next
		Response.Write "Test2: " & result2 & vbCrLf
		
		' Test 3: Nested For Each
		Dim arr3 : arr3 = Array()
		ReDim arr3(2)
		arr3(0) = "a"
		arr3(1) = "b"
		arr3(2) = "c"
		
		Dim outer, inner, result3
		result3 = ""
		For Each outer In arr3
			For Each inner In arr3
				result3 = result3 & outer & inner & ","
			Next
		Next
		Response.Write "Test3: " & result3 & vbCrLf
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, filepath.Join("..", "www", "test.asp"), w, r)
	if err != nil {
		t.Fatalf("ExecuteASPFile failed: %v", err)
	}

	output := w.Body.String()
	t.Logf("Output:\n%s", output)

	// Verify Test 1: For Each over ReDim array
	if !strings.Contains(output, "Test1: item1;item2;item3;") {
		t.Errorf("Test1 failed - expected 'item1;item2;item3;' in output, got: %s", output)
	}

	// Verify Test 2: For Each over array with dictionaries
	if !strings.Contains(output, "Test2: field1:text;field2:submit;") {
		t.Errorf("Test2 failed - expected 'field1:text;field2:submit;' in output, got: %s", output)
	}

	// Verify Test 3: Nested For Each
	if !strings.Contains(output, "Test3:") && !strings.Contains(output, "aa,") {
		t.Errorf("Test3 failed - expected nested iteration output, got: %s", output)
	}
}

func TestForEachVBArrayWithJSON(t *testing.T) {
	code := `
		<%
		Option Explicit
		
		' Simulate what aspForm.asp does
		Dim counter : counter = 3
		Dim arr : arr = Array()
		ReDim arr(counter - 1)
		
		Dim i
		For i = 0 To counter - 1
			Dim dict
			Set dict = Server.CreateObject("Scripting.Dictionary")
			dict.Add "type", "field" & i
			dict.Add "name", "name" & i
			Set arr(i) = dict
		Next
		
		' Now iterate like the JSON generator does
		Dim fieldkey, output
		output = "["
		Dim first : first = True
		For Each fieldkey In arr
			If Not first Then output = output & ","
			first = False
			
			If IsObject(fieldkey) Then
				output = output & "{""type"":""" & fieldkey("type") & """,""name"":""" & fieldkey("name") & """}"
			End If
		Next
		output = output & "]"
		
		Response.Write output
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, filepath.Join("..", "www", "test.asp"), w, r)
	if err != nil {
		t.Fatalf("ExecuteASPFile failed: %v", err)
	}

	output := w.Body.String()
	t.Logf("JSON Output:\n%s", output)

	// Verify JSON array is not empty
	if strings.Contains(output, "[]") {
		t.Errorf("JSON array is empty - For Each loop not working with VBArray")
	}

	// Verify JSON contains expected fields
	if !strings.Contains(output, `"type":"field0"`) {
		t.Errorf("Expected field0 in JSON, got: %s", output)
	}
	if !strings.Contains(output, `"type":"field1"`) {
		t.Errorf("Expected field1 in JSON, got: %s", output)
	}
	if !strings.Contains(output, `"type":"field2"`) {
		t.Errorf("Expected field2 in JSON, got: %s", output)
	}
}

func TestForEachArrayFunction(t *testing.T) {
	code := `
		<%
		' Test For Each with Array() function result
		Dim arr : arr = Array("x", "y", "z")
		Dim item, result
		result = ""
		For Each item In arr
			result = result & item
		Next
		Response.Write result
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, filepath.Join("..", "www", "test.asp"), w, r)
	if err != nil {
		t.Fatalf("ExecuteASPFile failed: %v", err)
	}

	output := w.Body.String()
	t.Logf("Output: %s", output)

	if !strings.Contains(output, "xyz") {
		t.Errorf("Expected 'xyz', got: %s", output)
	}
}

func TestASPLiteFormPattern(t *testing.T) {
	// This test simulates the exact pattern used in aspForm.asp
	code := `
		<%
		Option Explicit
		
		' Create a counter and dictionary collection (simulating allFields)
		Dim allFields : Set allFields = Server.CreateObject("Scripting.Dictionary")
		Dim counter : counter = 0
		
		' Add fields like aspForm.asp does
		Dim field1 : Set field1 = Server.CreateObject("Scripting.Dictionary")
		field1.Add "type", "text"
		field1.Add "name", "username"
		allFields.Add counter, field1
		counter = counter + 1
		
		Dim field2 : Set field2 = Server.CreateObject("Scripting.Dictionary")
		field2.Add "type", "hidden"
		field2.Add "name", "token"
		allFields.Add counter, field2
		counter = counter + 1
		
		Dim field3 : Set field3 = Server.CreateObject("Scripting.Dictionary")
		field3.Add "type", "submit"
		field3.Add "name", "submit"
		allFields.Add counter, field3
		counter = counter + 1
		
		' Now create array and fill it (exactly like aspForm.asp line 182-208)
		Dim arr : arr = Array()
		ReDim arr(counter - 1)
		
		Dim fieldkey
		For Each fieldkey In allFields
			Set arr(fieldkey) = allFields(fieldkey)
		Next
		
		' Now iterate over the array (like JSON generator does in asplite.asp)
		Dim item, output
		output = "Fields: "
		For Each item In arr
			If IsObject(item) Then
				If TypeName(item) = "Dictionary" Then
					output = output & item("type") & ":" & item("name") & ";"
				Else
					output = output & "Unknown;"
				End If
			Else
				output = output & "NotObject;"
			End If
		Next
		
		Response.Write output
		%>
	`

	w := httptest.NewRecorder()
	r := httptest.NewRequest("GET", "/test.asp", nil)

	processor := NewASPProcessor(&ASPProcessorConfig{
		RootDir:       "../www",
		ScriptTimeout: 30,
		DebugASP:      false,
	})

	err := processor.ExecuteASPFile(code, filepath.Join("..", "www", "test.asp"), w, r)
	if err != nil {
		t.Fatalf("ExecuteASPFile failed: %v", err)
	}

	output := w.Body.String()
	t.Logf("Output:\n%s", output)

	// Verify all fields are present
	if !strings.Contains(output, "text:username") {
		t.Errorf("Expected 'text:username' in output, got: %s", output)
	}
	if !strings.Contains(output, "hidden:token") {
		t.Errorf("Expected 'hidden:token' in output, got: %s", output)
	}
	if !strings.Contains(output, "submit:submit") {
		t.Errorf("Expected 'submit:submit' in output, got: %s", output)
	}
	
	// Verify no "NotObject" or "Unknown" entries
	if strings.Contains(output, "NotObject") {
		t.Errorf("Found 'NotObject' - items in array are not objects: %s", output)
	}
	if strings.Contains(output, "Unknown") {
		t.Errorf("Found 'Unknown' - items in array are not Dictionary type: %s", output)
	}
}
