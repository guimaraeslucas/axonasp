/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 */
package axonvm

import (
	"strings"
	"testing"
)

func TestVB6GoToAndLabels(t *testing.T) {
	// Test basic backward jump
	source1 := `<%
	Dim x
	x = 0
	MyLabel:
	x = x + 1
	If x < 5 Then GoTo MyLabel
	Response.Write x
	%>`
	out1 := runVBSAndGetOutput(t, source1)
	if out1 != "5" {
		t.Fatalf("Backward jump failed: expected '5', got %q", out1)
	}

	// Test basic forward jump
	source2 := `<%
	Response.Write "Start "
	GoTo SkipMe
	Response.Write "Middle "
	SkipMe:
	Response.Write "End"
	%>`
	out2 := runVBSAndGetOutput(t, source2)
	if out2 != "Start End" {
		t.Fatalf("Forward jump failed: expected 'Start End', got %q", out2)
	}

	// Test nested GoTo in Sub
	source3 := `<%
	Sub TestGoTo()
		Dim i
		i = 1
		Again:
		Response.Write i
		i = i + 1
		If i <= 3 Then GoTo Again
	End Sub
	TestGoTo()
	%>`
	out3 := runVBSAndGetOutput(t, source3)
	if out3 != "123" {
		t.Fatalf("Nested GoTo in Sub failed: expected '123', got %q", out3)
	}

	// Test label scoping (should fail if label is in another Sub)
	source4 := `<%
	Sub Sub1()
		MyLabelInSub1:
		Response.Write "In Sub1"
	End Sub
	Sub Sub2()
		GoTo MyLabelInSub1
	End Sub
	Sub2()
	%>`
	compiler := NewASPCompiler(source4)
	err := compiler.Compile()
	if err == nil {
		t.Fatal("Expected compile error for cross-procedure GoTo, but it compiled successfully")
	}
	if !strings.Contains(err.Error(), "Label 'mylabelinsub1' not defined in procedure 'Sub2'") {
		t.Fatalf("Expected scoping error message, got: %v", err)
	}
}

func TestVB6MultipleLabelsOnSameLine(t *testing.T) {
	source := `<%
	GoTo Label2
	Label1: Response.Write "1" : GoTo Label3
	Label2: Response.Write "2" : GoTo Label1
	Label3: Response.Write "3"
	%>`
	out := runVBSAndGetOutput(t, source)
	if out != "213" {
		t.Fatalf("Multiple labels on same line failed: expected '213', got %q", out)
	}
}

func TestVB6NumericLabels(t *testing.T) {
	// VB6 supports numeric line numbers as labels
	source := `<%
	Dim x
	x = 1
	GoTo 100
	x = 2
	100:
	Response.Write x
	%>`
	out := runVBSAndGetOutput(t, source)
	if out != "1" {
		t.Fatalf("Numeric labels failed: expected '1', got %q", out)
	}
}
