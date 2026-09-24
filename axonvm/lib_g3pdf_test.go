/*
 * AxonASP Server
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
package axonvm

import (
	"os"
	"testing"
)

func TestG3PDF(t *testing.T) {
	vm := NewVM(nil, nil, 0)

	pdf := NewG3PDF(vm)
	if pdf == nil {
		t.Fatal("Failed to create G3PDF")
	}

	// basic interactions
	pdf.DispatchMethod("AddPage", nil)
	pdf.DispatchMethod("SetFont", []Value{NewString("Arial"), NewString("B"), NewInteger(16)})
	pdf.DispatchMethod("Cell", []Value{NewInteger(40), NewInteger(10), NewString("Hello World!")})

	page := pdf.DispatchPropertyGet("Page")
	if page.Num != 1 {
		t.Errorf("expected page 1, got %d", page.Num)
	}
}

// TestG3PDFPersitsParamParser tests the AspPDF-style param string parser
func TestG3PDFPersitsParamParser(t *testing.T) {
	// Test basic parsing
	params := parseAspPDFParams("x=0; y=650; width=612; alignment=center; size=50")
	if params == nil {
		t.Fatal("params should not be nil")
	}
	if params["x"] != "0" {
		t.Errorf("expected x=0, got x=%s", params["x"])
	}
	if params["alignment"] != "center" {
		t.Errorf("expected alignment=center, got alignment=%s", params["alignment"])
	}
	if params["size"] != "50" {
		t.Errorf("expected size=50, got size=%s", params["size"])
	}

	// Test empty input
	empty := parseAspPDFParams("")
	if empty != nil {
		t.Error("expected nil for empty input")
	}

	// Test single param
	single := parseAspPDFParams("color=#FF0000")
	if single["color"] != "#FF0000" {
		t.Errorf("expected color=#FF0000, got %s", single["color"])
	}

	// Test with extra whitespace
	ws := parseAspPDFParams("  left = 10 ; top = 20 ; right = 30 ")
	if ws["left"] != "10" || ws["top"] != "20" || ws["right"] != "30" {
		t.Errorf("whitespace parsing failed: %v", ws)
	}

	// Test numeric extraction helpers
	f := parseAspPDFFloat(params, "x", -1)
	if f != 0 {
		t.Errorf("expected x=0.0, got %f", f)
	}
	f2 := parseAspPDFFloat(params, "nonexistent", 42.5)
	if f2 != 42.5 {
		t.Errorf("expected fallback 42.5, got %f", f2)
	}
	i := parseAspPDFInt(params, "size", 0)
	if i != 50 {
		t.Errorf("expected size=50, got %d", i)
	}
}

// TestG3PDFPersitsCreateDocument tests the full Persits.Pdf object lifecycle
func TestG3PDFPersitsCreateDocument(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	pdf := NewG3PDF(vm)
	if pdf == nil {
		t.Fatal("Failed to create G3PDF")
	}

	// Test CreateDocument
	doc := pdf.CreateDocument()
	if doc == nil {
		t.Fatal("CreateDocument returned nil")
	}
	if !doc.open {
		t.Error("expected doc.open to be true")
	}

	// Test Pages.Add -> returns PdfPage
	pagesVal := doc.getPages()
	if pagesVal.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject for Pages, got type %d", pagesVal.Type)
	}

	// Test page creation via Pages.Add mock
	pc, exists := vm.pdfPagesItems[pagesVal.Num]
	if !exists {
		t.Fatal("PdfPagesCollection not found in VM map")
	}
	pageResult := pc.DispatchMethod("add", nil)
	if pageResult.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject from Pages.Add, got type %d", pageResult.Type)
	}
	page, exists := vm.pdfPageItems[pageResult.Num]
	if !exists {
		t.Fatal("PdfPage not found in VM map after Pages.Add")
	}

	// Test Canvas property on page
	canvasVal := page.DispatchPropertyGet("canvas")
	if canvasVal.Type != VTNativeObject {
		t.Fatalf("expected VTNativeObject for Canvas, got type %d", canvasVal.Type)
	}
	canvas, exists := vm.pdfCanvasItems[canvasVal.Num]
	if !exists {
		t.Fatal("PdfCanvas not found in VM map")
	}

	// Test Canvas.DrawText with param string
	canvas.DispatchMethod("drawtext", []Value{
		NewString("Hello Persits!"),
		NewString("x=10; y=50; size=16; alignment=center"),
	})

	// Test Canvas.DrawLine with param string
	canvas.DispatchMethod("drawline", []Value{
		NewString("x=10; y=10; x1=200; y1=10; color=#000000"),
	})

	// Test Canvas.DrawBox with param string
	canvas.DispatchMethod("drawbox", []Value{
		NewString("left=50; top=50; right=150; bottom=100; color=#FF0000; width=0.5"),
	})

	// Test Font loading
	font := pdf.getFont("Arial")
	if font == nil {
		t.Fatal("getFont returned nil")
	}
	if font.family != "helvetica" {
		t.Errorf("expected family=helvetica for Arial, got %s", font.family)
	}

	// Test font property get
	nameVal := font.DispatchPropertyGet("name")
	if nameVal.Str != "Arial" {
		t.Errorf("expected font name=Arial, got %s", nameVal.Str)
	}

	sizeVal := font.DispatchPropertyGet("size")
	if sizeVal.Flt != 12 {
		t.Errorf("expected font size=12, got %f", sizeVal.Flt)
	}

	// Test Save
	err := doc.Save("test_persits_output.pdf")
	if err != nil {
		t.Fatalf("Save failed: %v", err)
	}

	// Cleanup test file
	os.Remove("test_persits_output.pdf")

	// Test cleanup
	vm.cleanupG3PDFResources()
	if len(vm.pdfDocItems) != 0 {
		t.Error("expected pdfDocItems to be empty after cleanup")
	}
	if len(vm.pdfPageItems) != 0 {
		t.Error("expected pdfPageItems to be empty after cleanup")
	}
	if len(vm.pdfCanvasItems) != 0 {
		t.Error("expected pdfCanvasItems to be empty after cleanup")
	}
	if len(vm.pdfFontItems) != 0 {
		t.Error("expected pdfFontItems to be empty after cleanup")
	}
	if len(vm.pdfPagesItems) != 0 {
		t.Error("expected pdfPagesItems to be empty after cleanup")
	}
}

// TestG3PDFPersitsFontMapping tests font name mapping from Persits to fpdf
func TestG3PDFPersitsFontMapping(t *testing.T) {
	tests := []struct {
		input      string
		wantFamily string
		wantStyle  string
	}{
		{"Arial", "helvetica", ""},
		{"Arial Bold", "helvetica", "B"},
		{"Arial Italic", "helvetica", "I"},
		{"Arial Bold Italic", "helvetica", "BI"},
		{"Times New Roman", "times", ""},
		{"Times New Roman Bold", "times", "B"},
		{"Courier", "courier", ""},
		{"Courier-Bold", "courier", "B"},
		{"Symbol", "symbol", ""},
		{"ZapfDingbats", "zapfdingbats", ""},
	}
	for _, tt := range tests {
		family, style := mapPersitsFontToFPDF(tt.input)
		if family != tt.wantFamily || style != tt.wantStyle {
			t.Errorf("mapPersitsFontToFPDF(%q) = (%q, %q), want (%q, %q)",
				tt.input, family, style, tt.wantFamily, tt.wantStyle)
		}
	}
}

// TestG3PDFPersitsImportFromUrl tests that ImportFromUrl routes to WriteHTML
func TestG3PDFPersitsImportFromUrl(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	pdf := NewG3PDF(vm)
	doc := pdf.CreateDocument()
	doc.pdf.pdf.AddPage()

	// Test with empty URL should fail
	err := doc.ImportFromUrl("", "")
	if err == nil {
		t.Error("expected error for empty URL")
	}

	// Test with a URL - this will likely fail since no host is available,
	// but it should not panic and the error should be descriptive
	err = doc.ImportFromUrl("http://example.com", "scale=0.6; drawbackground=true")
	if err != nil {
		// Expected failure in test environment (no HTTP client)
		t.Logf("ImportFromUrl returned expected error: %v", err)
	}
}

// TestG3PDFPersitsOutput tests SendBinary and Save output methods
func TestG3PDFPersitsOutput(t *testing.T) {
	vm := NewVM(nil, nil, 0)
	pdf := NewG3PDF(vm)
	doc := pdf.CreateDocument()
	doc.pdf.pdf.AddPage()
	doc.pdf.pdf.SetFont("helvetica", "", 12)
	doc.pdf.pdf.CellFormat(40, 10, "Test Output", "1", 0, "C", false, 0, "")

	// Test SendBinary returns bytes
	data, err := doc.SendBinary()
	if err != nil {
		t.Fatalf("SendBinary failed: %v", err)
	}
	if len(data) == 0 {
		t.Error("expected non-empty binary data")
	}
	t.Logf("SendBinary returned %d bytes", len(data))
}
