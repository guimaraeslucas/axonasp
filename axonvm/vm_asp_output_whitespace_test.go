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
	"os"
	"path/filepath"
	"testing"
)

// TestASPOutputSkipsIndentedLineBreakAfterPercentBlockEnd verifies IIS-compatible
// output suppression of one line break after %> when only horizontal whitespace
// appears before that line break.
func TestASPOutputSkipsIndentedLineBreakAfterPercentBlockEnd(t *testing.T) {
	source := "A<%= \"X\" %>   \t\r\nB"
	actual := runASPAndCollectOutput(t, source)
	if actual != "AXB" {
		t.Fatalf("unexpected output: got %q want %q", actual, "AXB")
	}
}

// TestASPOutputSkipsIndentedLineBreakAfterScriptServerEnd verifies the same
// suppression rule after </script> runat=server block termination.
func TestASPOutputSkipsIndentedLineBreakAfterScriptServerEnd(t *testing.T) {
	source := "<script runat=\"server\">Dim x : x = 1</script>  \t\nB"
	actual := runASPAndCollectOutput(t, source)
	if actual != "B" {
		t.Fatalf("unexpected output: got %q want %q", actual, "B")
	}
}

// TestASPOutputPreservesWhitespaceWhenNoLineBreakFollowsPercentBlockEnd verifies
// horizontal whitespace remains when no line break follows %>.
func TestASPOutputPreservesWhitespaceWhenNoLineBreakFollowsPercentBlockEnd(t *testing.T) {
	source := "A<%= \"X\" %>   B"
	actual := runASPAndCollectOutput(t, source)
	if actual != "AX   B" {
		t.Fatalf("unexpected output: got %q want %q", actual, "AX   B")
	}
}

func TestASPOutputStandaloneIncludeLineDoesNotEmitExtraBlankLine(t *testing.T) {
	rootDir := t.TempDir()
	includePath := filepath.Join(rootDir, "header.inc")
	if err := os.WriteFile(includePath, []byte("X"), 0o644); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	parentPath := filepath.Join(rootDir, "default.asp")
	source := "A\r\n<!--#include file=\"header.inc\"-->\r\nB"

	compiler := NewASPCompiler(source)
	compiler.SetSourceName(parentPath)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/default.asp")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "A\r\nXB" {
		t.Fatalf("unexpected output: got %q want %q", output.String(), "A\\r\\nXB")
	}
}

func TestASPOutputIncludeWithIndentedTrailingCRLFDoesNotEmitBlankLine(t *testing.T) {
	rootDir := t.TempDir()
	includePath := filepath.Join(rootDir, "header.inc")
	if err := os.WriteFile(includePath, []byte("X"), 0o644); err != nil {
		t.Fatalf("write include failed: %v", err)
	}

	parentPath := filepath.Join(rootDir, "default.asp")
	source := "<!--#include file=\"header.inc\"-->   \t\r\nB"

	compiler := NewASPCompiler(source)
	compiler.SetSourceName(parentPath)
	if err := compiler.Compile(); err != nil {
		t.Fatalf("compile failed: %v", err)
	}

	vm := NewVMFromCompiler(compiler)
	host := NewMockHost()
	host.Server().SetRootDir(rootDir)
	host.Server().SetRequestPath("/default.asp")
	var output bytes.Buffer
	host.SetOutput(&output)
	vm.SetHost(host)

	if err := vm.Run(); err != nil {
		t.Fatalf("vm run failed: %v", err)
	}
	host.Response().Flush()

	if output.String() != "XB" {
		t.Fatalf("unexpected output: got %q want %q", output.String(), "XB")
	}
}

func TestASPOutputSuppressesFormattingWhitespaceBeforeCodeBlock(t *testing.T) {
	source := "\r\n\t<% Dim x : x = 1 %>\r\nB"
	actual := runASPAndCollectOutput(t, source)
	if actual != "B" {
		t.Fatalf("unexpected output: got %q want %q", actual, "B")
	}
}

func TestASPOutputCollapsesWhitespaceBetweenConsecutiveCodeBlocks(t *testing.T) {
	source := "<% Dim x : x = 1 %>\r\n\r\n<% Dim y : y = 2 %>\r\nB"
	actual := runASPAndCollectOutput(t, source)
	if actual != "B" {
		t.Fatalf("unexpected output: got %q want %q", actual, "B")
	}
}
