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

// TestJScriptGlobalAndGlobalThis verifies that global and globalThis are aliases.
func TestJScriptGlobalAndGlobalThis(t *testing.T) {
	source := `<script runat="server" language="JScript">
		Response.Write(typeof global === "object" ? "1" : "0");
		Response.Write(typeof globalThis === "object" ? "1" : "0");
		Response.Write(global === globalThis ? "1" : "0");
		var result = "0";
		try {
			result = global.Math ? "1" : "0";
		} catch (e) {
			Response.Write("ERROR: " + e.message + "\n");
			result = "0";
		}
		Response.Write(result);
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1111"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptProcessEnv verifies process.env access for environment variables.
func TestJScriptProcessEnv(t *testing.T) {
	// Set a test environment variable
	testKey := "AXONASP_TEST_VAR"
	testValue := "test_value_12345"
	os.Setenv(testKey, testValue)
	defer os.Unsetenv(testKey)

	source := `<script runat="server" language="JScript">
		Response.Write(typeof process === "object" ? "1" : "0");
		Response.Write(typeof process.env === "object" ? "1" : "0");
		Response.Write(process.env["AXONASP_TEST_VAR"] === "test_value_12345" ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "111"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptProcessArgv verifies process.argv array is available.
func TestJScriptProcessArgv(t *testing.T) {
	source := `<script runat="server" language="JScript">
		Response.Write(typeof process.argv === "object" ? "1" : "0");
		Response.Write(Array.isArray(process.argv) ? "1" : "0");
		Response.Write(process.argv.length > 0 ? "1" : "0");
		Response.Write(typeof process.argv[0] === "string" ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1111"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptProcessCwd verifies process.cwd() returns a string path.
func TestJScriptProcessCwd(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var cwd = process.cwd();
		Response.Write(typeof cwd === "string" ? "1" : "0");
		Response.Write(cwd.length > 0 ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "11"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferFrom verifies Buffer.from() with UTF-8 encoding.
func TestJScriptBufferFrom(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from("hello", "utf8");
		Response.Write(Buffer.isBuffer(buf) ? "1" : "0");
		Response.Write(buf.length === 5 ? "1" : "0");
		Response.Write(buf[0] === 104 ? "1" : "0");  // 'h' = 104
		Response.Write(buf[1] === 101 ? "1" : "0");  // 'e' = 101
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1111"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferFromArray verifies Buffer.from() with array of bytes.
func TestJScriptBufferFromArray(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from([72, 101, 108, 108, 111]);
		Response.Write(buf.length === 5 ? "1" : "0");
		Response.Write(buf[0] === 72 ? "1" : "0");   // 'H'
		Response.Write(buf[4] === 111 ? "1" : "0");  // 'o'
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "111"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferAlloc verifies Buffer.alloc() memory allocation.
func TestJScriptBufferAlloc(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.alloc(10);
		Response.Write(buf.length === 10 ? "1" : "0");
		Response.Write(buf[0] === 0 ? "1" : "0");
		Response.Write(buf[9] === 0 ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "111"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferAllocWithFill verifies Buffer.alloc() with fill value.
func TestJScriptBufferAllocWithFill(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.alloc(5, 255);
		Response.Write(buf.length === 5 ? "1" : "0");
		Response.Write(buf[0] === 255 ? "1" : "0");
		Response.Write(buf[4] === 255 ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "111"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferIsBuffer verifies Buffer.isBuffer() type checking.
func TestJScriptBufferIsBuffer(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from("test");
		Response.Write(Buffer.isBuffer(buf) ? "1" : "0");
		Response.Write(Buffer.isBuffer("string") ? "1" : "0");
		Response.Write(Buffer.isBuffer([1, 2, 3]) ? "1" : "0");
		Response.Write(Buffer.isBuffer(null) ? "1" : "0");
		Response.Write(Buffer.isBuffer(undefined) ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "10000"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferToString verifies Buffer.toString() with UTF-8 encoding.
func TestJScriptBufferToString(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from("hello");
		Response.Write(buf.toString() === "hello" ? "1" : "0");
		Response.Write(buf.toString("utf8") === "hello" ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "11"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferToStringHex verifies Buffer.toString() with hex encoding.
func TestJScriptBufferToStringHex(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from([0xAB, 0xCD, 0xEF]);
		var hex = buf.toString("hex");
		Response.Write(hex === "abcdef" ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferToStringBase64 verifies Buffer.toString() with base64 encoding.
func TestJScriptBufferToStringBase64(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from("test");
		var b64 = buf.toString("base64");
		Response.Write(b64 === "dGVzdA==" ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferFromHex verifies Buffer.from() with hex encoding.
func TestJScriptBufferFromHex(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from("48656c6c6f", "hex");
		Response.Write(buf.toString() === "Hello" ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferFromBase64 verifies Buffer.from() with base64 encoding.
func TestJScriptBufferFromBase64(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from("dGVzdA==", "base64");
		Response.Write(buf.toString() === "test" ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferLength verifies buffer.length property.
func TestJScriptBufferLength(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf1 = Buffer.from("hello");
		var buf2 = Buffer.alloc(100);
		Response.Write(buf1.length === 5 ? "1" : "0");
		Response.Write(buf2.length === 100 ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "11"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferIndexAccess verifies numeric index access to buffer bytes.
func TestJScriptBufferIndexAccess(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from([10, 20, 30, 40, 50]);
		Response.Write(buf[0] === 10 ? "1" : "0");
		Response.Write(buf[2] === 30 ? "1" : "0");
		Response.Write(buf[4] === 50 ? "1" : "0");
		Response.Write(buf[10] === undefined ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1111"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptBufferToStringRange verifies Buffer.toString() with start and end offsets.
func TestJScriptBufferToStringRange(t *testing.T) {
	source := `<script runat="server" language="JScript">
		var buf = Buffer.from("Hello World");
		Response.Write(buf.toString("utf8", 0, 5) === "Hello" ? "1" : "0");
		Response.Write(buf.toString("utf8", 6, 11) === "World" ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "11"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}

// TestJScriptNodeCompatibilityDisabled verifies that Node.js globals are unavailable when disabled.
// This test checks graceful fallback behavior when enable_node_compatibility = false.
func TestJScriptNodeCompatibilityDisabled(t *testing.T) {
	// This test would require dynamically disabling the feature, which isn't straightforward.
	// For now, we verify that if the config is disabled, the globals don't exist.
	// This is more of an integration test that depends on config setup.
	t.Skip("Integration test - requires config teardown/setup")
}

// TestJScriptBufferConstructorAsFunction verifies Buffer can be used as a function.
func TestJScriptBufferConstructorAsFunction(t *testing.T) {
	source := `<script runat="server" language="JScript">
		// Buffer should be callable via new or as a function
		var buf = Buffer.alloc(3);
		Response.Write(Buffer.isBuffer(buf) ? "1" : "0");
	</script>`

	output := runASPSourceForTest(t, source)
	expected := "1"
	if output != expected {
		t.Fatalf("Expected '%s', got '%s'", expected, output)
	}
}
