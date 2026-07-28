/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 */
package axonvm

import (
	"strings"
	"testing"
)

func TestServerVariablesAndClientCertificateDispatchJScript(t *testing.T) {
	host := NewMockHost()
	host.Request().ServerVars.Add("SCRIPT_NAME", "/app/index.asp")
	host.Request().ClientCertificate.Add("Subject", "CN=TestCert")

	source := `<script runat="server" language="JScript">
		// 1. Direct access with index (1) on ServerVariables
		Response.Write("1:" + Request.ServerVariables("SCRIPT_NAME")(1) + "\n");

		// 2. Direct string coercion on ServerVariables
		Response.Write("2:" + String(Request.ServerVariables("SCRIPT_NAME")) + "\n");

		// 3. Count on ServerVariables collection item
		Response.Write("3:" + Request.ServerVariables("SCRIPT_NAME").Count + "\n");

		// 4. Direct access with index (1) on ClientCertificate
		Response.Write("4:" + Request.ClientCertificate("Subject")(1) + "\n");

		// 5. Direct string coercion on ClientCertificate
		Response.Write("5:" + String(Request.ClientCertificate("Subject")) + "\n");

		// 6. Count on ClientCertificate collection item
		Response.Write("6:" + Request.ClientCertificate("Subject").Count + "\n");

		// 7. Non-existent keys
		Response.Write("7:" + String(Request.ServerVariables("NON_EXISTENT")) + "\n");
		Response.Write("8:" + String(Request.ClientCertificate("NON_EXISTENT")) + "\n");
	</script>`

	out := runASPSourceForTestWithHost(t, source, host)
	lines := strings.Split(strings.TrimSpace(out), "\n")

	expected := []string{
		"1:/app/index.asp",
		"2:/app/index.asp",
		"3:1",
		"4:CN=TestCert",
		"5:CN=TestCert",
		"6:1",
		"7:undefined",
		"8:undefined",
	}

	if len(lines) != len(expected) {
		t.Fatalf("expected %d output lines, got %d:\n%s", len(expected), len(lines), out)
	}

	for i, exp := range expected {
		if lines[i] != exp {
			t.Errorf("line %d: expected %q, got %q", i+1, exp, lines[i])
		}
	}
}

func TestServerVariablesAndClientCertificateDispatchVBScript(t *testing.T) {
	host := NewMockHost()
	host.Request().ServerVars.Add("SCRIPT_NAME", "/app/index.asp")
	host.Request().ClientCertificate.Add("Subject", "CN=TestCert")

	source := `<%
		' 1. Direct access with index (1) on ServerVariables
		Response.Write "1:" & Request.ServerVariables("SCRIPT_NAME")(1) & vbCrLf

		' 2. Direct string coercion on ServerVariables
		Response.Write "2:" & Request.ServerVariables("SCRIPT_NAME") & vbCrLf

		' 3. Count on ServerVariables collection item
		Response.Write "3:" & Request.ServerVariables("SCRIPT_NAME").Count & vbCrLf

		' 4. Direct access with index (1) on ClientCertificate
		Response.Write "4:" & Request.ClientCertificate("Subject")(1) & vbCrLf

		' 5. Direct string coercion on ClientCertificate
		Response.Write "5:" & Request.ClientCertificate("Subject") & vbCrLf

		' 6. Count on ClientCertificate collection item
		Response.Write "6:" & Request.ClientCertificate("Subject").Count & vbCrLf

		' 7. Non-existent keys
		Response.Write "7:" & Request.ServerVariables("NON_EXISTENT") & vbCrLf
		Response.Write "8:" & Request.ClientCertificate("NON_EXISTENT") & vbCrLf
	%>`

	out := runASPSourceForTestWithHost(t, source, host)
	lines := strings.Split(strings.TrimSpace(out), "\r\n")

	expected := []string{
		"1:/app/index.asp",
		"2:/app/index.asp",
		"3:1",
		"4:CN=TestCert",
		"5:CN=TestCert",
		"6:1",
		"7:",
		"8:",
	}

	if len(lines) != len(expected) {
		t.Fatalf("expected %d output lines, got %d:\n%s", len(expected), len(lines), out)
	}

	for i, exp := range expected {
		if lines[i] != exp {
			t.Errorf("line %d: expected %q, got %q", i+1, exp, lines[i])
		}
	}
}
