package server

import (
	"net/http/httptest"
	"strings"
	"testing"
)

// ExecuteASPTest is a helper function to execute ASP code and return the output
func ExecuteASPTest(code string) (string, error) {
	// Create a mock HTTP request and response
	req := httptest.NewRequest("GET", "/test.asp", nil)
	w := httptest.NewRecorder()

	// Create executor
	config := &ASPProcessorConfig{
		RootDir:       "./www",
		ScriptTimeout: 30,
	}
	executor := NewASPExecutor(config)

	// Execute the code
	err := executor.Execute(code, w, req, "test-session")
	if err != nil {
		return "", err
	}

	// Get response body
	output := w.Body.String()
	return strings.TrimSpace(output), nil
}

// TestArithmeticOperators tests basic arithmetic operations
func TestArithmeticOperators(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		{`<% Response.Write 10 + 5 %>`, "15"},
		{`<% Response.Write 10 - 5 %>`, "5"},
		{`<% Response.Write 10 * 5 %>`, "50"},
		{`<% Response.Write 10 / 5 %>`, "2"},
	}

	for _, test := range tests {
		result, err := ExecuteASPTest(test.code)
		if err != nil {
			t.Errorf("Error executing code '%s': %v", test.code, err)
			continue
		}
		if result != test.expected {
			t.Errorf("Expected '%s' but got '%s' for code: %s", test.expected, result, test.code)
		}
	}
}

// TestIntDivisionAndMod tests integer division and modulo
func TestIntDivisionAndMod(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		{`<% Response.Write 17 \ 5 %>`, "3"},
		{`<% Response.Write 17 Mod 5 %>`, "2"},
		{`<% Response.Write 20 \ 3 %>`, "6"},
		{`<% Response.Write 20 Mod 3 %>`, "2"},
	}

	for _, test := range tests {
		result, err := ExecuteASPTest(test.code)
		if err != nil {
			t.Errorf("Error executing code '%s': %v", test.code, err)
			continue
		}
		if result != test.expected {
			t.Errorf("Expected '%s' but got '%s' for code: %s", test.expected, result, test.code)
		}
	}
}

// TestLogicalOperators tests And, Or, Not with boolean context
func TestLogicalOperators(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		// In VBScript, True = -1 and False = 0
		// True And True = -1 & -1 = -1 (True)
		// True And False = -1 & 0 = 0 (False)
		{`<% Response.Write CBool(True And True) %>`, "True"},
		{`<% Response.Write CBool(True And False) %>`, "False"},
		{`<% Response.Write CBool(False And False) %>`, "False"},
		{`<% Response.Write CBool(True Or False) %>`, "True"},
		{`<% Response.Write CBool(False Or False) %>`, "False"},
		{`<% Response.Write CBool(Not False) %>`, "True"},
	}

	for _, test := range tests {
		result, err := ExecuteASPTest(test.code)
		if err != nil {
			t.Errorf("Error executing code '%s': %v", test.code, err)
			continue
		}
		if result != test.expected {
			t.Errorf("Expected '%s' but got '%s' for code: %s", test.expected, result, test.code)
		}
	}
}

// TestBitwiseOperators tests bitwise And, Or, Not
func TestBitwiseOperators(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		{`<% Response.Write 12 And 10 %>`, "8"}, // 1100 & 1010 = 1000
		{`<% Response.Write 12 Or 10 %>`, "14"}, // 1100 | 1010 = 1110
		{`<% Response.Write Not 0 %>`, "-1"},    // ~0 = -1
		{`<% Response.Write 5 And 3 %>`, "1"},   // 0101 & 0011 = 0001
		{`<% Response.Write 5 Or 3 %>`, "7"},    // 0101 | 0011 = 0111
	}

	for _, test := range tests {
		result, err := ExecuteASPTest(test.code)
		if err != nil {
			t.Errorf("Error executing code '%s': %v", test.code, err)
			continue
		}
		if result != test.expected {
			t.Errorf("Expected '%s' but got '%s' for code: %s", test.expected, result, test.code)
		}
	}
}

// TestBooleanValues tests True and False constants
func TestBooleanValues(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		{`<% Dim x %><% x = True %><% Response.Write x %>`, "True"},
		{`<% Dim x %><% x = False %><% Response.Write x %>`, "False"},
		{`<% Response.Write CInt(True) %>`, "-1"},
		{`<% Response.Write CInt(False) %>`, "0"},
	}

	for _, test := range tests {
		result, err := ExecuteASPTest(test.code)
		if err != nil {
			t.Errorf("Error executing code '%s': %v", test.code, err)
			continue
		}
		if result != test.expected {
			t.Errorf("Expected '%s' but got '%s' for code: %s", test.expected, result, test.code)
		}
	}
}

// TestSpecialValues tests Null, Empty, Nothing
func TestSpecialValues(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		{`<% Dim x %><% x = Empty %><% Response.Write IsEmpty(x) %>`, "True"},
		{`<% Dim x %><% x = Null %><% Response.Write IsNull(x) %>`, "True"},
		{`<% Dim x %><% Set x = Nothing %><% Response.Write x Is Nothing %>`, "True"},
	}

	for _, test := range tests {
		result, err := ExecuteASPTest(test.code)
		if err != nil {
			t.Errorf("Error executing code '%s': %v", test.code, err)
			continue
		}
		if result != test.expected {
			t.Errorf("Expected '%s' but got '%s' for code: %s", test.expected, result, test.code)
		}
	}
}

// TestOperatorPrecedence tests operator precedence
func TestOperatorPrecedence(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		{`<% Response.Write 2 + 3 * 4 %>`, "14"},
		{`<% Response.Write 10 - 3 + 2 %>`, "9"},
		{`<% Response.Write 15 Mod 4 * 2 %>`, "6"},
	}

	for _, test := range tests {
		result, err := ExecuteASPTest(test.code)
		if err != nil {
			t.Errorf("Error executing code '%s': %v", test.code, err)
			continue
		}
		if result != test.expected {
			t.Errorf("Expected '%s' but got '%s' for code: %s", test.expected, result, test.code)
		}
	}
}

// TestNegativeNumbers tests operations with negative numbers
func TestNegativeNumbers(t *testing.T) {
	tests := []struct {
		code     string
		expected string
	}{
		{`<% Response.Write -10 + -5 %>`, "-15"},
		{`<% Response.Write -10 * -5 %>`, "50"},
		{`<% Response.Write -10 / 2 %>`, "-5"},
	}

	for _, test := range tests {
		result, err := ExecuteASPTest(test.code)
		if err != nil {
			t.Errorf("Error executing code '%s': %v", test.code, err)
			continue
		}
		if result != test.expected {
			t.Errorf("Expected '%s' but got '%s' for code: %s", test.expected, result, test.code)
		}
	}
}
