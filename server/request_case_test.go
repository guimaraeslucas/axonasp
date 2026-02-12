package server

import (
	"net/http/httptest"
	"testing"
)

func TestRequestQueryStringCaseInsensitivity(t *testing.T) {
	tests := []struct {
		name        string
		queryString string
		accessKey   string
		expected    string
	}{
		{
			name:        "Lowercase parameter, lowercase access",
			queryString: "iid=test123",
			accessKey:   "iid",
			expected:    "test123",
		},
		{
			name:        "Lowercase parameter, uppercase access",
			queryString: "iid=test456",
			accessKey:   "IID",
			expected:    "test456",
		},
		{
			name:        "Uppercase parameter, lowercase access",
			queryString: "IID=test789",
			accessKey:   "iid",
			expected:    "test789",
		},
		{
			name:        "Mixed case parameter, different case access",
			queryString: "iId=testABC",
			accessKey:   "IID",
			expected:    "testABC",
		},
		{
			name:        "Mixed case in URL, exact match access",
			queryString: "iID=testDEF",
			accessKey:   "iId",
			expected:    "testDEF",
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			// Create test HTTP request
			req := httptest.NewRequest("GET", "http://localhost:4050/test.asp?"+tt.queryString, nil)

			// Create Request object
			requestObj := NewRequestObject()

			// Populate QueryString from HTTP request (simulating executor initialization)
			for key, values := range req.URL.Query() {
				if len(values) > 0 {
					requestObj.QueryString.Add(key, values[0])
				}
			}

			// Test access with different case
			result := requestObj.QueryString.Get(tt.accessKey)

			if result == nil || toString(result) != tt.expected {
				t.Errorf("Expected '%s', got '%v'", tt.expected, result)
			}

			// Also test through CallMethod (Request("iId"))
			methodResult, _ := requestObj.CallMethod(tt.accessKey)
			if methodResult == nil || toString(methodResult) != tt.expected {
				t.Errorf("CallMethod: Expected '%s', got '%v'", tt.expected, methodResult)
			}
		})
	}
}
