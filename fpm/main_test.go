//go:build !windows

package main

import "testing"

func TestNormalizePoolSocketEndpoint(t *testing.T) {
	tests := []struct {
		name         string
		input        string
		wantEndpoint string
		wantPath     string
		wantUnix     bool
		wantErr      bool
	}{
		{
			name:         "unix endpoint with prefix",
			input:        "unix:/tmp/client.sock",
			wantEndpoint: "unix:/tmp/client.sock",
			wantPath:     "/tmp/client.sock",
			wantUnix:     true,
		},
		{
			name:         "absolute unix path without prefix",
			input:        "/tmp/client.sock",
			wantEndpoint: "unix:/tmp/client.sock",
			wantPath:     "/tmp/client.sock",
			wantUnix:     true,
		},
		{
			name:         "relative unix path without prefix",
			input:        "./run/client.sock",
			wantEndpoint: "unix:./run/client.sock",
			wantPath:     "./run/client.sock",
			wantUnix:     true,
		},
		{
			name:         "tcp endpoint untouched",
			input:        "127.0.0.1:9100",
			wantEndpoint: "127.0.0.1:9100",
			wantPath:     "",
			wantUnix:     false,
		},
		{
			name:    "empty endpoint returns error",
			input:   "",
			wantErr: true,
		},
		{
			name:    "invalid unix prefix path returns error",
			input:   "unix:",
			wantErr: true,
		},
	}

	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			gotEndpoint, gotPath, gotUnix, err := normalizePoolSocketEndpoint(tt.input)
			if tt.wantErr {
				if err == nil {
					t.Fatalf("expected error, got endpoint=%q path=%q isUnix=%t", gotEndpoint, gotPath, gotUnix)
				}
				return
			}
			if err != nil {
				t.Fatalf("unexpected error: %v", err)
			}
			if gotEndpoint != tt.wantEndpoint {
				t.Fatalf("endpoint mismatch: got %q want %q", gotEndpoint, tt.wantEndpoint)
			}
			if gotPath != tt.wantPath {
				t.Fatalf("socket path mismatch: got %q want %q", gotPath, tt.wantPath)
			}
			if gotUnix != tt.wantUnix {
				t.Fatalf("isUnix mismatch: got %t want %t", gotUnix, tt.wantUnix)
			}
		})
	}
}
