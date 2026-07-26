package caddy

import (
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"testing"

	"g3pix.com.br/axonasp/axonconfig"
)

func TestCustomIndexResolution(t *testing.T) {
	tempDir := t.TempDir()

	// Test 1: Priority order - index.asp, default.asp, home.asp, main.asp
	// Create all 4 files
	files := []string{"index.asp", "default.asp", "home.asp", "main.asp"}
	for _, f := range files {
		_ = os.WriteFile(filepath.Join(tempDir, f), []byte("<% Response.Write \"ok\" %>"), 0644)
	}

	res, ok := resolveDefaultASPPath(tempDir, "/")
	if !ok || res != "/index.asp" {
		t.Fatalf("expected /index.asp, got %s (ok=%v)", res, ok)
	}

	// Remove index.asp -> default.asp should be next
	_ = os.Remove(filepath.Join(tempDir, "index.asp"))
	res, ok = resolveDefaultASPPath(tempDir, "/")
	if !ok || res != "/default.asp" {
		t.Fatalf("expected /default.asp, got %s (ok=%v)", res, ok)
	}

	// Remove default.asp -> home.asp should be next
	_ = os.Remove(filepath.Join(tempDir, "default.asp"))
	res, ok = resolveDefaultASPPath(tempDir, "/")
	if !ok || res != "/home.asp" {
		t.Fatalf("expected /home.asp, got %s (ok=%v)", res, ok)
	}

	// Remove home.asp -> main.asp should be next
	_ = os.Remove(filepath.Join(tempDir, "home.asp"))
	res, ok = resolveDefaultASPPath(tempDir, "/")
	if !ok || res != "/main.asp" {
		t.Fatalf("expected /main.asp, got %s (ok=%v)", res, ok)
	}

	// Remove main.asp -> should return false so Caddy falls back
	_ = os.Remove(filepath.Join(tempDir, "main.asp"))
	res, ok = resolveDefaultASPPath(tempDir, "/")
	if ok || res != "/" {
		t.Fatalf("expected fallback (ok=false), got %s (ok=%v)", res, ok)
	}
}

func TestCloakSensitiveFiles(t *testing.T) {
	a := &AxonASP{
		SiteName: "test_cloak",
	}

	mockHandler := caddyhttpHandlerFunc(func(w http.ResponseWriter, r *http.Request) error {
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write([]byte("passed_to_next"))
		return nil
	})

	testCases := []struct {
		urlPath      string
		expectedCode int
	}{
		{"/global.asa", http.StatusNotFound},
		{"/GLOBAL.ASA", http.StatusNotFound},
		{"/MyInfo.xml", http.StatusNotFound},
		{"/myinfo.xml", http.StatusNotFound},
		{"/sub/global.asa", http.StatusNotFound},
		{"/sub/MyInfo.xml", http.StatusNotFound},
		{"/allowed.asp", http.StatusOK},
		{"/index.html", http.StatusOK},
	}

	for _, tc := range testCases {
		req := httptest.NewRequest("GET", tc.urlPath, nil)
		rec := httptest.NewRecorder()

		// Create dummy file for non-cloaked tests if needed
		_ = a.ServeHTTP(rec, req, mockHandler)

		if rec.Code != tc.expectedCode {
			t.Errorf("Path %s: expected status %d, got %d", tc.urlPath, tc.expectedCode, rec.Code)
		}
	}
}

func TestTempDirAlignmentAndSubfolders(t *testing.T) {
	a := &AxonASP{
		SiteName: "site_alpha",
	}

	req := httptest.NewRequest("GET", "http://site_alpha.local/", nil)
	siteTemp := a.resolveSiteTempDir(req)

	caddyTemp := caddyNativeTempDir()
	expectedPrefix := filepath.Join(caddyTemp, "axonasp", "site_alpha")
	if siteTemp != expectedPrefix {
		t.Fatalf("expected site temp dir %s, got %s", expectedPrefix, siteTemp)
	}

	a.setupSiteTempDir(siteTemp)

	cacheDir := filepath.Join(siteTemp, "cache")
	sessionsDir := filepath.Join(siteTemp, "sessions")

	if info, err := os.Stat(cacheDir); err != nil || !info.IsDir() {
		t.Errorf("expected cache subfolder at %s", cacheDir)
	}
	if info, err := os.Stat(sessionsDir); err != nil || !info.IsDir() {
		t.Errorf("expected sessions subfolder at %s", sessionsDir)
	}

	// Verify Viper override
	viperTemp := axonconfig.NewViper().GetString("global.temp_dir")
	if viperTemp != siteTemp {
		t.Errorf("expected Viper global.temp_dir to be %s, got %s", siteTemp, viperTemp)
	}
}

type caddyhttpHandlerFunc func(w http.ResponseWriter, r *http.Request) error

func (f caddyhttpHandlerFunc) ServeHTTP(w http.ResponseWriter, r *http.Request) error {
	return f(w, r)
}
