package caddy

import (
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"
	"unsafe"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
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

func TestAxonLiveEndpointHandling(t *testing.T) {
	tempDir := t.TempDir()

	a := &AxonASP{
		SiteName:      "test_axonlive",
		GlobalAsaPath: filepath.Join(tempDir, "global.asa"),
	}
	_ = os.WriteFile(filepath.Join(tempDir, "global.asa"), []byte(""), 0644)

	mockHandler := caddyhttpHandlerFunc(func(w http.ResponseWriter, r *http.Request) error {
		w.WriteHeader(http.StatusNotFound)
		return nil
	})

	// 1. GET /g3al/ -> 405 Method Not Allowed
	req := httptest.NewRequest("GET", "/g3al/", nil)
	rec := httptest.NewRecorder()
	_ = a.ServeHTTP(rec, req, mockHandler)
	if rec.Code != http.StatusMethodNotAllowed {
		t.Errorf("GET /g3al/: expected %d, got %d", http.StatusMethodNotAllowed, rec.Code)
	}

	// 2. POST /g3al/ without X-G3AxonLive header -> 400 Bad Request
	req = httptest.NewRequest("POST", "/g3al/", strings.NewReader(`{"sessionId":"s1","componentId":"c1","eventName":"click"}`))
	rec = httptest.NewRecorder()
	_ = a.ServeHTTP(rec, req, mockHandler)
	if rec.Code != http.StatusBadRequest {
		t.Errorf("POST /g3al/ missing header: expected %d, got %d", http.StatusBadRequest, rec.Code)
	}

	// 3. POST /g3al/ with header but unregistered session -> 404 Not Found
	req = httptest.NewRequest("POST", "/g3al/", strings.NewReader(`{"sessionId":"unregistered","componentId":"c1","eventName":"click"}`))
	req.Header.Set("X-G3AxonLive", "true")
	req.AddCookie(&http.Cookie{Name: "ASPSESSIONID", Value: "unregistered"})
	rec = httptest.NewRecorder()
	_ = a.ServeHTTP(rec, req, mockHandler)
	if rec.Code != http.StatusNotFound {
		t.Errorf("POST /g3al/ unregistered session: expected %d, got %d", http.StatusNotFound, rec.Code)
	}

	// 4. Register page and test valid AxonLive dispatch
	sessionID := "test-session-123"
	scriptURL := "/live.asp"
	aspContent := `<%
Dim AxonLive
Set AxonLive = Server.CreateObject("G3AXONLIVE")
If AxonLive.InitPage() Then
    Response.Write "{""success"":true}"
End If
%>`
	_ = os.WriteFile(filepath.Join(tempDir, "live.asp"), []byte(aspContent), 0644)

	// First execute live.asp once via GET to initialize scriptCache / VM if needed, or register directly
	axonvm.G3ALRegisterPage(sessionID, scriptURL)

	reqBody := `{"sessionId":"test-session-123","componentId":"btn1","eventName":"click"}`
	req = httptest.NewRequest("POST", "/g3al", strings.NewReader(reqBody))
	req.Header.Set("X-G3AxonLive", "true")
	req.Header.Set("Content-Type", "application/json")
	req.AddCookie(&http.Cookie{Name: "ASPSESSIONID", Value: sessionID})

	// Provision scriptCache if needed
	a.scriptCache = axonvm.NewScriptCache(axonvm.BytecodeCacheMemoryOnly, "", 64)
	a.vmPools = &vmPoolManager{pools: make(map[string]unsafe.Pointer)}
	a.application = asp.NewApplication()

	rec = httptest.NewRecorder()
	_ = a.ServeHTTP(rec, req, mockHandler)

	if rec.Code != http.StatusOK {
		t.Errorf("POST /g3al valid session: expected %d, got %d. Body: %s", http.StatusOK, rec.Code, rec.Body.String())
	}
}

