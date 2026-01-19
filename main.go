package main

import (
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"runtime/debug"
	"strconv"
	"strings"
	"time"

	"go-asp/server"

	"github.com/joho/godotenv"
)

// Configuration variables
var (
	Port            = "4050"
	RootDir         = "./www"
	DefaultTimezone = "America/Sao_Paulo"
	DefaultPage     = "default.asp"
	ScriptTimeout   = 30 // in seconds
	DebugASP        = false
)

func init() {
	// Load .env file
	err := godotenv.Load()
	if err != nil {
		fmt.Println("Info: No .env file found, using defaults or system environment.")
	}

	// Override configuration from Environment Variables
	if val := os.Getenv("SERVER_PORT"); val != "" {
		Port = val
	}
	if val := os.Getenv("WEB_ROOT"); val != "" {
		RootDir = val
	}
	if val := os.Getenv("TIMEZONE"); val != "" {
		DefaultTimezone = val
	}
	if val := os.Getenv("DEFAULT_PAGE"); val != "" {
		DefaultPage = val
	}
	if val := os.Getenv("SCRIPT_TIMEOUT"); val != "" {
		if i, err := strconv.Atoi(val); err == nil {
			ScriptTimeout = i
		}
	}
	if val := os.Getenv("DEBUG_ASP"); val == "TRUE" {
		DebugASP = true
	}

	// Set timezone
	os.Setenv("TZ", DefaultTimezone)
}

func main() {
	http.HandleFunc("/", handleRequest)

	// Initialize session manager and start cleanup routine
	sessionManager := server.GetSessionManager()
	sessionManager.StartCleanupRoutine(15 * time.Minute) // Cleanup every 15 minutes

	// Load Global.asa file
	globalASAManager := server.GetGlobalASAManager()
	err := globalASAManager.LoadGlobalASA(RootDir)
	if err != nil {
		fmt.Printf("Warning: Failed to load Global.asa: %v\n", err)
	}

	// Check if events are loaded
	fmt.Printf("Application_OnStart defined: %v\n", globalASAManager.HasApplicationOnStart())
	fmt.Printf("Session_OnStart defined: %v\n", globalASAManager.HasSessionOnStart())

	// Execute Application_OnStart if defined
	if globalASAManager.HasApplicationOnStart() {
		// Create temporary executor for Application_OnStart
		processor := server.NewASPProcessor(&server.ASPProcessorConfig{
			RootDir:       RootDir,
			ScriptTimeout: ScriptTimeout,
			DebugASP:      DebugASP,
		})
		executor := server.NewASPExecutor(processor.GetConfig())

		// Create a dummy response writer for Application_OnStart
		// Since it's not tied to a specific request, we can use a no-op writer
		dummyWriter := NewDummyResponseWriter()

		// Create a dummy request
		dummyRequest := &http.Request{
			Header: make(http.Header),
		}

		// Create execution context for Application_OnStart
		ctx := server.NewExecutionContext(dummyWriter, dummyRequest, "app_startup", time.Duration(ScriptTimeout)*time.Second)
		ctx.RootDir = RootDir

		// Set the context in executor
		executor.SetContext(ctx)

		if err := globalASAManager.ExecuteApplicationOnStart(executor, ctx); err != nil {
			fmt.Printf("Error in Application_OnStart: %v\n", err)
		}
	}

	fmt.Printf("Starting G3pix AxonASP on http://localhost:%s\n", Port)
	fmt.Printf("Serving files from %s\n", RootDir)
	if DebugASP {
		fmt.Println("[DEBUG] DEBUG_ASP mode is enabled")
	}

	err = http.ListenAndServe(":"+Port, nil)
	if err != nil {
		log.Fatal(err)
	}
}

// DummyResponseWriter is a no-op response writer for Application_OnStart
type DummyResponseWriter struct{}

func (w *DummyResponseWriter) Header() http.Header {
	return make(http.Header)
}

func (w *DummyResponseWriter) Write([]byte) (int, error) {
	return 0, nil
}

func (w *DummyResponseWriter) WriteHeader(statusCode int) {
}

func NewDummyResponseWriter() *DummyResponseWriter {
	return &DummyResponseWriter{}
}

func handleRequest(w http.ResponseWriter, r *http.Request) {
	path := r.URL.Path
	if path == "/" {
		path = "/" + DefaultPage
	}

	fullPath := filepath.Join(RootDir, path)

	// Security check: prevent directory traversal
	cleanPath := filepath.Clean(fullPath)
	cleanRoot := filepath.Clean(RootDir)
	if !strings.HasPrefix(cleanPath, cleanRoot) {
		http.Error(w, "AxonASP: Forbidden", http.StatusForbidden)
		return
	}

	// Check if file exists
	info, err := os.Stat(fullPath)
	if os.IsNotExist(err) {
		http.Error(w, "AxonASP: 404 page not found", http.StatusNotFound)
		return
	}

	// If it's a directory, try to serve the default page
	if info.IsDir() {
		fullPath = filepath.Join(fullPath, DefaultPage)
		if _, err := os.Stat(fullPath); os.IsNotExist(err) {
			http.Error(w, "AxonASP: 404 page not found", http.StatusNotFound)
			return
		}
	}

	// Serve static files if not ASP
	if !strings.HasSuffix(strings.ToLower(fullPath), ".asp") {
		http.ServeFile(w, r, fullPath)
		return
	}

	// Process ASP file
	content, err := os.ReadFile(fullPath)
	if err != nil {
		http.Error(w, "AxonASP: error reading file", http.StatusInternalServerError)
		return
	}

	// Recover from panics to avoid crashing server
	defer func() {
		if r := recover(); r != nil {
			fmt.Printf("Runtime panic in %s: %v\n", path, r)

			// Check if debug mode is enabled
			isDebug := os.Getenv("DEBUG_ASP") == "TRUE"

			w.Header().Set("Content-Type", "text/html; charset=utf-8")
			w.WriteHeader(http.StatusInternalServerError)

			fmt.Fprintf(w, "<br><hr style='border-top: 1px dashed red;'>\n")
			fmt.Fprintf(w, "<div style='color: red; font-family: monospace; background: #ffe6e6; padding: 10px; border: 1px solid red;'>\n")

			if isDebug {
				// Detailed error output with stack trace
				stack := string(debug.Stack())
				stack = strings.ReplaceAll(stack, "<", "&lt;")
				stack = strings.ReplaceAll(stack, ">", "&gt;")

				fmt.Fprintf(w, "<strong>G3pix AxonASP panic</strong><br>\n")
				fmt.Fprintf(w, "Error: %v<br>\n", r)
				fmt.Fprintf(w, "<pre>%s</pre>\n", stack)
			} else {
				// Simple error output
				fmt.Fprintf(w, "<strong>G3pix AxonASP error</strong><br>\n")
				fmt.Fprintf(w, "Description: %v<br>\n", r)
			}

			fmt.Fprintf(w, "</div>\n")
		}
	}()

	// Create ASP processor and execute
	processor := server.NewASPProcessor(&server.ASPProcessorConfig{
		RootDir:       RootDir,
		ScriptTimeout: ScriptTimeout,
		DebugASP:      DebugASP,
	})

	err = processor.ExecuteASPFile(string(content), fullPath, w, r)
	if err != nil {
		fmt.Printf("ASP processing error in %s: %v\n", path, err)
		http.Error(w, fmt.Sprintf("AxonASP: %v", err), http.StatusInternalServerError)
	}
}
