package main

import (
	"go-asp/asp"
	"fmt"
	"log"
	"net/http"
	"os"
	"path/filepath"
	"runtime/debug"
	"strconv"
	"strings"
	"time"

	"github.com/joho/godotenv"
)

// Configuration variables (defaults can be overridden by .env)
var (
	Port            = "4050"
	RootDir         = "./www"
	DefaultTimezone = "America/Sao_Paulo"
	DefaultPage     = "default.asp"
	ScriptTimeout   = 30 // in seconds

)

// DummyResponseWriter for background events
type DummyResponseWriter struct{}

func (d *DummyResponseWriter) Header() http.Header        { return http.Header{} }
func (d *DummyResponseWriter) Write([]byte) (int, error)  { return 0, nil }
func (d *DummyResponseWriter) WriteHeader(statusCode int) {}

func init() {
	// Load .env file
	err := godotenv.Load()
	if err != nil {
		fmt.Println("Info: No .env file found, using defaults or system environment.")
	}

	// Override configuration from Environment Variables if present
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

	// Set timezone for the application
	os.Setenv("TZ", DefaultTimezone)
}

func main() {
	// Load Global.asa
	globalPath := filepath.Join(RootDir, "global.asa")
	if content, err := os.ReadFile(globalPath); err == nil {
		fmt.Println("Info: Loading global.asa...")
		code := asp.ParseGlobalASA(string(content))
		asp.AppState.GlobalASACode = code

		// Trigger Application_OnStart
		dummyReq, _ := http.NewRequest("GET", "/", nil)
		dummyCtx := asp.NewExecutionContext(&DummyResponseWriter{}, dummyReq, RootDir)
		asp.RunGlobalEvent("Application_OnStart", dummyCtx)
	}

	// Start Session Scavenger
	go SessionScavenger()

	http.HandleFunc("/", handleRequest)

	fmt.Printf("Starting G3pix AxonASP on http://localhost:%s\n", Port)
	fmt.Printf("Serving files from %s\n", RootDir)

	err := http.ListenAndServe(":"+Port, nil)
	if err != nil {
		log.Fatal(err)
	}
}

func SessionScavenger() {
	// Check every minute
	ticker := time.NewTicker(1 * time.Minute)
	for range ticker.C {
		now := time.Now()
		asp.SessionStorage.Range(func(key, value interface{}) bool {
			s := value.(*asp.Session)
			// Default check: 20 mins or custom
			limit := time.Duration(s.Timeout) * time.Minute

			if !s.Abandoned && now.Sub(s.LastAccessed) > limit {
				// Expired: Trigger Session_OnEnd
				// We need a context with this session
				dummyReq, _ := http.NewRequest("GET", "/", nil)
				// We need to inject the EXPIRED session into the context,
				// but NewExecutionContext gets a NEW session if cookie missing/invalid.
				// We must manually construct Context to point to this session.

				dummyCtx := &asp.ExecutionContext{
					Response:    &DummyResponseWriter{},
					Request:     dummyReq,
					Session:     s, // Use the expiring session
					Application: asp.AppState,
					Variables:   make(map[string]interface{}),
					ResponseState: &asp.ResponseState{
						Buffer: true,
					},
					RootDir: RootDir,
				}

				asp.RunGlobalEvent("Session_OnEnd", dummyCtx)

				// Remove
				asp.SessionStorage.Delete(s.ID)
			}
			return true
		})
	}
}

func handleRequest(w http.ResponseWriter, r *http.Request) {
	path := r.URL.Path
	if path == "/" {
		path = "/" + DefaultPage
	}

	fullPath := filepath.Join(RootDir, path)

	// Security check
	if !strings.HasPrefix(filepath.Clean(fullPath), filepath.Clean(RootDir)) {
		http.Error(w, "AxonASP: Forbidden", http.StatusForbidden)
		return
	}

	info, err := os.Stat(fullPath)
	if os.IsNotExist(err) {
		http.Error(w, "AxonASP: 404 page not found ", http.StatusNotFound)
		//http.NotFound(w, r)
		return
	}
	if info.IsDir() {
		// try default
		fullPath = filepath.Join(fullPath, DefaultPage)
		if _, err := os.Stat(fullPath); os.IsNotExist(err) {
			http.Error(w, "AxonASP: 404 page not found ", http.StatusNotFound)
			//http.NotFound(w, r)
			return
		}
	}

	// Serve Static if not ASP
	if !strings.HasSuffix(strings.ToLower(fullPath), ".asp") {
		http.ServeFile(w, r, fullPath)
		return
	}

	// Process ASP
	content, err := os.ReadFile(fullPath)
	if err != nil {
		http.Error(w, "AxonASP: error reading file", http.StatusInternalServerError)
		return
	}

	// 1. Initialize Context
	ctx := asp.NewExecutionContext(w, r, RootDir)

	// Handle Includes (Server-Side Includes)
	contentStr := string(content)
	contentStr, err = asp.ProcessIncludes(contentStr, filepath.Dir(fullPath), RootDir)
	if err != nil {
		fmt.Printf("Warning: Include processing error in %s: %v\n", fullPath, err)
		ctx.Write(fmt.Sprintf("Include error: %s: %v<br>", fullPath, err))
	}

	// Trigger Session_OnStart if new
	if ctx.Session.IsNew {
		asp.RunGlobalEvent("Session_OnStart", ctx)
	}

	// 2. Parse Code
	tokens := asp.ParseRaw(contentStr)

	// 3. Prepare Engine
	engine := asp.Prepare(tokens)

	// 4. Run
	// (Recover from panics in interpreter to avoid crashing server)
	defer func() {
		if r := recover(); r != nil {
			fmt.Printf("Runtime Error in %s: %v\n", path, r)

			// Check for DEBUG_ASP_CODE variable
			isDebug := false
			if val, ok := ctx.Variables["debug_asp_code"]; ok {
				// Check string "TRUE" or boolean true
				sVal := fmt.Sprintf("%v", val)
				if strings.ToUpper(sVal) == "TRUE" {
					isDebug = true
				}
			}

			// Format Error Message
			line := engine.CurrentLine

			ctx.Write("<br><hr style='border-top: 1px dashed red;'>")
			ctx.Write("<div style='color: red; font-family: monospace; background: #ffe6e6; padding: 10px; border: 1px solid red;'>")

			if isDebug {
				// Detailed Error
				stack := string(debug.Stack())
				// Sanitize stack for HTML
				stack = strings.ReplaceAll(stack, "<", "&lt;")
				stack = strings.ReplaceAll(stack, ">", "&gt;")

				ctx.Write("<strong>G3pix AxonASP panic</strong><br>")
				ctx.Write(fmt.Sprintf("Error: %v<br>", r))
				ctx.Write(fmt.Sprintf("Line: %d<br>", line))
				ctx.Write(fmt.Sprintf("<pre>%s</pre>", stack))
			} else {
				// Simple Error
				ctx.Write("<strong>G3pix AxonASP error</strong><br>")
				ctx.Write(fmt.Sprintf("Description: %v<br>", r))
				ctx.Write(fmt.Sprintf("Line: %d", line))
			}

			ctx.Write("</div>")
			ctx.Flush()
		}
	}()

	engine.Run(ctx)

	// 5. Send Output
	ctx.Flush()
}
