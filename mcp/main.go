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
//Use go install github.com/josephspurrier/goversioninfo/cmd/goversioninfo@latest
//Then run "go generate" in the project root to embed version info into the executable
//You need to specify -64=false/-arm=true if you're trying to create an 32-bit or ARM windows binary, this is required by the new version of golang
//go:generate goversioninfo -icon=icon_mcp.ico -64=true
package main

import (
	"bufio"
	"context"
	"flag"
	"fmt"
	"os"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
	"github.com/fsnotify/fsnotify"
	"github.com/joho/godotenv"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/sahilm/fuzzy"
	"github.com/spf13/viper"
)

// DocEntry represents a single AxonASP built-in function or library.
type DocEntry struct {
	Title        string
	Keywords     string
	Description  string
	Observations string
	Syntax       string
}

// Global state for in-memory caching and hot-reloading
var (
	Version       = "0.0.0.0"
	mu            sync.RWMutex
	cachedDocs    []DocEntry
	lastModTime   time.Time
	docFilePath   = "mcp/docs.md"
	styleFilePath = "mcp/aspcodingstyle.md"
	mcpMode       = "stdio"
	mcpSSEPort    = 8000
)

// init loads environment variables and applies TOML-based configuration through Viper.
func init() {
	_ = godotenv.Load()
	loadMCPConfig()
}

// loadMCPConfig loads and applies MCP settings from config/axonasp.toml using Viper.
func loadMCPConfig() {
	v := axonconfig.NewViper()
	if strings.TrimSpace(v.ConfigFileUsed()) == "" {
		fmt.Fprintf(os.Stderr, "[G3pix AxonASP MCP] Warning: failed to read configuration file, using defaults\n")
	}
	applyMCPConfigValues(v)

	axonconfig.EnableWatchIfConfigured(v, func(fsnotify.Event) {
		applyMCPConfigValues(v)
	})
}

// applyMCPConfigValues applies the active MCP settings from the loaded Viper instance.
func applyMCPConfigValues(v *viper.Viper) {
	if mode := strings.ToLower(strings.TrimSpace(v.GetString("mcp.mcp_mode"))); mode != "" {
		switch mode {
		case "stdio", "sse":
			mcpMode = mode
		default:
			fmt.Fprintf(os.Stderr, "[G3pix AxonASP MCP] Warning: invalid mcp.mcp_mode '%s', using stdio\n", mode)
			mcpMode = "stdio"
		}
	}

	if port := v.GetInt("mcp.mcp_sse_port"); port > 0 {
		mcpSSEPort = port
	}

	if docsPath := strings.TrimSpace(v.GetString("mcp.mcp_docs")); docsPath != "" {
		docFilePath = docsPath
	}
}

// LoadDocsFromMarkdown parses the markdown file into a slice of DocEntry structs.
func LoadDocsFromMarkdown(filePath string) ([]DocEntry, error) {
	file, err := os.Open(filePath)
	if err != nil {
		return nil, err
	}
	defer file.Close()

	var docs []DocEntry
	var currentDoc *DocEntry
	inCodeBlock := false

	scanner := bufio.NewScanner(file)
	for scanner.Scan() {
		line := scanner.Text()
		trimmedLine := strings.TrimSpace(line)

		// Detect new entry
		if strings.HasPrefix(trimmedLine, "## ") {
			if currentDoc != nil {
				docs = append(docs, *currentDoc)
			}
			currentDoc = &DocEntry{
				Title: strings.TrimPrefix(trimmedLine, "## "),
			}
			continue
		}

		if currentDoc == nil {
			continue
		}

		// Parse metadata fields
		if strings.HasPrefix(trimmedLine, "**Keywords:**") {
			currentDoc.Keywords = strings.TrimSpace(strings.TrimPrefix(trimmedLine, "**Keywords:**"))
			continue
		}
		if strings.HasPrefix(trimmedLine, "**Description:**") {
			currentDoc.Description = strings.TrimSpace(strings.TrimPrefix(trimmedLine, "**Description:**"))
			continue
		}
		if strings.HasPrefix(trimmedLine, "**Observations:**") {
			currentDoc.Observations = strings.TrimSpace(strings.TrimPrefix(trimmedLine, "**Observations:**"))
			continue
		}

		// Handle code blocks for syntax
		if strings.HasPrefix(trimmedLine, "```") {
			inCodeBlock = !inCodeBlock
			continue
		}

		if inCodeBlock {
			currentDoc.Syntax += line + "\n"
		}
	}

	if currentDoc != nil {
		docs = append(docs, *currentDoc)
	}

	if err := scanner.Err(); err != nil {
		return nil, err
	}

	return docs, nil
}

// reloadDocsIfNeeded checks if docs.md was modified and reloads it into memory.
func reloadDocsIfNeeded() error {
	info, err := os.Stat(docFilePath)
	if err != nil {
		return fmt.Errorf("could not stat %s: %w", docFilePath, err)
	}

	// If file was modified after our last load, trigger a reload
	if info.ModTime().After(lastModTime) {
		mu.Lock()
		defer mu.Unlock()

		// Double-check inside lock to prevent race conditions
		if info.ModTime().After(lastModTime) {
			docs, err := LoadDocsFromMarkdown(docFilePath)
			if err != nil {
				return fmt.Errorf("failed to parse markdown: %w", err)
			}
			cachedDocs = docs
			lastModTime = info.ModTime()
			// Firing a log to stderr (so it doesn't break stdout MCP communication)
			fmt.Fprintf(os.Stderr, "[G3pix AxonASP MCP] Reloaded %d functions from %s\n", len(docs), docFilePath)
		}
	}
	return nil
}

// getSearchableStrings generates a slice of strings for the fuzzy matcher.
func getSearchableStrings(docs []DocEntry) []string {
	var list []string
	for _, doc := range docs {
		// Combine Title and Keywords to increase match accuracy
		list = append(list, doc.Title+" "+doc.Keywords)
	}
	return list
}

// SearchHandler executes the fuzzy search logic and returns the Markdown snippet.
func searchHandler(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, ok := request.Params.Arguments.(map[string]interface{})
	if !ok {
		return mcp.NewToolResultError("Invalid arguments format. Expected a map."), nil
	}

	query, ok := args["query"].(string)
	if !ok || strings.TrimSpace(query) == "" {
		return mcp.NewToolResultError("Argument 'query' is required and must be a non-empty string."), nil
	}

	// Ensure our in-memory docs are up-to-date before searching
	if err := reloadDocsIfNeeded(); err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("Database error: %v", err)), nil
	}

	mu.RLock()
	docsToSearch := cachedDocs
	mu.RUnlock()

	searchList := getSearchableStrings(docsToSearch)
	matches := fuzzy.Find(query, searchList)

	if len(matches) == 0 {
		return mcp.NewToolResultText(fmt.Sprintf("No documentation found for: '%s'. Try different keywords. You should ensure your query is short using the least number os keywords. Eg.: g3db select", query)), nil
	}

	// Build the response in Markdown format
	var responseBuilder strings.Builder
	responseBuilder.WriteString(fmt.Sprintf("The 3 best matches for your search '%s'. You must carefully read all observations and implementation parameters before proceeding. ALWAYS prioritize existing server-side implementations over recreating Classic ASP code from scratch; for example, strictly use the native G3JSON object for JSON manipulation rather than raw parsing or custom ASP classes. Furthermore, you must adhere to standard Classic ASP coding patterns by avoiding single-line syntax where distinct lines are required, and always explicitly closing conditional blocks with 'End If':\n\n", query))

	// Return top 3 matches maximum
	limit := 3
	if len(matches) < 3 {
		limit = len(matches)
	}

	for i := 0; i < limit; i++ {
		doc := docsToSearch[matches[i].Index]
		responseBuilder.WriteString(fmt.Sprintf("### Option %d: %s\n", i+1, doc.Title))
		responseBuilder.WriteString(fmt.Sprintf("- **Description:** %s\n", doc.Description))
		if doc.Observations != "" {
			responseBuilder.WriteString(fmt.Sprintf("- **Observations:** %s\n", doc.Observations))
		}
		responseBuilder.WriteString(fmt.Sprintf("- **Syntax:**\n```vbscript\n%s\n```\n\n", strings.TrimSpace(doc.Syntax)))
	}

	return mcp.NewToolResultText(responseBuilder.String()), nil
}

// getASPCodingStyleHandler returns the full ASP/VBScript coding-style guide for formatting and compatibility guidance.
func getASPCodingStyleHandler(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_ = ctx
	_ = request

	content, err := os.ReadFile(styleFilePath)
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("Failed to read coding style guide: %v", err)), nil
	}

	if len(content) == 0 {
		return mcp.NewToolResultError("The coding style guide is empty."), nil
	}

	return mcp.NewToolResultText(string(content)), nil
}

// isDocBlockValid returns true when all required markdown fields are present.
func isDocBlockValid(hasKeywords, hasDescription, hasObservations, hasSyntax, hasCodeBlock bool) bool {
	return hasKeywords && hasDescription && hasObservations && hasSyntax && hasCodeBlock
}

// printDocBlockError prints a detailed validation error for a malformed function block.
func printDocBlockError(functionName string, line int, hasKeywords, hasDescription, hasObservations, hasSyntax, hasCodeBlock bool) {
	fmt.Printf("\nStructural error detected in function block: '%s' (near line %d)\n", functionName, line)
	fmt.Println("Missing or incorrectly formatted required fields:")
	if !hasKeywords {
		fmt.Println("- Missing required field: **Keywords:**")
	}
	if !hasDescription {
		fmt.Println("- Missing required field: **Description:**")
	}
	if !hasObservations {
		fmt.Println("- Missing required field: **Observations:**")
	}
	if !hasSyntax {
		fmt.Println("- Missing required field: **Syntax:**")
	}
	if !hasCodeBlock {
		fmt.Println("- Missing required VBScript code block (```vbscript ... ```)")
	}
}

// runDocsMarkdownLinter validates docs markdown structure and returns process exit code.
func runDocsMarkdownLinter(filePath string) int {
	file, err := os.Open(filePath)
	if err != nil {
		fmt.Printf("Fatal error: could not open file %s: %v\n", filePath, err)
		return 1
	}
	defer file.Close()

	scanner := bufio.NewScanner(file)
	lineNum := 0
	currentFunction := ""
	errorsFound := 0

	hasKeywords := false
	hasDescription := false
	hasObservations := false
	hasSyntax := false
	inCodeBlock := false
	hasCodeBlock := false

	fmt.Printf("Starting strict validation for %s...\n", filePath)

	for scanner.Scan() {
		lineNum++
		line := scanner.Text()
		trimmed := strings.TrimSpace(line)

		if strings.HasPrefix(trimmed, "## ") {
			if currentFunction != "" {
				if !isDocBlockValid(hasKeywords, hasDescription, hasObservations, hasSyntax, hasCodeBlock) {
					printDocBlockError(currentFunction, lineNum, hasKeywords, hasDescription, hasObservations, hasSyntax, hasCodeBlock)
					errorsFound++
				}
			}

			currentFunction = strings.TrimPrefix(trimmed, "## ")
			hasKeywords, hasDescription, hasObservations, hasSyntax, hasCodeBlock = false, false, false, false, false
			inCodeBlock = false
			continue
		}

		if currentFunction != "" {
			switch {
			case strings.HasPrefix(trimmed, "**Keywords:**"):
				hasKeywords = true
			case strings.HasPrefix(trimmed, "**Description:**"):
				hasDescription = true
			case strings.HasPrefix(trimmed, "**Observations:**"):
				hasObservations = true
			case strings.HasPrefix(trimmed, "**Syntax:**"):
				hasSyntax = true
			case strings.HasPrefix(trimmed, "```vbscript"):
				inCodeBlock = true
			case inCodeBlock && strings.HasPrefix(trimmed, "```"):
				inCodeBlock = false
				hasCodeBlock = true
			}
		}
	}

	if currentFunction != "" {
		if !isDocBlockValid(hasKeywords, hasDescription, hasObservations, hasSyntax, hasCodeBlock) {
			printDocBlockError(currentFunction, lineNum, hasKeywords, hasDescription, hasObservations, hasSyntax, hasCodeBlock)
			errorsFound++
		}
	}

	if err := scanner.Err(); err != nil {
		fmt.Printf("Error while reading file: %v\n", err)
		return 1
	}

	if errorsFound > 0 {
		fmt.Printf("\nVALIDATION FAILED: %d formatting error(s) found by the linter.\n", errorsFound)
		fmt.Println("The file was rejected because it does not follow the required MCP documentation structure.")
		return 1
	}

	fmt.Println("\nSUCCESS: The file passed all formatting rules and is ready for MCP.")
	return 0
}

func main() {
	lintDocs := flag.Bool("lint-docs", false, "Run docs markdown linter and exit.")
	docsFile := flag.String("docs-file", docFilePath, "Path to the documentation markdown file.")
	flag.Parse()

	docFilePath = *docsFile

	if *lintDocs {
		os.Exit(runDocsMarkdownLinter(docFilePath))
	}

	// Initialize the file check on boot
	if err := reloadDocsIfNeeded(); err != nil {
		fmt.Fprintf(os.Stderr, "[G3pix AxonASP MCP] Critical Error: %v\n", err)
		os.Exit(1)
	}

	// 1. Instantiate the MCP Server
	s := server.NewMCPServer(
		"G3pix AxonASP Docs",
		"1.0.0",
		server.WithPromptCapabilities(true),
	)

	// 2. Create the Tool Definition (English for token optimization)
	tool := mcp.NewTool(
		"search_axonasp_docs",
		mcp.WithDescription("Search for AxonASP built-in functions, custom objects, and libraries to be used in the Classic ASP implementation of AxonASP. Always use this tool before creating complex code to get the correct syntax and avoid hallucinating or unnecessary manual implementations. Use keywords like function names (e.g., G3JSON), actions (e.g., parse json, database, session, upload), or general topics (e.g., file handling, error handling). Don't use more than 3 keywords. Use get_asp_coding_style tool to get the official coding style guide for Classic ASP and VBScript. You must use english keywords."), // The description is intentionally verbose to guide the user towards effective queries and to optimize token usage by providing clear instructions on how to search for documentation.
		mcp.WithString("query", mcp.Required(), mcp.Description("Search term, module name, or action (e.g., G3JSON, parse json, database). Max of 3 keywords. You must use english keywords.")),
	)

	// 3. Register the tool
	s.AddTool(tool, searchHandler)

	styleTool := mcp.NewTool(
		"get_asp_coding_style",
		mcp.WithDescription("Get the official Classic ASP and VBScript coding-style instructions used by AxonASP. Call this tool whenever you need code formatting rules, control-structure conventions, object assignment rules, or compatibility guidance before writing or refactoring ASP code. Returns the full content of mcp/aspcodingstyle.md."),
	)

	s.AddTool(styleTool, getASPCodingStyleHandler)

	// 4. Start server based on configured mode (stdio or SSE)
	if mcpMode == "sse" {
		addr := fmt.Sprintf(":%d", mcpSSEPort)
		fmt.Fprintf(os.Stderr, "[G3pix AxonASP MCP] Starting in SSE mode on %s\n", addr)

		sseServer := server.NewSSEServer(
			s,
			server.WithBaseURL(fmt.Sprintf("http://localhost:%d", mcpSSEPort)),
			server.WithUseFullURLForMessageEndpoint(true),
		)

		if err := sseServer.Start(addr); err != nil {
			fmt.Fprintf(os.Stderr, "MCP SSE Server error: %v\n", err)
		}
		return
	}

	fmt.Fprintln(os.Stderr, "[G3pix AxonASP MCP] Starting in stdio mode")
	if err := server.ServeStdio(s); err != nil {
		fmt.Fprintf(os.Stderr, "MCP Server error: %v\n", err)
	}
}
