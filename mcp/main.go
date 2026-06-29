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
//go:generate goversioninfo -icon=icon_mcp.ico -64=true
package main

import (
	"context"
	"flag"
	"fmt"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"

	"g3pix.com.br/axonasp/axonconfig"
	"github.com/blugelabs/bluge"
	"github.com/fsnotify/fsnotify"
	"github.com/joho/godotenv"
	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"github.com/spf13/viper"
)

// Global state and configuration
var (
	Version       = "0.0.0.0"
	mu            sync.RWMutex
	docFilePath   = "./www/manual/md/"
	styleFilePath = "mcp/aspcodingstyle.md"
	indexPath     = "./mcp/search-index/manual/"
	mcpMode       = "stdio"
	mcpSSEPort    = 8000
	globalReader  *bluge.Reader
	readerOnce    sync.Once
)

// SearchIndex manages the Bluge index lifecycle.
type SearchIndex struct {
	indexPath string
	docsPath  string
}

// canonicalIndexPath normalizes a path so containment checks are stable across relative paths and separators.
func canonicalIndexPath(path string) string {
	trimmed := strings.TrimSpace(path)
	if trimmed == "" {
		return ""
	}
	absPath, err := filepath.Abs(trimmed)
	if err != nil {
		absPath = trimmed
	}
	cleaned := filepath.Clean(absPath)
	if resolved, err := filepath.EvalSymlinks(cleaned); err == nil {
		cleaned = resolved
	}
	return cleaned
}

// pathWithin reports whether candidate is inside or equal to root.
func pathWithin(root, candidate string) bool {
	rootPath := canonicalIndexPath(root)
	candidatePath := canonicalIndexPath(candidate)
	if rootPath == "" || candidatePath == "" {
		return false
	}

	if runtime.GOOS == "windows" {
		rootPath = strings.ToLower(rootPath)
		candidatePath = strings.ToLower(candidatePath)
	}

	if candidatePath == rootPath {
		return true
	}

	sep := string(os.PathSeparator)
	if !strings.HasSuffix(rootPath, sep) {
		rootPath += sep
	}

	return strings.HasPrefix(candidatePath, rootPath)
}

// canonicalDocLookupPath resolves one documentation path relative to the configured docs root.
func canonicalDocLookupPath(root string, relPath string) string {
	trimmed := strings.TrimSpace(relPath)
	if trimmed == "" {
		return ""
	}
	return canonicalIndexPath(filepath.Join(root, filepath.FromSlash(trimmed)))
}

// init loads environment variables and applies TOML-based configuration through Viper.
func init() {
	_ = godotenv.Load()
	loadMCPConfig()
}

// loadMCPConfig loads and applies MCP settings from config/axonasp.toml using Viper.
func loadMCPConfig() {
	for i := 1; i < len(os.Args); i++ {
		arg := os.Args[i]
		if (arg == "-c" || arg == "--config.config_file") && i+1 < len(os.Args) {
			axonconfig.SetCustomConfigPath(os.Args[i+1])
			break
		}
	}

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

// Rebuild clears and recreates the Bluge index by scanning the documentation directory.
func (s *SearchIndex) Rebuild() error {
	s.indexPath = canonicalIndexPath(s.indexPath)
	s.docsPath = canonicalIndexPath(s.docsPath)

	fmt.Fprintf(os.Stderr, "[G3pix AxonASP MCP] Rebuilding index at %s from %s...\n", s.indexPath, s.docsPath)

	// Remove old index
	_ = os.RemoveAll(s.indexPath)
	if err := os.MkdirAll(s.indexPath, 0755); err != nil {
		return fmt.Errorf("failed to create index directory: %w", err)
	}

	config := bluge.DefaultConfig(s.indexPath)
	writer, err := bluge.OpenWriter(config)
	if err != nil {
		return fmt.Errorf("failed to open index writer: %w", err)
	}
	defer writer.Close()

	visited := make(map[string]bool)
	return s.walkAndIndex(writer, s.docsPath, "", visited, 0)
}

// walkAndIndex recursively scans a directory and indexes .md files with symlink loop protection.
func (s *SearchIndex) walkAndIndex(writer *bluge.Writer, absPath, relPath string, visited map[string]bool, depth int) error {
	if depth > 20 {
		return nil // Safety limit for recursion
	}

	// Canonicalize path to detect symlink loops
	realPath, err := filepath.EvalSymlinks(absPath)
	if err != nil {
		return nil // Skip broken symlinks or inaccessible paths
	}
	if visited[realPath] {
		return nil // Loop detected
	}
	visited[realPath] = true

	entries, err := os.ReadDir(absPath)
	if err != nil {
		return fmt.Errorf("failed to read directory %s: %w", absPath, err)
	}

	batch := bluge.NewBatch()
	for _, entry := range entries {
		name := entry.Name()
		entryAbs := filepath.Join(absPath, name)
		entryRel := name
		if relPath != "" {
			entryRel = filepath.Join(relPath, name)
		}

		// Handle directory or symlink to directory
		info, err := entry.Info()
		if err != nil {
			continue
		}

		isDir := info.IsDir()
		if info.Mode()&os.ModeSymlink != 0 {
			resolvedInfo, err := os.Stat(entryAbs)
			if err == nil && resolvedInfo.IsDir() {
				isDir = true
			}
		}

		if isDir {
			// Skip the index directory subtree when it lives inside the docs tree.
			if pathWithin(s.indexPath, entryAbs) {
				continue
			}
			if err := s.walkAndIndex(writer, entryAbs, entryRel, visited, depth+1); err != nil {
				return err
			}
		} else if strings.HasSuffix(strings.ToLower(name), ".md") {
			content, err := os.ReadFile(entryAbs)
			if err != nil {
				continue
			}

			docID := filepath.ToSlash(entryRel)
			doc := bluge.NewDocument(docID)
			// Index content for search, store it for snippets
			doc.AddField(bluge.NewTextField("content", string(content)).StoreValue())
			// Store the relative path for retrieval
			doc.AddField(bluge.NewStoredOnlyField("path", []byte(docID)))

			batch.Update(doc.ID(), doc)
		}
	}

	if err := writer.Batch(batch); err != nil {
		return fmt.Errorf("failed to execute batch for %s: %w", absPath, err)
	}

	return nil
}

// createSnippet extracts a brief snippet from the content for search results.
func createSnippet(content string, length int) string {
	content = strings.ReplaceAll(content, "\n", " ")
	content = strings.ReplaceAll(content, "\r", "")
	content = strings.TrimSpace(content)
	if len(content) <= length {
		return content
	}
	return content[:length] + "..."
}

// searchHandler executes the fuzzy search logic using Bluge and returns a list of paths.
func searchHandler(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, ok := request.Params.Arguments.(map[string]any)
	if !ok {
		return mcp.NewToolResultError("Invalid arguments format. Expected a map."), nil
	}

	queryTerm, ok := args["query"].(string)
	if !ok || strings.TrimSpace(queryTerm) == "" {
		return mcp.NewToolResultError("Argument 'query' is required and must be a non-empty string."), nil
	}

	mu.RLock()
	reader := globalReader
	mu.RUnlock()

	if reader == nil {
		return mcp.NewToolResultError("Search index is not initialized."), nil
	}

	// Create a MatchQuery with fuzziness
	query := bluge.NewMatchQuery(queryTerm).SetField("content").SetFuzziness(1)
	searchRequest := bluge.NewTopNSearch(5, query)

	iter, err := reader.Search(ctx, searchRequest)
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("Search error: %v", err)), nil
	}

	var results []struct {
		Path    string
		Score   float64
		Snippet string
	}

	match, err := iter.Next()
	for err == nil && match != nil {
		var path string
		var content string
		match.VisitStoredFields(func(field string, value []byte) bool {
			if field == "path" {
				path = string(value)
			} else if field == "content" {
				content = string(value)
			}
			return true
		})

		results = append(results, struct {
			Path    string
			Score   float64
			Snippet string
		}{
			Path:    path,
			Score:   match.Score,
			Snippet: createSnippet(content, 150),
		})
		match, err = iter.Next()
	}

	if len(results) == 0 {
		return mcp.NewToolResultText(fmt.Sprintf("No documentation found for: '%s'. Try different keywords. Use short queries like 'G3JSON' or 'database connection'.", queryTerm)), nil
	}

	// Build the response in Markdown format
	var responseBuilder strings.Builder
	responseBuilder.WriteString(fmt.Sprintf("Ranked manual matches for '%s'. Step 1: inspect the snippet of the most probable match below. Step 2: choose a returned path and call 'read_axonasp_doc' to receive the full Markdown content. Do not browse the workspace or manual folders directly when this tool returns a likely match. ALWAYS prioritize existing server-side implementations over recreating Classic ASP code from scratch; for example, strictly use the native G3JSON object for JSON manipulation rather than raw parsing or custom ASP classes. Furthermore, you must adhere to standard Classic ASP coding patterns by avoiding single-line syntax where distinct lines are required, and always explicitly closing conditional blocks with 'End If':\n\n", queryTerm))

	for i, res := range results {
		heading := fmt.Sprintf("Match %d", i+1)
		if i == 0 {
			heading = "Most probable match"
		}
		responseBuilder.WriteString(fmt.Sprintf("### %s: %s\n", heading, res.Path))
		responseBuilder.WriteString(fmt.Sprintf("- **Score:** %.4f\n", res.Score))
		responseBuilder.WriteString(fmt.Sprintf("- **Snippet:** %s\n\n", res.Snippet))
	}

	return mcp.NewToolResultText(responseBuilder.String()), nil
}

// readDocHandler retrieves the full content of a specific documentation file.
func readDocHandler(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, ok := request.Params.Arguments.(map[string]any)
	if !ok {
		return mcp.NewToolResultError("Invalid arguments format. Expected a map."), nil
	}

	relPath, ok := args["path"].(string)
	if !ok || strings.TrimSpace(relPath) == "" {
		return mcp.NewToolResultError("Argument 'path' is required and must be a non-empty string."), nil
	}

	// Construct full path and validate security
	absDocs := canonicalIndexPath(docFilePath)
	absFile := canonicalDocLookupPath(docFilePath, relPath)

	if !pathWithin(absDocs, absFile) || absDocs == "" || absFile == "" {
		return mcp.NewToolResultError("Security error: requested path is outside the documentation directory."), nil
	}

	content, err := os.ReadFile(absFile)
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("Failed to read documentation file: %v", err)), nil
	}

	return mcp.NewToolResultText(string(content)), nil
}

// getASPCodingStyleHandler returns the full ASP/VBScript coding-style guide.
func getASPCodingStyleHandler(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	content, err := os.ReadFile(styleFilePath)
	if err != nil {
		return mcp.NewToolResultError(fmt.Sprintf("Failed to read coding style guide: %v", err)), nil
	}

	if len(content) == 0 {
		return mcp.NewToolResultError("The coding style guide is empty."), nil
	}

	return mcp.NewToolResultText(string(content)), nil
}

func main() {
	rebuildIndex := flag.Bool("rebuild-index", false, "Force rebuild of the search index.")
	docsPathFlag := flag.String("docs-path", docFilePath, "Path to the documentation directory.")
	flag.Parse()

	docFilePath = *docsPathFlag

	// Rebuild index if missing or forced
	if _, err := os.Stat(indexPath); os.IsNotExist(err) || *rebuildIndex {
		idx := &SearchIndex{indexPath: indexPath, docsPath: docFilePath}
		if err := idx.Rebuild(); err != nil {
			fmt.Fprintf(os.Stderr, "[G3pix AxonASP MCP] Critical Error: %v\n", err)
			os.Exit(1)
		}
	}

	// Open global reader (Singleton)
	config := bluge.DefaultConfig(indexPath)
	reader, err := bluge.OpenReader(config)
	if err != nil {
		fmt.Fprintf(os.Stderr, "[G3pix AxonASP MCP] Failed to open index reader: %v\n", err)
		os.Exit(1)
	}
	globalReader = reader
	defer globalReader.Close()

	// 1. Instantiate the MCP Server
	s := server.NewMCPServer(
		"G3pix AxonASP Docs",
		"2.2.0",
		server.WithPromptCapabilities(true),
	)

	// 2. Register Search Tool
	searchTool := mcp.NewTool(
		"search_axonasp_docs",
		mcp.WithDescription("Step 1 of 2 for AxonASP manual lookup. Search manual pages, examples, runtime docs, built-in functions, custom objects, and libraries. Returns ranked file paths and snippets only; then call read_axonasp_doc to receive the full Markdown content. Do not browse workspace files directly. Use english keywords."),
		mcp.WithString("query", mcp.Required(), mcp.Description("Search term or action (e.g., G3JSON, database connection).")),
	)
	s.AddTool(searchTool, searchHandler)

	// 3. Register Read Tool
	readTool := mcp.NewTool(
		"read_axonasp_doc",
		mcp.WithDescription("Step 2 of 2 for AxonASP manual lookup. Read the full Markdown content of a documentation file using the path returned by search_axonasp_docs. Use this instead of browsing the manual folder directly."),
		mcp.WithString("path", mcp.Required(), mcp.Description("The relative path of the file (e.g., libraries/g3json/overview.md).")),
	)
	s.AddTool(readTool, readDocHandler)

	// 4. Register Style Tool
	styleTool := mcp.NewTool(
		"get_asp_coding_style",
		mcp.WithDescription("Get the official Classic ASP and VBScript coding-style guide used by AxonASP."),
	)
	s.AddTool(styleTool, getASPCodingStyleHandler)

	// 5. Start server
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
