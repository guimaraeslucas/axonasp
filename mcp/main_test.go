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
 */

package main

import (
	"context"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/blugelabs/bluge"
	"github.com/mark3labs/mcp-go/mcp"
)

func TestSearchAndReadWorkflow(t *testing.T) {
	// 1. Setup temporary test environment
	tmpDocs, err := os.MkdirTemp("", "mcp-docs-test")
	if err != nil {
		t.Fatalf("Failed to create temp docs dir: %v", err)
	}
	defer os.RemoveAll(tmpDocs)

	tmpIndex, err := os.MkdirTemp("", "mcp-index-test")
	if err != nil {
		t.Fatalf("Failed to create temp index dir: %v", err)
	}
	defer os.RemoveAll(tmpIndex)

	// Create a dummy global.asa documentation file
	docsSubDir := filepath.Join(tmpDocs, "asp")
	err = os.MkdirAll(docsSubDir, 0755)
	if err != nil {
		t.Fatalf("Failed to create docs subdir: %v", err)
	}

	docPath := "asp/global-asa.md"
	docContent := "# global.asa\nThis is the documentation for global.asa file in AxonASP."
	err = os.WriteFile(filepath.Join(tmpDocs, docPath), []byte(docContent), 0644)
	if err != nil {
		t.Fatalf("Failed to write test doc: %v", err)
	}

	// Override global paths
	originalDocsPath := docFilePath
	originalIndexPath := indexPath
	docFilePath = tmpDocs
	indexPath = tmpIndex
	defer func() {
		docFilePath = originalDocsPath
		indexPath = originalIndexPath
	}()

	// 2. Test Index Rebuild
	idx := &SearchIndex{indexPath: indexPath, docsPath: docFilePath}
	err = idx.Rebuild()
	if err != nil {
		t.Fatalf("Failed to rebuild index: %v", err)
	}

	// Verify index files exist
	files, _ := os.ReadDir(indexPath)
	if len(files) == 0 {
		t.Errorf("Index directory is empty, expected bluge index files")
	}

	// 3. Open Global Reader for testing handlers
	config := bluge.DefaultConfig(indexPath)
	reader, err := bluge.OpenReader(config)
	if err != nil {
		t.Fatalf("Failed to open reader: %v", err)
	}
	globalReader = reader
	defer globalReader.Close()

	// 4. Test Search Handler (Step A) - Fuzzy matching for "global.asa"
	// We'll test with a slight typo "globa.asa" to verify fuzzy matching
	req := mcp.CallToolRequest{
		Params: mcp.CallToolParams{
			Name: "search_axonasp_docs",
			Arguments: map[string]any{
				"query": "globa.asa",
			},
		},
	}

	result, err := searchHandler(context.Background(), req)
	if err != nil {
		t.Fatalf("searchHandler failed: %v", err)
	}

	if result.IsError {
		t.Fatalf("searchHandler returned error: %s", result.Content[0].(mcp.TextContent).Text)
	}

	searchText := result.Content[0].(mcp.TextContent).Text
	if !strings.Contains(searchText, "asp/global-asa.md") {
		t.Errorf("Search result does not contain expected path 'asp/global-asa.md'. Got: %s", searchText)
	}
	if !strings.Contains(searchText, "Most probable match") {
		t.Errorf("Search result should identify the highest-ranked entry. Got: %s", searchText)
	}
	if !strings.Contains(searchText, "read_axonasp_doc") {
		t.Errorf("Search result should instruct callers to use read_axonasp_doc. Got: %s", searchText)
	}

	// 5. Test Read Handler (Step B)
	readReq := mcp.CallToolRequest{
		Params: mcp.CallToolParams{
			Name: "read_axonasp_doc",
			Arguments: map[string]any{
				"path": "asp/global-asa.md",
			},
		},
	}

	readResult, err := readDocHandler(context.Background(), readReq)
	if err != nil {
		t.Fatalf("readDocHandler failed: %v", err)
	}

	if readResult.IsError {
		t.Fatalf("readDocHandler returned error: %s", readResult.Content[0].(mcp.TextContent).Text)
	}

	readText := readResult.Content[0].(mcp.TextContent).Text
	if !strings.Contains(readText, "# global.asa") {
		t.Errorf("Read result does not contain expected content. Got: %s", readText)
	}
}

func TestSymlinkLoopProtection(t *testing.T) {
	tmpDocs, err := os.MkdirTemp("", "mcp-symlink-test")
	if err != nil {
		t.Fatalf("Failed to create temp docs dir: %v", err)
	}
	defer os.RemoveAll(tmpDocs)

	tmpIndex, err := os.MkdirTemp("", "mcp-symlink-index")
	if err != nil {
		t.Fatalf("Failed to create temp index dir: %v", err)
	}
	defer os.RemoveAll(tmpIndex)

	// Create a nested directory structure
	dir1 := filepath.Join(tmpDocs, "dir1")
	dir2 := filepath.Join(tmpDocs, "dir1", "dir2")
	os.MkdirAll(dir2, 0755)

	// Create an inverted symlink: dir2/link_to_dir1 -> dir1
	// This creates an infinite recursion if not handled
	err = os.Symlink(dir1, filepath.Join(dir2, "link_to_dir1"))
	if err != nil {
		// On some Windows environments, symlinks might require admin privileges.
		// Skip the test if symlink creation fails.
		t.Skip("Skipping symlink test: failed to create symlink (likely permission issue)")
	}

	idx := &SearchIndex{indexPath: tmpIndex, docsPath: tmpDocs}
	err = idx.Rebuild()
	if err != nil {
		t.Errorf("Rebuild failed or timed out due to symlink loop: %v", err)
	}
}

func TestRebuildSkipsIndexSubtreeInsideDocs(t *testing.T) {
	tmpDocs, err := os.MkdirTemp("", "mcp-docs-index-subtree")
	if err != nil {
		t.Fatalf("Failed to create temp docs dir: %v", err)
	}
	defer os.RemoveAll(tmpDocs)

	indexDir := filepath.Join(tmpDocs, "nested", "manual-index")
	if err := os.MkdirAll(indexDir, 0755); err != nil {
		t.Fatalf("Failed to create nested index dir: %v", err)
	}

	if err := os.WriteFile(filepath.Join(tmpDocs, "root.md"), []byte("root token"), 0644); err != nil {
		t.Fatalf("Failed to write root doc: %v", err)
	}
	if err := os.WriteFile(filepath.Join(indexDir, "generated.md"), []byte("index token"), 0644); err != nil {
		t.Fatalf("Failed to write generated doc: %v", err)
	}

	idx := &SearchIndex{indexPath: indexDir, docsPath: tmpDocs}
	if err := idx.Rebuild(); err != nil {
		t.Fatalf("Failed to rebuild index: %v", err)
	}

	config := bluge.DefaultConfig(idx.indexPath)
	reader, err := bluge.OpenReader(config)
	if err != nil {
		t.Fatalf("Failed to open reader: %v", err)
	}
	defer reader.Close()

	rootQuery := bluge.NewMatchQuery("root").SetField("content")
	rootResults, err := reader.Search(context.Background(), bluge.NewTopNSearch(5, rootQuery))
	if err != nil {
		t.Fatalf("Failed to search root token: %v", err)
	}
	match, err := rootResults.Next()
	if err != nil {
		t.Fatalf("Failed to iterate root results: %v", err)
	}
	if match == nil {
		t.Fatalf("Expected root doc to be indexed")
	}

	indexQuery := bluge.NewMatchQuery("index").SetField("content")
	indexResults, err := reader.Search(context.Background(), bluge.NewTopNSearch(5, indexQuery))
	if err != nil {
		t.Fatalf("Failed to search nested index token: %v", err)
	}
	match, err = indexResults.Next()
	if err != nil {
		t.Fatalf("Failed to iterate nested index results: %v", err)
	}
	if match != nil {
		t.Fatalf("Expected index subtree content to be skipped")
	}
}
