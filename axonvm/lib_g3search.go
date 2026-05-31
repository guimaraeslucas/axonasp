//go:build !wasm && !lib_g3search_disabled

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

package axonvm

import (
	"context"
	"os"
	"path/filepath"
	"runtime"
	"strings"
	"sync"

	"g3pix.com.br/axonasp/vbscript"
	"github.com/blugelabs/bluge"
)

const g3SearchBuiltMarkerFile = ".axonasp_g3search_built"

var (
	globalBlugeReader     *bluge.Reader
	globalBlugeReaderPath string
	globalReaderMutex     sync.RWMutex
	globalBuiltIndexes    = make(map[string]bool)
	globalBuildingIndexes = make(map[string]chan struct{})
	globalBuildRuns       = make(map[string]int)
)

// G3Search provides high-performance document indexing and searching using Bluge.
type G3Search struct {
	indexPath string
	docsPath  string
	extension string
	vm        *VM
}

// NewG3Search creates a new G3Search object with default extension.
func NewG3Search(vm *VM) *G3Search {

	return &G3Search{
		extension: ".md",
		vm:        vm,
	}
}

// closeGlobalReaderLocked closes the shared reader and clears global reader state.
func closeGlobalReaderLocked() error {
	if globalBlugeReader == nil {
		return nil
	}
	err := globalBlugeReader.Close()
	globalBlugeReader = nil
	globalBlugeReaderPath = ""
	return err
}

// openGlobalReaderLocked opens the shared reader for one index path under write lock.
func openGlobalReaderLocked(indexPath string) error {
	config := bluge.DefaultConfig(indexPath)
	reader, err := bluge.OpenReader(config)
	if err != nil {
		return err
	}
	globalBlugeReader = reader
	globalBlugeReaderPath = indexPath
	return nil
}

// canonicalIndexPath normalizes one path so equivalent directories map to one state key.
func canonicalIndexPath(path string) string {
	trimmed := strings.TrimSpace(path)
	if trimmed == "" {
		return ""
	}

	normalized := filepath.Clean(trimmed)
	absPath, err := filepath.Abs(normalized)
	if err == nil {
		normalized = absPath
	}

	return normalized
}

// canonicalIndexKey returns the map key used for per-directory build state.
func canonicalIndexKey(path string) string {
	normalized := canonicalIndexPath(path)
	if runtime.GOOS == "windows" {
		return strings.ToLower(normalized)
	}
	return normalized
}

// hasPersistentBuildMarker reports whether this index path was already built.
func hasPersistentBuildMarker(indexPath string) bool {
	markerPath := filepath.Join(indexPath, g3SearchBuiltMarkerFile)
	info, err := os.Stat(markerPath)
	if err != nil {
		return false
	}
	return !info.IsDir()
}

// writePersistentBuildMarker writes a durable marker after one successful build.
func writePersistentBuildMarker(indexPath string, docsPath string) error {
	markerPath := filepath.Join(indexPath, g3SearchBuiltMarkerFile)
	return os.WriteFile(markerPath, []byte("docs="+canonicalIndexPath(docsPath)), 0o644)
}

// readPersistentBuildMarkerDocsPath returns docsPath metadata from the marker.
// It returns ok=false for legacy/unparseable markers.
func readPersistentBuildMarkerDocsPath(indexPath string) (docsPath string, ok bool) {
	markerPath := filepath.Join(indexPath, g3SearchBuiltMarkerFile)
	raw, err := os.ReadFile(markerPath)
	if err != nil {
		return "", false
	}
	text := strings.TrimSpace(string(raw))
	if !strings.HasPrefix(text, "docs=") {
		return "", false
	}
	value := strings.TrimSpace(strings.TrimPrefix(text, "docs="))
	if value == "" {
		return "", false
	}
	return canonicalIndexPath(value), true
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

// beginGlobalIndexBuild reserves exactly one build slot for one canonical index path.
func beginGlobalIndexBuild(indexPath string) (<-chan struct{}, bool) {
	indexKey := canonicalIndexKey(indexPath)
	if indexKey == "" {
		return nil, false
	}

	globalReaderMutex.Lock()
	defer globalReaderMutex.Unlock()

	if globalBuiltIndexes[indexKey] {
		return nil, false
	}

	if hasPersistentBuildMarker(indexPath) {
		globalBuiltIndexes[indexKey] = true
		return nil, false
	}

	if waitCh, exists := globalBuildingIndexes[indexKey]; exists {
		return waitCh, false
	}

	waitCh := make(chan struct{})
	globalBuildingIndexes[indexKey] = waitCh
	globalBuildRuns[indexKey]++
	return waitCh, true
}

// finishGlobalIndexBuild marks one build slot as complete and wakes waiters.
func finishGlobalIndexBuild(indexPath string, buildSucceeded bool) {
	indexKey := canonicalIndexKey(indexPath)
	if indexKey == "" {
		return
	}

	globalReaderMutex.Lock()
	if buildSucceeded {
		globalBuiltIndexes[indexKey] = true
	}

	if waitCh, exists := globalBuildingIndexes[indexKey]; exists {
		delete(globalBuildingIndexes, indexKey)
		close(waitCh)
	}
	globalReaderMutex.Unlock()
}

// acquireGlobalReaderForSearch returns a shared reader pinned by RLock for one search call.
func (s *G3Search) acquireGlobalReaderForSearch() (*bluge.Reader, func(), error) {
	indexPath := canonicalIndexPath(s.indexPath)
	if indexPath == "" {
		return nil, nil, os.ErrInvalid
	}
	s.indexPath = indexPath

	for {
		globalReaderMutex.RLock()
		if globalBlugeReader != nil && strings.EqualFold(globalBlugeReaderPath, indexPath) {
			return globalBlugeReader, globalReaderMutex.RUnlock, nil
		}
		globalReaderMutex.RUnlock()

		globalReaderMutex.Lock()
		if globalBlugeReader != nil && !strings.EqualFold(globalBlugeReaderPath, indexPath) {
			if err := closeGlobalReaderLocked(); err != nil {
				globalReaderMutex.Unlock()
				return nil, nil, err
			}
		}
		if globalBlugeReader == nil {
			if err := openGlobalReaderLocked(indexPath); err != nil {
				globalReaderMutex.Unlock()
				return nil, nil, err
			}
		}
		globalReaderMutex.Unlock()
	}
}

// DispatchMethod executes methods and property Let/Set behavior for G3Search.
func (s *G3Search) DispatchMethod(methodName string, args []Value) Value {

	switch {
	case strings.EqualFold(methodName, "BuildIndex"):
		if len(args) > 0 {
			s.indexPath = canonicalIndexPath(args[0].String())
		}
		if len(args) > 1 {
			s.docsPath = canonicalIndexPath(args[1].String())
		}
		if len(args) > 2 {
			s.extension = args[2].String()
			if s.extension != "" && !strings.HasPrefix(s.extension, ".") {
				s.extension = "." + s.extension
			}
		}
		return s.buildIndex()
	case strings.EqualFold(methodName, "Search"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		return s.search(args[0].String())
	case strings.EqualFold(methodName, "IndexPath"):
		if len(args) > 0 {
			s.indexPath = canonicalIndexPath(args[0].String())
			return Value{Type: VTEmpty}
		}
		return NewString(s.indexPath)
	case strings.EqualFold(methodName, "DocsPath"):
		if len(args) > 0 {
			s.docsPath = canonicalIndexPath(args[0].String())
			return Value{Type: VTEmpty}
		}
		return NewString(s.docsPath)
	case strings.EqualFold(methodName, "Extension"):
		if len(args) > 0 {
			s.extension = args[0].String()
			if s.extension != "" && !strings.HasPrefix(s.extension, ".") {
				s.extension = "." + s.extension
			}
			return Value{Type: VTEmpty}
		}
		return NewString(s.extension)
	default:
		return Value{Type: VTEmpty}
	}
}

// DispatchPropertyGet resolves property reads for G3Search.
func (s *G3Search) DispatchPropertyGet(propertyName string) Value {

	switch {
	case strings.EqualFold(propertyName, "IndexPath"):
		return NewString(s.indexPath)
	case strings.EqualFold(propertyName, "DocsPath"):
		return NewString(s.docsPath)
	case strings.EqualFold(propertyName, "Extension"):
		return NewString(s.extension)
	default:
		return Value{Type: VTEmpty}
	}
}

// DispatchPropertySet handles Let assignments on G3Search properties.
func (s *G3Search) DispatchPropertySet(propertyName string, val Value) {

	switch {
	case strings.EqualFold(propertyName, "IndexPath"):
		s.indexPath = canonicalIndexPath(val.String())
	case strings.EqualFold(propertyName, "DocsPath"):
		s.docsPath = canonicalIndexPath(val.String())
	case strings.EqualFold(propertyName, "Extension"):
		s.extension = val.String()
		if s.extension != "" && !strings.HasPrefix(s.extension, ".") {
			s.extension = "." + s.extension
		}
	}
}

// buildIndex iterates through DocsPath and indexes files matching Extension.
func (s *G3Search) buildIndex() Value {
	indexPath := canonicalIndexPath(s.indexPath)
	docsPath := canonicalIndexPath(s.docsPath)
	s.indexPath = indexPath
	s.docsPath = docsPath

	if docsPath == "" {
		s.vm.raise(vbscript.InternalError, ErrG3SearchDocsPathMissing.String())
		return Value{Type: VTEmpty}
	}
	if indexPath == "" {
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexPathMissing.String())
		return Value{Type: VTEmpty}
	}

	waitCh, shouldBuild := beginGlobalIndexBuild(indexPath)
	if !shouldBuild {
		if waitCh != nil {
			<-waitCh
		}
		return Value{Type: VTEmpty}
	}

	buildSucceeded := false
	defer func() {
		finishGlobalIndexBuild(indexPath, buildSucceeded)
	}()

	globalReaderMutex.Lock()

	if err := closeGlobalReaderLocked(); err != nil {
		globalReaderMutex.Unlock()
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexOpenFailed.String()+": "+err.Error())
		return Value{Type: VTEmpty}
	}

	config := bluge.DefaultConfig(indexPath)
	writer, err := bluge.OpenWriter(config)
	if err != nil {
		globalReaderMutex.Unlock()
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexOpenFailed.String()+": "+err.Error())
		return Value{Type: VTEmpty}
	}

	batch := bluge.NewBatch()

	err = filepath.Walk(docsPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		if info.IsDir() {
			if pathWithin(indexPath, path) {
				return filepath.SkipDir
			}
			return nil
		}

		if s.extension == "" || strings.EqualFold(filepath.Ext(path), s.extension) {
			content, err := os.ReadFile(path)
			if err != nil {
				return err
			}

			doc := bluge.NewDocument(path)
			doc.AddField(bluge.NewTextField("content", string(content)))
			doc.AddField(bluge.NewStoredOnlyField("filename", []byte(path)))

			batch.Update(doc.ID(), doc)
		}
		return nil
	})

	if err != nil {
		_ = writer.Close()
		globalReaderMutex.Unlock()
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexWriteFailed.String()+": "+err.Error())
		return Value{Type: VTEmpty}
	}

	if err := writer.Batch(batch); err != nil {
		_ = writer.Close()
		globalReaderMutex.Unlock()
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexWriteFailed.String()+": "+err.Error())
		return Value{Type: VTEmpty}
	}

	if err := writer.Close(); err != nil {
		globalReaderMutex.Unlock()
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexWriteFailed.String()+": "+err.Error())
		return Value{Type: VTEmpty}
	}

	if err := writePersistentBuildMarker(indexPath, s.docsPath); err != nil {
		globalReaderMutex.Unlock()
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexWriteFailed.String()+": "+err.Error())
		return Value{Type: VTEmpty}
	}

	// Mark the index as built once write and commit complete.
	buildSucceeded = true

	if err := openGlobalReaderLocked(indexPath); err != nil {
		globalReaderMutex.Unlock()
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexOpenFailed.String()+": "+err.Error())
		return Value{Type: VTEmpty}
	}

	globalReaderMutex.Unlock()

	return Value{Type: VTEmpty}
}

// search performs a search on the index and returns [filename, score] tuples.
func (s *G3Search) search(term string) Value {

	if s.indexPath == "" {
		s.vm.raise(vbscript.InternalError, ErrG3SearchIndexPathMissing.String())
		return Value{Type: VTEmpty}
	}

	searchOnce := func() (Value, bool) {
		reader, release, err := s.acquireGlobalReaderForSearch()
		if err != nil {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, nil)}, true
		}
		defer release()

		query := bluge.NewMatchQuery(term).SetField("content").SetFuzziness(1)
		searchRequest := bluge.NewAllMatches(query)

		iter, searchErr := reader.Search(context.Background(), searchRequest)
		if searchErr != nil {
			s.vm.raise(vbscript.InternalError, ErrG3SearchSearchFailed.String()+": "+searchErr.Error())
			return Value{Type: VTEmpty}, false
		}

		results := make([]Value, 0, 16)
		match, iterErr := iter.Next()
		for iterErr == nil && match != nil {
			filename := ""
			match.VisitStoredFields(func(field string, value []byte) bool {
				if field == "filename" {
					filename = string(value)
				}
				return true
			})

			if filename != "" {
				tuple := [2]Value{NewString(filename), NewDouble(match.Score)}
				results = append(results, Value{Type: VTArray, Arr: NewVBArrayFromValues(0, tuple[:])})
			}

			match, iterErr = iter.Next()
		}

		if iterErr != nil {
			s.vm.raise(vbscript.InternalError, ErrG3SearchSearchFailed.String()+": "+iterErr.Error())
			return Value{Type: VTEmpty}, false
		}

		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, results)}, len(results) == 0
	}

	indexHasAnyDocument := func() (bool, error) {
		reader, release, err := s.acquireGlobalReaderForSearch()
		if err != nil {
			return false, err
		}
		defer release()

		iter, searchErr := reader.Search(context.Background(), bluge.NewAllMatches(bluge.NewMatchAllQuery()))
		if searchErr != nil {
			return false, searchErr
		}

		match, iterErr := iter.Next()
		if iterErr != nil {
			return false, iterErr
		}
		return match != nil, nil
	}

	result, empty := searchOnce()
	if !empty || s.docsPath == "" {
		return result
	}

	markerDocsPath, markerOK := readPersistentBuildMarkerDocsPath(s.indexPath)
	if markerOK && strings.EqualFold(markerDocsPath, canonicalIndexPath(s.docsPath)) {
		hasDocs, hasDocsErr := indexHasAnyDocument()
		if hasDocsErr == nil && hasDocs {
			return result
		}
	}

	// Self-heal stale/empty indexes: force one rebuild and retry once.
	indexPath := canonicalIndexPath(s.indexPath)
	if indexPath == "" {
		return result
	}

	globalReaderMutex.Lock()
	_ = closeGlobalReaderLocked()
	delete(globalBuiltIndexes, canonicalIndexKey(indexPath))
	globalReaderMutex.Unlock()

	_ = os.Remove(filepath.Join(indexPath, g3SearchBuiltMarkerFile))
	_ = s.buildIndex()

	retryResult, _ := searchOnce()
	return retryResult
}
