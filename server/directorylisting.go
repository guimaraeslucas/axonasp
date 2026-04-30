/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Lucas GuimarÃ£es - G3pix Ltda
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
package main

import (
	"encoding/base64"
	"fmt"
	"html/template"
	"io"
	"mime"
	"net/http"
	"net/url"
	"os"
	"path"
	"path/filepath"
	"sort"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
)

// DirectoryListingRenderer renders folder listings using a configurable HTML template.
type DirectoryListingRenderer struct {
	rootDir      string
	templatePath string
	template     *template.Template
	blockedFiles map[string]struct{}
	blockedDirs  map[string]struct{}
	blockedExts  map[string]struct{}
}

type directoryListingEntry struct {
	Name        string
	Href        string
	TypeLabel   string
	MimeType    string
	SizeDisplay string
	Modified    string
	IsDir       bool
}

type directoryListingPageData struct {
	Title       string
	RequestPath string
	ParentPath  string
	GeneratedAt string
	InlineCSS   template.CSS
	LogoDataURI template.URL
	Entries     []directoryListingEntry
}

type directoryListingAssets struct {
	inlineCSS   string
	logoDataURI string
}

var directoryListingAssetsOnce sync.Once
var cachedDirectoryListingAssets directoryListingAssets

// NewDirectoryListingRenderer compiles the HTML template and initializes listing filters.
func NewDirectoryListingRenderer(rootDir string, templatePath string) (*DirectoryListingRenderer, error) {
	tmpl, err := template.New("directory-listing").ParseFiles(templatePath)
	if err != nil {
		return nil, err
	}

	renderer := &DirectoryListingRenderer{
		rootDir:      rootDir,
		templatePath: templatePath,
		template:     tmpl,
		blockedFiles: makeLookupSet(BlockedFiles),
		blockedDirs:  makeLookupSet(BlockedDirs),
		blockedExts:  makeLookupSet(BlockedExtensions),
	}
	return renderer, nil
}

// Render writes one directory listing response using the configured template.
func (d *DirectoryListingRenderer) Render(w http.ResponseWriter, r *http.Request, absDirPath string, requestPath string) error {
	entries, err := os.ReadDir(absDirPath)
	if err != nil {
		return err
	}

	visible := make([]directoryListingEntry, 0, len(entries))
	for _, entry := range entries {
		name := entry.Name()
		if name == "" {
			continue
		}
		nameLower := strings.ToLower(name)

		if entry.IsDir() {
			if d.isBlockedDirName(nameLower) {
				continue
			}
			childPath := filepath.Join(absDirPath, name)
			childAbs, err := filepath.Abs(childPath)
			if err == nil && isBlockedDirectory(childAbs) {
				continue
			}
			if strings.EqualFold(name, ".") || strings.EqualFold(name, "..") {
				continue
			}
			entryInfo, infoErr := entry.Info()
			if infoErr != nil {
				continue
			}
			visible = append(visible, directoryListingEntry{
				Name:        name,
				Href:        joinRequestPath(requestPath, name, true),
				TypeLabel:   "Directory",
				MimeType:    "-",
				SizeDisplay: "-",
				Modified:    entryInfo.ModTime().Format("2006-01-02 15:04:05"),
				IsDir:       true,
			})
			continue
		}

		if d.isBlockedFileName(nameLower) || d.isBlockedExtension(filepath.Ext(nameLower)) {
			continue
		}

		entryInfo, infoErr := entry.Info()
		if infoErr != nil {
			continue
		}
		visible = append(visible, directoryListingEntry{
			Name:        name,
			Href:        joinRequestPath(requestPath, name, false),
			TypeLabel:   "File",
			MimeType:    detectFileMIME(filepath.Join(absDirPath, name), nameLower),
			SizeDisplay: formatFileSize(entryInfo.Size()),
			Modified:    entryInfo.ModTime().Format("2006-01-02 15:04:05"),
			IsDir:       false,
		})
	}

	sort.Slice(visible, func(i, j int) bool {
		if visible[i].IsDir != visible[j].IsDir {
			return visible[i].IsDir
		}
		return strings.ToLower(visible[i].Name) < strings.ToLower(visible[j].Name)
	})

	assets := loadDirectoryListingAssets()
	pageData := directoryListingPageData{
		Title:       "❖ AxonASP Directory Listing",
		RequestPath: requestPath,
		ParentPath:  resolveParentPath(requestPath),
		GeneratedAt: time.Now().In(serverLocation).Format("2006-01-02 15:04:05 MST"),
		InlineCSS:   template.CSS(assets.inlineCSS),
		LogoDataURI: template.URL(assets.logoDataURI),
		Entries:     visible,
	}

	w.Header().Set("Content-Type", "text/html; charset=utf-8")
	w.WriteHeader(http.StatusOK)
	return d.template.ExecuteTemplate(w, filepath.Base(d.templatePath), pageData)
}

// loadDirectoryListingAssets reads the configured logo and CSS assets once and caches the output.
func loadDirectoryListingAssets() directoryListingAssets {
	directoryListingAssetsOnce.Do(func() {
		assets := directoryListingAssets{}
		v := axonconfig.NewViper()

		assets.inlineCSS = readTextFileOrEmpty(v.GetString("axfunctions.ax_default_css_path"))
		assets.logoDataURI = readLogoAsDataURIOrEmpty(v.GetString("axfunctions.ax_default_logo_path"))
		cachedDirectoryListingAssets = assets
	})
	return cachedDirectoryListingAssets
}

// readTextFileOrEmpty reads a UTF-8 text file and returns its full content or empty string on failure.
func readTextFileOrEmpty(filePath string) string {
	cleanPath := strings.TrimSpace(filePath)
	if cleanPath == "" {
		return ""
	}
	data, err := os.ReadFile(cleanPath)
	if err != nil {
		return ""
	}
	return string(data)
}

// readLogoAsDataURIOrEmpty reads a binary image file and encodes it as a data URI string.
func readLogoAsDataURIOrEmpty(filePath string) string {
	cleanPath := strings.TrimSpace(filePath)
	if cleanPath == "" {
		return ""
	}
	data, err := os.ReadFile(cleanPath)
	if err != nil || len(data) == 0 {
		return ""
	}

	ext := strings.ToLower(filepath.Ext(cleanPath))
	mimeType := mime.TypeByExtension(ext)
	if mimeType == "" {
		mimeType = http.DetectContentType(data)
	}
	if mimeType == "" {
		mimeType = "application/octet-stream"
	}

	return "data:" + mimeType + ";base64," + base64.StdEncoding.EncodeToString(data)
}

// detectFileMIME determines MIME type using extension first and falls back to content sniffing.
func detectFileMIME(absPath string, nameLower string) string {
	mimeType := mime.TypeByExtension(filepath.Ext(nameLower))
	if mimeType != "" {
		return mimeType
	}

	f, err := os.Open(absPath)
	if err != nil {
		return "application/octet-stream"
	}
	defer f.Close()

	buf := make([]byte, 512)
	n, readErr := io.ReadFull(f, buf)
	if readErr != nil && readErr != io.ErrUnexpectedEOF && readErr != io.EOF {
		return "application/octet-stream"
	}
	if n == 0 {
		return "application/octet-stream"
	}

	return http.DetectContentType(buf[:n])
}

func (d *DirectoryListingRenderer) isBlockedFileName(nameLower string) bool {
	_, blocked := d.blockedFiles[nameLower]
	return blocked
}

func (d *DirectoryListingRenderer) isBlockedDirName(nameLower string) bool {
	_, blocked := d.blockedDirs[nameLower]
	if blocked {
		return true
	}
	baseName := strings.ToLower(filepath.Base(filepath.Clean(nameLower)))
	_, blocked = d.blockedDirs[baseName]
	return blocked
}

func (d *DirectoryListingRenderer) isBlockedExtension(ext string) bool {
	ext = strings.ToLower(strings.TrimSpace(ext))
	if ext == "" {
		return false
	}
	_, blocked := d.blockedExts[ext]
	return blocked
}

func makeLookupSet(values []string) map[string]struct{} {
	lookup := make(map[string]struct{}, len(values))
	for _, value := range values {
		clean := strings.ToLower(strings.TrimSpace(value))
		if clean == "" {
			continue
		}
		lookup[clean] = struct{}{}
		lookup[strings.ToLower(filepath.Base(filepath.Clean(clean)))] = struct{}{}
	}
	return lookup
}

func joinRequestPath(requestPath string, name string, isDir bool) string {
	base := requestPath
	if base == "" {
		base = "/"
	}
	if !strings.HasSuffix(base, "/") {
		base += "/"
	}
	escapedName := url.PathEscape(name)
	joined := path.Clean(base + escapedName)
	if !strings.HasPrefix(joined, "/") {
		joined = "/" + joined
	}
	if isDir && !strings.HasSuffix(joined, "/") {
		joined += "/"
	}
	return joined
}

func resolveParentPath(requestPath string) string {
	clean := path.Clean("/" + strings.TrimSpace(requestPath))
	if clean == "/" {
		return ""
	}
	parent := path.Dir(clean)
	if parent == "." || parent == "/" {
		return "/"
	}
	if !strings.HasSuffix(parent, "/") {
		parent += "/"
	}
	return parent
}

func formatFileSize(size int64) string {
	if size < 1024 {
		return fmt.Sprintf("%d B", size)
	}
	if size < 1024*1024 {
		return fmt.Sprintf("%.1f KB", float64(size)/1024.0)
	}
	if size < 1024*1024*1024 {
		return fmt.Sprintf("%.1f MB", float64(size)/(1024.0*1024.0))
	}
	return fmt.Sprintf("%.1f GB", float64(size)/(1024.0*1024.0*1024.0))
}
