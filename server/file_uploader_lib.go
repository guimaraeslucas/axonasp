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
package server

import (
	"fmt"
	"io"
	"mime"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// G3FileUploader implements file upload functionality for ASP
type G3FileUploader struct {
	ctx                  *ExecutionContext
	blockedExtensions    map[string]bool
	allowedExtensions    map[string]bool
	useAllowedExtOnly    bool
	maxFileSize          int64
	preserveOriginalName bool
	debugMode            bool
}

// FileUploadInfo contains information about an uploaded file
type FileUploadInfo struct {
	OriginalFileName string
	NewFileName      string
	Size             int64
	MimeType         string
	Extension        string
	TemporaryPath    string
	UploadedAt       string
	IsSuccess        bool
	ErrorMessage     string
}

func (f *G3FileUploader) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "blockedextensions":
		extList := make([]string, 0)
		for ext := range f.blockedExtensions {
			extList = append(extList, ext)
		}
		return extList
	case "allowedextensions":
		extList := make([]string, 0)
		for ext := range f.allowedExtensions {
			extList = append(extList, ext)
		}
		return extList
	case "maxfilesize":
		return f.maxFileSize
	case "preserveoriginalname":
		return f.preserveOriginalName
	case "debugmode":
		return f.debugMode
	}
	return nil
}

func (f *G3FileUploader) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "maxfilesize":
		if v, ok := value.(int64); ok {
			f.maxFileSize = v
		} else if v, ok := value.(float64); ok {
			f.maxFileSize = int64(v)
		} else if v, ok := value.(int); ok {
			f.maxFileSize = int64(v)
		}
	case "preserveoriginalname":
		if v, ok := value.(bool); ok {
			f.preserveOriginalName = v
		}
	case "debugmode":
		if v, ok := value.(bool); ok {
			f.debugMode = v
		}
	}
}

func (f *G3FileUploader) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	switch method {
	case "blockextension":
		if len(args) > 0 {
			f.blockExtension(args[0])
		}
		return nil

	case "allowextension":
		if len(args) > 0 {
			f.allowExtension(args[0])
		}
		return nil

	case "blockextensions":
		if len(args) > 0 {
			f.blockExtensions(args[0])
		}
		return nil

	case "allowextensions":
		if len(args) > 0 {
			f.allowExtensions(args[0])
		}
		return nil

	case "setuseallowedonly":
		if len(args) > 0 {
			if v, ok := args[0].(bool); ok {
				f.useAllowedExtOnly = v
			}
		}
		return nil

	case "process":
		if len(args) < 1 {
			return nil
		}
		fieldName := fmt.Sprintf("%v", args[0])
		targetDir := "./"
		if len(args) > 1 {
			targetDir = fmt.Sprintf("%v", args[1])
		}
		newFileName := ""
		if len(args) > 2 {
			newFileName = fmt.Sprintf("%v", args[2])
		}
		return f.processUpload(fieldName, targetDir, newFileName)

	case "processall":
		targetDir := "./"
		if len(args) > 0 {
			targetDir = fmt.Sprintf("%v", args[0])
		}
		return f.processAllUploads(targetDir)

	case "getfileinfo":
		if len(args) < 1 {
			return nil
		}
		fieldName := fmt.Sprintf("%v", args[0])
		return f.getFileInfo(fieldName)

	case "getallfilesinfo":
		return f.getAllFilesInfo()

	case "isvalidextension":
		if len(args) < 1 {
			return false
		}
		ext := fmt.Sprintf("%v", args[0])
		return f.isValidExtension(ext)

	default:
		return nil
	}
}

// blockExtension adds an extension to the blocked list
func (f *G3FileUploader) blockExtension(ext interface{}) {
	extStr := strings.ToLower(strings.TrimSpace(fmt.Sprintf("%v", ext)))
	if !strings.HasPrefix(extStr, ".") {
		extStr = "." + extStr
	}
	f.blockedExtensions[extStr] = true
}

// blockExtensions adds multiple extensions to the blocked list (comma-separated)
func (f *G3FileUploader) blockExtensions(exts interface{}) {
	extStr := fmt.Sprintf("%v", exts)
	parts := strings.Split(extStr, ",")
	for _, part := range parts {
		f.blockExtension(strings.TrimSpace(part))
	}
}

// allowExtension adds an extension to the allowed list
func (f *G3FileUploader) allowExtension(ext interface{}) {
	extStr := strings.ToLower(strings.TrimSpace(fmt.Sprintf("%v", ext)))
	if !strings.HasPrefix(extStr, ".") {
		extStr = "." + extStr
	}
	f.allowedExtensions[extStr] = true
}

// allowExtensions adds multiple extensions to the allowed list (comma-separated)
func (f *G3FileUploader) allowExtensions(exts interface{}) {
	extStr := fmt.Sprintf("%v", exts)
	parts := strings.Split(extStr, ",")
	for _, part := range parts {
		f.allowExtension(strings.TrimSpace(part))
	}
}

// isValidExtension checks if an extension is valid
func (f *G3FileUploader) isValidExtension(ext string) bool {
	ext = strings.ToLower(strings.TrimSpace(ext))
	if !strings.HasPrefix(ext, ".") {
		ext = "." + ext
	}

	// Check blocked extensions
	if f.blockedExtensions[ext] {
		return false
	}

	// Check allowed extensions
	if f.useAllowedExtOnly {
		return f.allowedExtensions[ext]
	}

	return true
}

// processUpload handles uploading a single file
func (f *G3FileUploader) processUpload(fieldName, targetDir, newFileName string) interface{} {
	// Parse multipart form with adaptive size limit based on maxFileSize
	var parseLimit int64 = 32 << 20 // 32MB default
	if f.maxFileSize > 0 {
		parseLimit = f.maxFileSize + (5 << 20) // Add 5MB buffer for form headers
	}

	err := f.ctx.httpRequest.ParseMultipartForm(parseLimit)
	if err != nil {
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] ParseMultipartForm error: %v\n", err)
		}
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("Failed to parse form data: %v", err),
		}
	}

	// Verify form was parsed
	if f.ctx.httpRequest.MultipartForm == nil {
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "No multipart form data received",
		}
	}

	// Debug: List available fields
	if f.debugMode && f.ctx.httpRequest.MultipartForm.File != nil {
		fmt.Printf("[G3FileUploader DEBUG] Available file fields: ")
		for key := range f.ctx.httpRequest.MultipartForm.File {
			fmt.Printf("%s ", key)
		}
		fmt.Printf("\n")
	}

	file, fileHeader, err := f.ctx.httpRequest.FormFile(fieldName)
	if err != nil {
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] FormFile error for field '%s': %v\n", fieldName, err)
		}
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("File field '%s' not found: %v", fieldName, err),
		}
	}
	defer file.Close()

	// Get file info
	fileName := strings.TrimSpace(fileHeader.Filename)
	fileSize := fileHeader.Size
	ext := strings.ToLower(filepath.Ext(fileName))

	// Validate extension
	if !f.isValidExtension(ext) {
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("File extension '%s' is not allowed", ext),
		}
	}

	// Check file size
	if f.maxFileSize > 0 && fileSize > f.maxFileSize {
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("File size exceeds maximum allowed size of %d bytes", f.maxFileSize),
		}
	}

	// Determine the final filename
	finalFileName := newFileName
	if finalFileName == "" {
		if f.preserveOriginalName {
			finalFileName = fileName
		} else {
			finalFileName = f.generateUniqueFileName(ext)
		}
	} else {
		// Ensure extension is preserved if not provided in newFileName
		if !strings.Contains(finalFileName, ".") {
			finalFileName = finalFileName + ext
		}
	}

	// Map target directory path
	mappedDir := f.ctx.Server_MapPath(targetDir)
	if mappedDir == "" {
		mappedDir = filepath.Join(f.ctx.RootDir, targetDir)
	}

	// Create directory if it doesn't exist
	os.MkdirAll(mappedDir, 0755)

	// Create temporary file first
	tempDir := filepath.Join(f.ctx.RootDir, "temp", "uploads")
	os.MkdirAll(tempDir, 0755)

	tempFile, err := os.CreateTemp(tempDir, "upload_*.tmp")
	if err != nil {
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("Failed to create temporary file: %v", err),
		}
	}
	tempPath := tempFile.Name()
	defer tempFile.Close()

	// Copy file to temporary location
	bytesWritten, err := io.Copy(tempFile, file)
	if err != nil {
		tempFile.Close()
		os.Remove(tempPath)
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] io.Copy error: %v (bytes written: %d)\n", err, bytesWritten)
		}
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("Failed to write temporary file: %v", err),
		}
	}

	// Ensure all data is written to disk
	err = tempFile.Sync()
	if err != nil {
		tempFile.Close()
		os.Remove(tempPath)
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] Sync error: %v\n", err)
		}
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("Failed to sync temporary file: %v", err),
		}
	}
	tempFile.Close()

	// Move to final location
	finalPath := filepath.Join(mappedDir, finalFileName)
	err = os.Rename(tempPath, finalPath)
	if err != nil {
		os.Remove(tempPath)
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("Failed to move file to final location: %v", err),
		}
	}

	// Detect MIME type
	mimeType := mime.TypeByExtension(ext)
	if mimeType == "" {
		mimeType = "application/octet-stream"
	}

	// Get relative path for response
	relPath, _ := filepath.Rel(f.ctx.RootDir, finalPath)

	return map[string]interface{}{
		"IsSuccess":        true,
		"OriginalFileName": fileName,
		"NewFileName":      finalFileName,
		"Size":             fileSize,
		"MimeType":         mimeType,
		"Extension":        ext,
		"TemporaryPath":    "",
		"FinalPath":        finalPath,
		"RelativePath":     relPath,
		"UploadedAt":       getCurrentTimeString(),
		"ErrorMessage":     "",
	}
}

// processAllUploads handles uploading all files from the request
func (f *G3FileUploader) processAllUploads(targetDir string) interface{} {
	// Parse multipart form with adaptive size limit based on maxFileSize
	var parseLimit int64 = 32 << 20 // 32MB default
	if f.maxFileSize > 0 {
		parseLimit = f.maxFileSize + (5 << 20) // Add 5MB buffer for form headers
	}

	err := f.ctx.httpRequest.ParseMultipartForm(parseLimit)
	if err != nil {
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] ParseMultipartForm error in ProcessAll: %v\n", err)
		}
		return make([]interface{}, 0)
	}

	results := make([]interface{}, 0)

	if f.ctx.httpRequest.MultipartForm == nil {
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] No multipart form in ProcessAll\n")
		}
		return results
	}

	if f.debugMode {
		fmt.Printf("[G3FileUploader DEBUG] ProcessAll - available fields: ")
		for fieldName := range f.ctx.httpRequest.MultipartForm.File {
			fmt.Printf("%s ", fieldName)
		}
		fmt.Printf("\n")
	}

	// Iterate through each file field and each file in that field
	for fieldName, fileHeaders := range f.ctx.httpRequest.MultipartForm.File {
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] ProcessAll - field %s has %d files\n", fieldName, len(fileHeaders))
		}

		// Process each file in this field
		for _, fileHeader := range fileHeaders {
			// We need to handle each file individually
			// Since processUpload uses FormFile which gets the first only,
			// we'll need to duplicate its logic here
			file, err := fileHeader.Open()
			if err != nil {
				results = append(results, map[string]interface{}{
					"IsSuccess":    false,
					"ErrorMessage": fmt.Sprintf("Failed to open file: %v", err),
				})
				continue
			}
			defer file.Close()

			// Get file info
			fileName := strings.TrimSpace(fileHeader.Filename)
			fileSize := fileHeader.Size
			ext := strings.ToLower(filepath.Ext(fileName))

			// Validate extension
			if !f.isValidExtension(ext) {
				results = append(results, map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     fmt.Sprintf("File extension '%s' is not allowed", ext),
				})
				continue
			}

			// Check file size
			if f.maxFileSize > 0 && fileSize > f.maxFileSize {
				results = append(results, map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     fmt.Sprintf("File size exceeds maximum allowed size of %d bytes", f.maxFileSize),
				})
				continue
			}

			// Determine the final filename
			finalFileName := ""
			if f.preserveOriginalName {
				finalFileName = fileName
			} else {
				finalFileName = f.generateUniqueFileName(ext)
			}

			// Map target directory path
			mappedDir := f.ctx.Server_MapPath(targetDir)
			if mappedDir == "" {
				mappedDir = filepath.Join(f.ctx.RootDir, targetDir)
			}

			// Create directory if it doesn't exist
			os.MkdirAll(mappedDir, 0755)

			// Create temporary file first
			tempDir := filepath.Join(f.ctx.RootDir, "temp", "uploads")
			os.MkdirAll(tempDir, 0755)

			tempFile, err := os.CreateTemp(tempDir, "upload_*.tmp")
			if err != nil {
				results = append(results, map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     fmt.Sprintf("Failed to create temporary file: %v", err),
				})
				continue
			}
			tempPath := tempFile.Name()

			// Copy file to temporary location
			bytesWritten, err := io.Copy(tempFile, file)
			if err != nil {
				tempFile.Close()
				os.Remove(tempPath)
				if f.debugMode {
					fmt.Printf("[G3FileUploader DEBUG] io.Copy error: %v (bytes written: %d)\n", err, bytesWritten)
				}
				results = append(results, map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     fmt.Sprintf("Failed to write temporary file: %v", err),
				})
				continue
			}

			// Ensure all data is written to disk
			err = tempFile.Sync()
			if err != nil {
				tempFile.Close()
				os.Remove(tempPath)
				if f.debugMode {
					fmt.Printf("[G3FileUploader DEBUG] Sync error: %v\n", err)
				}
				results = append(results, map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     fmt.Sprintf("Failed to sync temporary file: %v", err),
				})
				continue
			}
			tempFile.Close()

			// Move to final location
			finalPath := filepath.Join(mappedDir, finalFileName)
			err = os.Rename(tempPath, finalPath)
			if err != nil {
				os.Remove(tempPath)
				results = append(results, map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     fmt.Sprintf("Failed to move file to final location: %v", err),
				})
				continue
			}

			// Detect MIME type
			mimeType := mime.TypeByExtension(ext)
			if mimeType == "" {
				mimeType = "application/octet-stream"
			}

			// Get relative path for response
			relPath, _ := filepath.Rel(f.ctx.RootDir, finalPath)

			results = append(results, map[string]interface{}{
				"IsSuccess":        true,
				"OriginalFileName": fileName,
				"NewFileName":      finalFileName,
				"Size":             fileSize,
				"MimeType":         mimeType,
				"Extension":        ext,
				"TemporaryPath":    "",
				"FinalPath":        finalPath,
				"RelativePath":     relPath,
				"UploadedAt":       getCurrentTimeString(),
				"ErrorMessage":     "",
			})
		}
	}

	return results
}

// getFileInfo retrieves information about an uploaded file without saving it
func (f *G3FileUploader) getFileInfo(fieldName string) interface{} {
	// Parse multipart form with adaptive size limit based on maxFileSize
	var parseLimit int64 = 32 << 20 // 32MB default
	if f.maxFileSize > 0 {
		parseLimit = f.maxFileSize + (5 << 20) // Add 5MB buffer for form headers
	}

	err := f.ctx.httpRequest.ParseMultipartForm(parseLimit)
	if err != nil {
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] ParseMultipartForm error in GetFileInfo: %v\n", err)
		}
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("Failed to parse form data: %v", err),
		}
	}

	_, fileHeader, err := f.ctx.httpRequest.FormFile(fieldName)
	if err != nil {
		if f.debugMode {
			fmt.Printf("[G3FileUploader DEBUG] FormFile error for field '%s' in GetFileInfo: %v\n", fieldName, err)
		}
		return map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("File field '%s' not found: %v", fieldName, err),
		}
	}

	fileName := fileHeader.Filename
	fileSize := fileHeader.Size
	ext := strings.ToLower(filepath.Ext(fileName))
	mimeType := mime.TypeByExtension(ext)
	if mimeType == "" {
		mimeType = "application/octet-stream"
	}

	return map[string]interface{}{
		"OriginalFileName": fileName,
		"Size":             fileSize,
		"MimeType":         mimeType,
		"Extension":        ext,
		"IsValid":          f.isValidExtension(ext),
		"ExceedsMaxSize":   f.maxFileSize > 0 && fileSize > f.maxFileSize,
	}
}

// getAllFilesInfo retrieves information about all uploaded files
func (f *G3FileUploader) getAllFilesInfo() interface{} {
	f.ctx.httpRequest.ParseMultipartForm(32 << 20)

	results := make([]interface{}, 0)

	if f.ctx.httpRequest.MultipartForm != nil {
		for fieldName := range f.ctx.httpRequest.MultipartForm.File {
			info := f.getFileInfo(fieldName)
			results = append(results, info)
		}
	}

	return results
}

// generateUniqueFileName creates a unique filename based on timestamp and random string
func (f *G3FileUploader) generateUniqueFileName(ext string) string {
	timestamp := getCurrentTimeString()
	// Replace colons and dots with underscores for filesystem compatibility
	timestamp = strings.ReplaceAll(timestamp, ":", "")
	timestamp = strings.ReplaceAll(timestamp, "-", "")
	timestamp = strings.ReplaceAll(timestamp, " ", "_")

	// Generate a simple unique identifier
	randomPart := fmt.Sprintf("%d", os.Getpid())

	return fmt.Sprintf("upload_%s_%s%s", timestamp, randomPart, ext)
}

// getCurrentTimeString returns current time as formatted string
func getCurrentTimeString() string {
	return fmt.Sprintf("%04d-%02d-%02d %02d:%02d:%02d",
		time.Now().Year(),
		time.Now().Month(),
		time.Now().Day(),
		time.Now().Hour(),
		time.Now().Minute(),
		time.Now().Second(),
	)
}

// NewFileUploaderLibrary creates a new FileUploader library instance
func NewFileUploaderLibrary(ctx *ExecutionContext) *FileUploaderLibrary {
	return &FileUploaderLibrary{
		lib: &G3FileUploader{
			ctx:                  ctx,
			blockedExtensions:    make(map[string]bool),
			allowedExtensions:    make(map[string]bool),
			useAllowedExtOnly:    false,
			maxFileSize:          10 * 1024 * 1024, // 10MB default
			preserveOriginalName: false,
			debugMode:            false,
		},
	}
}

// FileUploaderLibrary wraps G3FileUploader for ASPLibrary interface compatibility
type FileUploaderLibrary struct {
	lib *G3FileUploader
}

// CallMethod calls a method on the FileUploader library
func (fu *FileUploaderLibrary) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return fu.lib.CallMethod(name, args...), nil
}

// GetProperty gets a property from the FileUploader library
func (fu *FileUploaderLibrary) GetProperty(name string) interface{} {
	return fu.lib.GetProperty(name)
}

// SetProperty sets a property on the FileUploader library
func (fu *FileUploaderLibrary) SetProperty(name string, value interface{}) error {
	fu.lib.SetProperty(name, value)
	return nil
}
