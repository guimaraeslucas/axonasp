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
	"fmt"
	"io"
	"mime"
	"os"
	"path/filepath"
	"strings"
	"time"
)

type G3FileUploader struct {
	vm                   *VM
	blockedExtensions    map[string]bool
	allowedExtensions    map[string]bool
	useAllowedExtOnly    bool
	maxFileSize          int64
	preserveOriginalName bool
	debugMode            bool
}

// newG3FileUploaderObject instantiates the G3FileUploader custom functions library.
func (vm *VM) newG3FileUploaderObject() Value {
	obj := &G3FileUploader{
		vm:                   vm,
		blockedExtensions:    make(map[string]bool),
		allowedExtensions:    make(map[string]bool),
		useAllowedExtOnly:    false,
		maxFileSize:          10 * 1024 * 1024, // 10MB default
		preserveOriginalName: false,
		debugMode:            false,
	}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.fileUploaderItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (f *G3FileUploader) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "blockedextensions":
		extList := make([]Value, 0)
		for ext := range f.blockedExtensions {
			extList = append(extList, NewString(ext))
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, extList)}
	case "allowedextensions":
		extList := make([]Value, 0)
		for ext := range f.allowedExtensions {
			extList = append(extList, NewString(ext))
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, extList)}
	case "maxfilesize":
		return NewInteger(f.maxFileSize)
	case "preserveoriginalname":
		return NewBool(f.preserveOriginalName)
	case "debugmode":
		return NewBool(f.debugMode)
	}
	return f.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet acts as a setter.
func (f *G3FileUploader) DispatchPropertySet(propertyName string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	val := args[0]
	switch strings.ToLower(propertyName) {
	case "maxfilesize":
		f.maxFileSize = int64(f.vm.asInt(val))
		return true
	case "preserveoriginalname":
		f.preserveOriginalName = (val.Type == VTBool && val.Num != 0) || f.vm.asInt(val) != 0
		return true
	case "debugmode":
		f.debugMode = (val.Type == VTBool && val.Num != 0) || f.vm.asInt(val) != 0
		return true
	}
	return false
}

// DispatchMethod provides O(1) string matching resolution.
func (f *G3FileUploader) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(methodName)

	switch method {
	case "blockextension":
		if len(args) > 0 {
			f.blockExtension(args[0].String())
		}
		return NewEmpty()

	case "allowextension":
		if len(args) > 0 {
			f.allowExtension(args[0].String())
		}
		return NewEmpty()

	case "blockextensions":
		if len(args) > 0 {
			f.blockExtensions(args[0].String())
		}
		return NewEmpty()

	case "allowextensions":
		if len(args) > 0 {
			f.allowExtensions(args[0].String())
		}
		return NewEmpty()

	case "setuseallowedonly":
		if len(args) > 0 {
			val := args[0]
			f.useAllowedExtOnly = (val.Type == VTBool && val.Num != 0) || f.vm.asInt(val) != 0
		}
		return NewEmpty()

	case "process", "save":
		if len(args) < 1 {
			return NewEmpty()
		}
		fieldName := args[0].String()
		targetDir := "./"
		if len(args) > 1 {
			targetDir = args[1].String()
		}
		newFileName := ""
		if len(args) > 2 {
			newFileName = args[2].String()
		}
		return f.processUpload(fieldName, targetDir, newFileName)

	case "processall", "saveall":
		targetDir := "./"
		if len(args) > 0 {
			targetDir = args[0].String()
		}
		return f.processAllUploads(targetDir)

	case "getfileinfo":
		if len(args) < 1 {
			return NewEmpty()
		}
		fieldName := args[0].String()
		return f.getFileInfo(fieldName)

	case "getallfilesinfo":
		return f.getAllFilesInfo()

	case "form", "formvalue":
		if len(args) < 1 {
			return NewEmpty()
		}
		fieldName := args[0].String()
		if f.vm.host != nil && f.vm.host.Request() != nil && f.vm.host.Request().HTTPRequest() != nil {
			req := f.vm.host.Request().HTTPRequest()
			if req.MultipartForm == nil {
				// Attempt to parse if not already parsed
				var parseLimit int64 = 32 << 20
				if f.maxFileSize > 0 {
					parseLimit = f.maxFileSize + (5 << 20)
				}
				_ = req.ParseMultipartForm(parseLimit)
			}
			if req.MultipartForm != nil && req.MultipartForm.Value != nil {
				if vals, ok := req.MultipartForm.Value[fieldName]; ok && len(vals) > 0 {
					return NewString(vals[0])
				}
			}
		}
		return NewEmpty()

	case "isvalidextension":
		if len(args) < 1 {
			return NewBool(false)
		}
		ext := args[0].String()
		return NewBool(f.isValidExtension(ext))
	}

	return NewEmpty()
}

func (f *G3FileUploader) blockExtension(ext string) {
	extStr := strings.ToLower(strings.TrimSpace(ext))
	if !strings.HasPrefix(extStr, ".") {
		extStr = "." + extStr
	}
	f.blockedExtensions[extStr] = true
}

func (f *G3FileUploader) blockExtensions(exts string) {
	parts := strings.Split(exts, ",")
	for _, part := range parts {
		f.blockExtension(strings.TrimSpace(part))
	}
}

func (f *G3FileUploader) allowExtension(ext string) {
	extStr := strings.ToLower(strings.TrimSpace(ext))
	if !strings.HasPrefix(extStr, ".") {
		extStr = "." + extStr
	}
	f.allowedExtensions[extStr] = true
}

func (f *G3FileUploader) allowExtensions(exts string) {
	parts := strings.Split(exts, ",")
	for _, part := range parts {
		f.allowExtension(strings.TrimSpace(part))
	}
}

func (f *G3FileUploader) isValidExtension(ext string) bool {
	ext = strings.ToLower(strings.TrimSpace(ext))
	if !strings.HasPrefix(ext, ".") {
		ext = "." + ext
	}
	if f.blockedExtensions[ext] {
		return false
	}
	if f.useAllowedExtOnly {
		return f.allowedExtensions[ext]
	}
	return true
}

func (f *G3FileUploader) wrapResultAsDict(m map[string]interface{}) Value {
	dictVal := f.vm.newDictionaryObject()
	for k, v := range m {
		var mapped Value
		switch val := v.(type) {
		case string:
			mapped = NewString(val)
		case bool:
			mapped = NewBool(val)
		case int:
			mapped = NewInteger(int64(val))
		case int64:
			mapped = NewInteger(val)
		case float64:
			mapped = NewDouble(val)
		default:
			mapped = NewString(fmt.Sprintf("%v", v))
		}
		f.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString(k), mapped})
	}
	return dictVal
}

func (f *G3FileUploader) processUpload(fieldName, targetDir, newFileName string) Value {
	if f.vm.host == nil || f.vm.host.Request() == nil || f.vm.host.Request().HTTPRequest() == nil {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "No HTTP request context",
		})
	}

	req := f.vm.host.Request().HTTPRequest()

	var parseLimit int64 = 32 << 20
	if f.maxFileSize > 0 {
		parseLimit = f.maxFileSize + (5 << 20)
	}

	err := req.ParseMultipartForm(parseLimit)
	if err != nil {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("Failed to parse form data: %v", err),
		})
	}

	if req.MultipartForm == nil {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "No multipart form data received",
		})
	}

	file, fileHeader, err := req.FormFile(fieldName)
	if err != nil {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("File field '%s' not found", fieldName),
		})
	}
	defer file.Close()

	fileName := strings.TrimSpace(fileHeader.Filename)
	fileSize := fileHeader.Size
	ext := strings.ToLower(filepath.Ext(fileName))

	if !f.isValidExtension(ext) {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("File extension '%s' is not allowed", ext),
		})
	}

	if f.maxFileSize > 0 && fileSize > f.maxFileSize {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "File size exceeds maximum allowed size",
		})
	}

	finalFileName := newFileName
	if finalFileName == "" {
		if f.preserveOriginalName {
			finalFileName = fileName
		} else {
			finalFileName = f.generateUniqueFileName(ext)
		}
	} else {
		if !strings.Contains(finalFileName, ".") {
			finalFileName = finalFileName + ext
		}
	}

	mappedDir := f.vm.host.Server().MapPath(targetDir)
	if mappedDir == "" {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "Invalid target directory",
		})
	}

	os.MkdirAll(mappedDir, 0755)

	// Since we don't have access to RootDir directly from host interface here without exposing it,
	// let's just use os.TempDir or Server().MapPath("temp/uploads")
	tempDir := f.vm.host.Server().MapPath("/temp/uploads")
	if tempDir == "" {
		tempDir = os.TempDir()
	}
	os.MkdirAll(tempDir, 0755)

	tempFile, err := os.CreateTemp(tempDir, "upload_*.tmp")
	if err != nil {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "Failed to create temporary file",
		})
	}
	tempPath := tempFile.Name()
	defer tempFile.Close()

	_, err = io.Copy(tempFile, file)
	if err != nil {
		tempFile.Close()
		os.Remove(tempPath)
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "Failed to write temporary file",
		})
	}

	err = tempFile.Sync()
	if err != nil {
		tempFile.Close()
		os.Remove(tempPath)
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "Failed to sync temporary file",
		})
	}
	tempFile.Close()

	finalPath := filepath.Join(mappedDir, finalFileName)
	err = os.Rename(tempPath, finalPath)
	if err != nil {
		os.Remove(tempPath)
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "Failed to move file to final location",
		})
	}

	mimeType := mime.TypeByExtension(ext)
	if mimeType == "" {
		mimeType = "application/octet-stream"
	}

	relPath := targetDir + "/" + finalFileName

	return f.wrapResultAsDict(map[string]interface{}{
		"IsSuccess":        true,
		"OriginalFileName": fileName,
		"NewFileName":      finalFileName,
		"Size":             fileSize,
		"MimeType":         mimeType,
		"Extension":        ext,
		"FinalPath":        finalPath,
		"RelativePath":     relPath,
		"ErrorMessage":     "",
	})
}

func (f *G3FileUploader) processAllUploads(targetDir string) Value {
	if f.vm.host == nil || f.vm.host.Request() == nil || f.vm.host.Request().HTTPRequest() == nil {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}

	req := f.vm.host.Request().HTTPRequest()

	var parseLimit int64 = 32 << 20
	if f.maxFileSize > 0 {
		parseLimit = f.maxFileSize + (5 << 20)
	}

	_ = req.ParseMultipartForm(parseLimit)

	var results []Value

	if req.MultipartForm == nil || req.MultipartForm.File == nil {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, results)}
	}

	for _, fileHeaders := range req.MultipartForm.File {
		for _, fileHeader := range fileHeaders {
			file, err := fileHeader.Open()
			if err != nil {
				results = append(results, f.wrapResultAsDict(map[string]interface{}{
					"IsSuccess":    false,
					"ErrorMessage": "Failed to open file",
				}))
				continue
			}

			fileName := strings.TrimSpace(fileHeader.Filename)
			fileSize := fileHeader.Size
			ext := strings.ToLower(filepath.Ext(fileName))

			if !f.isValidExtension(ext) {
				results = append(results, f.wrapResultAsDict(map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     "File extension not allowed",
				}))
				file.Close()
				continue
			}

			if f.maxFileSize > 0 && fileSize > f.maxFileSize {
				results = append(results, f.wrapResultAsDict(map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     "File size exceeds maximum",
				}))
				file.Close()
				continue
			}

			finalFileName := ""
			if f.preserveOriginalName {
				finalFileName = fileName
			} else {
				finalFileName = f.generateUniqueFileName(ext)
			}

			mappedDir := f.vm.host.Server().MapPath(targetDir)
			if mappedDir == "" {
				results = append(results, f.wrapResultAsDict(map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     "Invalid target directory",
				}))
				file.Close()
				continue
			}

			os.MkdirAll(mappedDir, 0755)

			tempDir := f.vm.host.Server().MapPath("/temp/uploads")
			if tempDir == "" {
				tempDir = os.TempDir()
			}
			os.MkdirAll(tempDir, 0755)

			tempFile, err := os.CreateTemp(tempDir, "upload_*.tmp")
			if err != nil {
				results = append(results, f.wrapResultAsDict(map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     "Failed to create temporary file",
				}))
				file.Close()
				continue
			}
			tempPath := tempFile.Name()

			_, err = io.Copy(tempFile, file)
			if err != nil {
				tempFile.Close()
				os.Remove(tempPath)
				results = append(results, f.wrapResultAsDict(map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     "Failed to write temporary file",
				}))
				file.Close()
				continue
			}

			err = tempFile.Sync()
			if err != nil {
				tempFile.Close()
				os.Remove(tempPath)
				results = append(results, f.wrapResultAsDict(map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     "Failed to sync temporary file",
				}))
				file.Close()
				continue
			}
			tempFile.Close()
			file.Close()

			finalPath := filepath.Join(mappedDir, finalFileName)
			err = os.Rename(tempPath, finalPath)
			if err != nil {
				os.Remove(tempPath)
				results = append(results, f.wrapResultAsDict(map[string]interface{}{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     "Failed to move file to final location",
				}))
				continue
			}

			mimeType := mime.TypeByExtension(ext)
			if mimeType == "" {
				mimeType = "application/octet-stream"
			}

			relPath := targetDir + "/" + finalFileName

			results = append(results, f.wrapResultAsDict(map[string]interface{}{
				"IsSuccess":        true,
				"OriginalFileName": fileName,
				"NewFileName":      finalFileName,
				"Size":             fileSize,
				"MimeType":         mimeType,
				"Extension":        ext,
				"FinalPath":        finalPath,
				"RelativePath":     relPath,
				"ErrorMessage":     "",
			}))
		}
	}

	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, results)}
}

func (f *G3FileUploader) getFileInfo(fieldName string) Value {
	if f.vm.host == nil || f.vm.host.Request() == nil || f.vm.host.Request().HTTPRequest() == nil {
		return NewEmpty()
	}

	req := f.vm.host.Request().HTTPRequest()
	_ = req.ParseMultipartForm(32 << 20)

	_, fileHeader, err := req.FormFile(fieldName)
	if err != nil {
		return f.wrapResultAsDict(map[string]interface{}{
			"IsSuccess":    false,
			"ErrorMessage": "File not found",
		})
	}

	fileName := fileHeader.Filename
	fileSize := fileHeader.Size
	ext := strings.ToLower(filepath.Ext(fileName))
	mimeType := mime.TypeByExtension(ext)
	if mimeType == "" {
		mimeType = "application/octet-stream"
	}

	return f.wrapResultAsDict(map[string]interface{}{
		"OriginalFileName": fileName,
		"Size":             fileSize,
		"MimeType":         mimeType,
		"Extension":        ext,
		"IsValid":          f.isValidExtension(ext),
		"ExceedsMaxSize":   f.maxFileSize > 0 && fileSize > f.maxFileSize,
	})
}

func (f *G3FileUploader) getAllFilesInfo() Value {
	if f.vm.host == nil || f.vm.host.Request() == nil || f.vm.host.Request().HTTPRequest() == nil {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}

	req := f.vm.host.Request().HTTPRequest()
	_ = req.ParseMultipartForm(32 << 20)

	var results []Value

	if req.MultipartForm != nil && req.MultipartForm.File != nil {
		for fieldName := range req.MultipartForm.File {
			results = append(results, f.getFileInfo(fieldName))
		}
	}

	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, results)}
}

func (f *G3FileUploader) generateUniqueFileName(ext string) string {
	timestamp := fmt.Sprintf("%04d%02d%02d_%02d%02d%02d",
		time.Now().Year(), time.Now().Month(), time.Now().Day(),
		time.Now().Hour(), time.Now().Minute(), time.Now().Second(),
	)
	randomPart := fmt.Sprintf("%d", os.Getpid())
	return fmt.Sprintf("upload_%s_%s%s", timestamp, randomPart, ext)
}
