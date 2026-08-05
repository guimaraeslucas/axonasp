//go:build !wasm && !lib_g3fileuploader_disabled

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
	"image"
	_ "image/gif"
	_ "image/jpeg"
	_ "image/png"
	"io"
	"mime"
	"mime/multipart"
	"os"
	"path/filepath"
	"strings"
	"time"
)

type UploaderMode int

const (
	ModeG3FileUploader UploaderMode = iota
	ModePersitsUpload
	ModeSAFileUp
)

type fuKind int

const (
	fuKindUploader fuKind = iota
	fuKindFormCollection
	fuKindFilesCollection
	fuKindFile
)

type G3FileItem struct {
	fieldName   string
	fileName    string
	fileSize    int64
	contentType string
	fileHeader  *multipart.FileHeader
	savedPath   string
	saved       bool
	imageWidth  int
	imageHeight int
}

type G3FileUploader struct {
	vm                   *VM
	blockedExtensions    map[string]bool
	allowedExtensions    map[string]bool
	useAllowedExtOnly    bool
	maxFileSize          int64
	preserveOriginalName bool
	allowAbsolutePaths   bool
	debugMode            bool

	mode   UploaderMode
	kind   fuKind
	parent *G3FileUploader

	parsed         bool
	totalBytes     int64
	overwriteFiles bool
	logonUser      string
	progressID     string
	saPath         string
	maxBytes       int64

	formValues map[string][]string
	files      map[string][]*G3FileItem
	allFiles   []*G3FileItem
	allKeys    []string

	fileItem *G3FileItem
}

// newG3FileUploaderObject instantiates the G3FileUploader custom functions library.
func (vm *VM) newG3FileUploaderObject() Value {
	return vm.newG3FileUploaderObjectWithProgID("g3fileuploader")
}

// newG3FileUploaderObjectWithProgID instantiates the G3FileUploader custom functions library with alias mode.
func (vm *VM) newG3FileUploaderObjectWithProgID(progID string) Value {
	mode := ModeG3FileUploader
	progIDLower := strings.ToLower(progID)
	if strings.Contains(progIDLower, "persits") || strings.Contains(progIDLower, "aspupload") {
		mode = ModePersitsUpload
	} else if strings.Contains(progIDLower, "softartisans") || strings.Contains(progIDLower, "fileup") {
		mode = ModeSAFileUp
	}

	obj := &G3FileUploader{
		vm:                   vm,
		blockedExtensions:    make(map[string]bool),
		allowedExtensions:    make(map[string]bool),
		useAllowedExtOnly:    false,
		maxFileSize:          10 * 1024 * 1024, // 10MB default
		preserveOriginalName: false,
		allowAbsolutePaths:   mode != ModeG3FileUploader,
		debugMode:            false,
		mode:                 mode,
		kind:                 fuKindUploader,
		overwriteFiles:       true,
	}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.fileUploaderItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

func (f *G3FileUploader) newSubObject(kind fuKind, fileItem *G3FileItem) Value {
	sub := &G3FileUploader{
		vm:                   f.vm,
		blockedExtensions:    f.blockedExtensions,
		allowedExtensions:    f.allowedExtensions,
		useAllowedExtOnly:    f.useAllowedExtOnly,
		maxFileSize:          f.maxFileSize,
		preserveOriginalName: f.preserveOriginalName,
		allowAbsolutePaths:   f.allowAbsolutePaths,
		debugMode:            f.debugMode,
		mode:                 f.mode,
		kind:                 kind,
		parent:               f,
		fileItem:             fileItem,
	}
	id := f.vm.nextDynamicNativeID
	f.vm.nextDynamicNativeID++
	f.vm.fileUploaderItems[id] = sub
	return Value{Type: VTNativeObject, Num: id}
}

func (f *G3FileUploader) ensureParsed() error {
	if f.parent != nil {
		return f.parent.ensureParsed()
	}
	if f.parsed {
		return nil
	}
	f.parsed = true

	if f.vm.host == nil || f.vm.host.Request() == nil || f.vm.host.Request().HTTPRequest() == nil {
		return fmt.Errorf("no HTTP request context")
	}
	req := f.vm.host.Request().HTTPRequest()

	f.totalBytes = max(req.ContentLength, 0)

	var parseLimit int64 = 32 << 20
	if f.maxFileSize > 0 {
		parseLimit = f.maxFileSize + (5 << 20)
	}

	err := req.ParseMultipartForm(parseLimit)
	if err != nil {
		return err
	}

	if req.MultipartForm == nil {
		return fmt.Errorf("no multipart form data")
	}

	f.formValues = make(map[string][]string)
	f.files = make(map[string][]*G3FileItem)
	f.allFiles = []*G3FileItem{}
	f.allKeys = []string{}

	if req.MultipartForm.Value != nil {
		for k, vals := range req.MultipartForm.Value {
			f.formValues[strings.ToLower(k)] = vals
			f.allKeys = append(f.allKeys, k)
		}
	}

	if req.MultipartForm.File != nil {
		for fieldName, fileHeaders := range req.MultipartForm.File {
			var items []*G3FileItem
			for _, fh := range fileHeaders {
				item := &G3FileItem{
					fieldName:   fieldName,
					fileName:    fh.Filename,
					fileSize:    fh.Size,
					contentType: fh.Header.Get("Content-Type"),
					fileHeader:  fh,
				}
				items = append(items, item)
				f.allFiles = append(f.allFiles, item)
			}
			f.files[strings.ToLower(fieldName)] = items
			f.allKeys = append(f.allKeys, fieldName)
		}
	}

	return nil
}

func (f *G3FileUploader) EnumValues() (Value, error) {
	_ = f.ensureParsed()

	target := f
	if f.parent != nil {
		target = f.parent
	}

	if f.kind == fuKindFilesCollection {
		values := make([]Value, len(target.allFiles))
		for i, item := range target.allFiles {
			values[i] = target.newSubObject(fuKindFile, item)
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	} else if f.kind == fuKindFormCollection {
		values := make([]Value, len(target.allKeys))
		for i, k := range target.allKeys {
			values[i] = NewString(k)
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	} else if f.kind == fuKindUploader {
		if f.mode == ModeSAFileUp {
			values := make([]Value, len(target.allKeys))
			for i, k := range target.allKeys {
				values[i] = NewString(k)
			}
			return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
		}
		values := make([]Value, len(target.allFiles))
		for i, item := range target.allFiles {
			values[i] = target.newSubObject(fuKindFile, item)
		}
		return ValueFromVBArray(NewVBArrayFromValues(0, values)), nil
	}
	return ValueFromVBArray(NewVBArrayFromValues(0, nil)), nil
}

// DispatchPropertyGet acts as a getter.
func (f *G3FileUploader) DispatchPropertyGet(propertyName string) Value {
	prop := strings.ToLower(propertyName)

	switch f.kind {
	case fuKindUploader:
		switch prop {
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
		case "allowabsolutepaths":
			return NewBool(f.allowAbsolutePaths)
		case "formfields":
			return f.getFormFields()
		case "debugmode":
			return NewBool(f.debugMode)
		case "totalbytes":
			_ = f.ensureParsed()
			return NewInteger(f.totalBytes)
		case "overwritefiles":
			return NewBool(f.overwriteFiles)
		case "logonuser":
			return NewString(f.logonUser)
		case "progressid":
			return NewString(f.progressID)
		case "form":
			return f.newSubObject(fuKindFormCollection, nil)
		case "files":
			return f.newSubObject(fuKindFilesCollection, nil)
		case "path":
			return NewString(f.saPath)
		case "isempty":
			_ = f.ensureParsed()
			return NewBool(len(f.allFiles) == 0)
		case "maxbytes":
			return NewInteger(f.maxBytes)
		}

	case fuKindFormCollection:
		switch prop {
		case "count":
			_ = f.ensureParsed()
			target := f
			if f.parent != nil {
				target = f.parent
			}
			return NewInteger(int64(len(target.allKeys)))
		}

	case fuKindFilesCollection:
		switch prop {
		case "count":
			_ = f.ensureParsed()
			target := f
			if f.parent != nil {
				target = f.parent
			}
			return NewInteger(int64(len(target.allFiles)))
		}

	case fuKindFile:
		if f.fileItem == nil {
			return NewEmpty()
		}
		switch prop {
		case "path":
			return NewString(f.fileItem.savedPath)
		case "filename":
			return NewString(f.fileItem.fileName)
		case "ext":
			return NewString(filepath.Ext(f.fileItem.fileName))
		case "size":
			return NewInteger(f.fileItem.fileSize)
		case "contenttype":
			return NewString(f.fileItem.contentType)
		case "binary":
			return f.fileItem.getBinary(f.vm)
		case "imagewidth":
			f.fileItem.ensureImageDimensions()
			return NewInteger(int64(f.fileItem.imageWidth))
		case "imageheight":
			f.fileItem.ensureImageDimensions()
			return NewInteger(int64(f.fileItem.imageHeight))
		}
	}

	return f.DispatchMethod(propertyName, nil)
}

// DispatchPropertySet acts as a setter.
func (f *G3FileUploader) DispatchPropertySet(propertyName string, args []Value) bool {
	if len(args) == 0 {
		return false
	}
	val := args[0]
	prop := strings.ToLower(propertyName)

	switch f.kind {
	case fuKindUploader:
		switch prop {
		case "maxfilesize":
			f.maxFileSize = int64(f.vm.asInt(val))
			return true
		case "preserveoriginalname":
			f.preserveOriginalName = (val.Type == VTBool && val.Num != 0) || f.vm.asInt(val) != 0
			return true
		case "allowabsolutepaths":
			f.allowAbsolutePaths = (val.Type == VTBool && val.Num != 0) || f.vm.asInt(val) != 0
			return true
		case "debugmode":
			f.debugMode = (val.Type == VTBool && val.Num != 0) || f.vm.asInt(val) != 0
			return true
		case "overwritefiles":
			f.overwriteFiles = (val.Type == VTBool && val.Num != 0) || f.vm.asInt(val) != 0
			return true
		case "logonuser":
			f.logonUser = val.String()
			return true
		case "progressid":
			f.progressID = val.String()
			return true
		case "path":
			f.saPath = val.String()
			return true
		case "maxbytes":
			f.maxBytes = int64(f.vm.asInt(val))
			f.maxFileSize = f.maxBytes
			return true
		}
	}

	return false
}

// DispatchMethod provides O(1) string matching resolution.
func (f *G3FileUploader) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(methodName)

	switch f.kind {
	case fuKindUploader:
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

		case "process":
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

		case "processall":
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

		case "form":
			if len(args) > 0 {
				return f.getFormItem(args[0])
			}
			return f.newSubObject(fuKindFormCollection, nil)

		case "files":
			if len(args) > 0 {
				return f.getFileItem(args[0])
			}
			return f.newSubObject(fuKindFilesCollection, nil)

		case "formfields":
			return f.getFormFields()

		case "formvalue":
			if len(args) < 1 {
				return NewEmpty()
			}
			fieldName := args[0].String()
			_ = f.ensureParsed()
			if vals, exists := f.formValues[strings.ToLower(fieldName)]; exists && len(vals) > 0 {
				return NewString(vals[0])
			}
			return NewEmpty()

		case "isvalidextension":
			if len(args) < 1 {
				return NewBool(false)
			}
			ext := args[0].String()
			return NewBool(f.isValidExtension(ext))

		case "save":
			if f.mode == ModeSAFileUp || len(args) == 0 {
				_ = f.ensureParsed()
				savedCount := 0
				for _, item := range f.allFiles {
					err := f.saveFile(item, f.saPath, true)
					if err == nil {
						savedCount++
					}
				}
				return NewEmpty()
			} else {
				targetDir := args[0].String()
				_ = f.ensureParsed()
				savedCount := 0
				for _, item := range f.allFiles {
					err := f.saveFile(item, targetDir, f.overwriteFiles)
					if err == nil {
						savedCount++
					}
				}
				return NewInteger(int64(savedCount))
			}

		case "saveall":
			targetDir := "./"
			if len(args) > 0 {
				targetDir = args[0].String()
			}
			_ = f.ensureParsed()
			savedCount := 0
			for _, item := range f.allFiles {
				err := f.saveFile(item, targetDir, f.overwriteFiles)
				if err == nil {
					savedCount++
				}
			}
			return NewInteger(int64(savedCount))

		case "savevirtual":
			if len(args) < 1 {
				return NewInteger(0)
			}
			virtualDir := args[0].String()
			_ = f.ensureParsed()
			physicalDir := f.vm.host.Server().MapPath(virtualDir)
			savedCount := 0
			for _, item := range f.allFiles {
				err := f.saveFile(item, physicalDir, f.overwriteFiles)
				if err == nil {
					savedCount++
				}
			}
			return NewInteger(int64(savedCount))

		case "createdirectory":
			if len(args) < 1 {
				return NewEmpty()
			}
			dir := args[0].String()
			var physicalPath string
			if f.allowAbsolutePaths && filepath.IsAbs(dir) {
				physicalPath = filepath.Clean(dir)
			} else {
				physicalPath = f.vm.host.Server().MapPath(dir)
			}
			_ = os.MkdirAll(physicalPath, 0755)
			return NewEmpty()

		case "sendbinary", "transferfile":
			if len(args) < 1 {
				return NewEmpty()
			}
			path := args[0].String()
			contentType := ""
			if len(args) > 1 {
				contentType = args[1].String()
			}
			_ = f.SendBinary(path, contentType)
			return NewEmpty()

		case "setmaxsize":
			if len(args) > 0 {
				f.maxFileSize = int64(f.vm.asInt(args[0]))
				f.maxBytes = f.maxFileSize
			}
			return NewEmpty()

		case "flush":
			f.parsed = false
			f.formValues = nil
			f.files = nil
			f.allFiles = nil
			f.allKeys = nil
			return NewEmpty()
		}

	case fuKindFormCollection:
		switch method {
		case "", "item":
			if len(args) > 0 {
				return f.getFormItem(args[0])
			}
		}

	case fuKindFilesCollection:
		switch method {
		case "", "item":
			if len(args) > 0 {
				return f.getFileItem(args[0])
			}
		}

	case fuKindFile:
		if f.fileItem == nil {
			return NewEmpty()
		}
		switch method {
		case "saveas":
			if len(args) > 0 {
				path := args[0].String()
				var physicalPath string
				if f.allowAbsolutePaths && filepath.IsAbs(path) {
					physicalPath = filepath.Clean(path)
				} else {
					physicalPath = f.vm.host.Server().MapPath(path)
				}
				err := f.writeItemToPath(f.fileItem, physicalPath)
				if err == nil {
					f.fileItem.savedPath = physicalPath
					f.fileItem.saved = true
				}
			}
			return NewEmpty()

		case "saveasvirtual":
			if len(args) > 0 {
				virtualPath := args[0].String()
				physicalPath := f.vm.host.Server().MapPath(virtualPath)
				err := f.writeItemToPath(f.fileItem, physicalPath)
				if err == nil {
					f.fileItem.savedPath = physicalPath
					f.fileItem.saved = true
				}
			}
			return NewEmpty()

		case "delete":
			if f.fileItem.saved && f.fileItem.savedPath != "" {
				_ = os.Remove(f.fileItem.savedPath)
				f.fileItem.saved = false
				f.fileItem.savedPath = ""
			}
			return NewEmpty()

		case "copy":
			if len(args) > 0 {
				path := args[0].String()
				var physicalPath string
				if f.allowAbsolutePaths && filepath.IsAbs(path) {
					physicalPath = filepath.Clean(path)
				} else {
					physicalPath = f.vm.host.Server().MapPath(path)
				}
				if f.fileItem.saved && f.fileItem.savedPath != "" {
					src, err := os.Open(f.fileItem.savedPath)
					if err == nil {
						defer src.Close()
						dir := filepath.Dir(physicalPath)
						_ = os.MkdirAll(dir, 0755)
						dst, err := os.Create(physicalPath)
						if err == nil {
							defer dst.Close()
							_, _ = io.Copy(dst, src)
						}
					}
				} else {
					_ = f.writeItemToPath(f.fileItem, physicalPath)
				}
			}
			return NewEmpty()
		}
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
	parts := strings.SplitSeq(exts, ",")
	for part := range parts {
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
	parts := strings.SplitSeq(exts, ",")
	for part := range parts {
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

func (f *G3FileUploader) wrapResultAsDict(m map[string]any) Value {
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

// moveUploadedFile relocates a fully written temp upload file to its final
// destination. It prefers an atomic os.Rename when both paths share the same
// volume, and falls back to a streaming copy+delete when the temp directory and
// the target directory live on different drives (os.Rename cannot cross volumes
// on Windows). A short retry loop absorbs transient sharing violations raised by
// antivirus scanners right after the temp file is closed.
func moveUploadedFile(tempPath, finalPath string) error {
	const maxAttempts = 3
	var err error
	for i := range maxAttempts {
		err = os.Rename(tempPath, finalPath)
		if err == nil {
			return nil
		}
		// Brief pause to let transient locks (AV scans, editors) be released
		// before retrying or falling back to a cross-volume copy.
		time.Sleep(time.Duration(50*(i+1)) * time.Millisecond)
	}
	return copyUploadedFile(tempPath, finalPath)
}

// copyUploadedFile streams the temp file to finalPath and removes the source,
// providing a cross-volume-safe fallback when os.Rename is not possible. All
// handles are closed before the temp file is removed so the cleanup never
// trips a Windows sharing violation on its own open handle.
func copyUploadedFile(tempPath, finalPath string) error {
	src, err := os.Open(tempPath)
	if err != nil {
		return err
	}

	dst, err := os.Create(finalPath)
	if err != nil {
		src.Close()
		return err
	}

	_, copyErr := io.Copy(dst, src)
	srcCloseErr := src.Close()
	syncErr := dst.Sync()
	dstCloseErr := dst.Close()

	if copyErr != nil {
		os.Remove(finalPath)
		return copyErr
	}
	if syncErr != nil {
		return syncErr
	}
	if dstCloseErr != nil {
		return dstCloseErr
	}
	if srcCloseErr != nil {
		return srcCloseErr
	}

	// The destination file is fully written, synced, and closed; the source temp
	// file is closed too. Best-effort removal with a short retry loop absorbs
	// transient sharing violations (e.g. antivirus scanners) without failing an
	// upload whose final destination was already persisted.
	for i := range 3 {
		if err := os.Remove(tempPath); err == nil {
			return nil
		} else if i < 2 {
			time.Sleep(time.Duration(50*(i+1)) * time.Millisecond)
		}
	}
	return nil
}

func (f *G3FileUploader) processUpload(fieldName, targetDir, newFileName string) Value {
	if f.vm.host == nil || f.vm.host.Request() == nil || f.vm.host.Request().HTTPRequest() == nil {
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": ErrG3FUNoHTTPRequest.String(),
		})
	}

	req := f.vm.host.Request().HTTPRequest()

	var parseLimit int64 = 32 << 20
	if f.maxFileSize > 0 {
		parseLimit = f.maxFileSize + (5 << 20)
	}

	err := req.ParseMultipartForm(parseLimit)
	if err != nil {
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf("%s: %v", ErrG3FUFormParseFailed.String(), err),
		})
	}

	if req.MultipartForm == nil {
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": ErrG3FUNoMultipartData.String(),
		})
	}

	file, fileHeader, err := req.FormFile(fieldName)
	if err != nil {
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf(ErrG3FUFileFieldNotFound.String(), fieldName),
		})
	}
	defer file.Close()

	fileName := strings.TrimSpace(fileHeader.Filename)
	fileSize := fileHeader.Size
	ext := strings.ToLower(filepath.Ext(fileName))

	if !f.isValidExtension(ext) {
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf(ErrG3FUExtensionNotAllowed.String(), ext),
		})
	}

	if f.maxFileSize > 0 && fileSize > f.maxFileSize {
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": ErrG3FUFileSizeExceedsMax.String(),
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

	var mappedDir string
	if f.allowAbsolutePaths && (filepath.IsAbs(targetDir) || strings.HasPrefix(targetDir, "\\\\")) {
		mappedDir = filepath.Clean(targetDir)
	} else {
		mappedDir = f.vm.host.Server().MapPath(targetDir)
	}

	if mappedDir == "" {
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": ErrG3FUInvalidTargetDir.String(),
		})
	}

	os.MkdirAll(mappedDir, 0755)

	tempDir := filepath.Join(resolveConfiguredTempDir(), "uploads")
	if tempDir == "" {
		tempDir = os.TempDir()
	}
	os.MkdirAll(tempDir, 0755)

	tempFile, err := os.CreateTemp(tempDir, "upload_*.tmp")
	if err != nil {
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": ErrG3FUTempFileCreateFailed.String(),
		})
	}
	tempPath := tempFile.Name()
	defer tempFile.Close()

	_, err = io.Copy(tempFile, file)
	if err != nil {
		tempFile.Close()
		os.Remove(tempPath)
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": ErrG3FUTempFileWriteFailed.String(),
		})
	}

	err = tempFile.Sync()
	if err != nil {
		tempFile.Close()
		os.Remove(tempPath)
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": ErrG3FUTempFileSyncFailed.String(),
		})
	}
	tempFile.Close()

	finalPath := filepath.Join(mappedDir, finalFileName)
	err = moveUploadedFile(tempPath, finalPath)
	if err != nil {
		os.Remove(tempPath)
		return f.wrapResultAsDict(map[string]any{
			"IsSuccess":    false,
			"ErrorMessage": fmt.Sprintf(ErrG3FUFinalMoveFailedDetail.String(), err),
		})
	}

	mimeType := mime.TypeByExtension(ext)
	if mimeType == "" {
		mimeType = "application/octet-stream"
	}

	relPath := targetDir + "/" + finalFileName

	return f.wrapResultAsDict(map[string]any{
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

	var mappedDir string
	if f.allowAbsolutePaths && (filepath.IsAbs(targetDir) || strings.HasPrefix(targetDir, "\\\\")) {
		mappedDir = filepath.Clean(targetDir)
	} else {
		mappedDir = f.vm.host.Server().MapPath(targetDir)
	}

	for _, fileHeaders := range req.MultipartForm.File {
		for _, fileHeader := range fileHeaders {
			file, err := fileHeader.Open()
			if err != nil {
				results = append(results, f.wrapResultAsDict(map[string]any{
					"IsSuccess":    false,
					"ErrorMessage": ErrG3FUOpenFileFailed.String(),
				}))
				continue
			}

			fileName := strings.TrimSpace(fileHeader.Filename)
			fileSize := fileHeader.Size
			ext := strings.ToLower(filepath.Ext(fileName))

			if !f.isValidExtension(ext) {
				results = append(results, f.wrapResultAsDict(map[string]any{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     fmt.Sprintf(ErrG3FUExtensionNotAllowed.String(), ext),
				}))
				file.Close()
				continue
			}

			if f.maxFileSize > 0 && fileSize > f.maxFileSize {
				results = append(results, f.wrapResultAsDict(map[string]any{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     ErrG3FUFileSizeExceedsMax.String(),
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

			if mappedDir == "" {
				results = append(results, f.wrapResultAsDict(map[string]any{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     ErrG3FUInvalidTargetDir.String(),
				}))
				file.Close()
				continue
			}

			os.MkdirAll(mappedDir, 0755)

			tempDir := filepath.Join(resolveConfiguredTempDir(), "uploads")
			if tempDir == "" {
				tempDir = os.TempDir()
			}
			os.MkdirAll(tempDir, 0755)

			tempFile, err := os.CreateTemp(tempDir, "upload_*.tmp")
			if err != nil {
				results = append(results, f.wrapResultAsDict(map[string]any{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     ErrG3FUTempFileCreateFailed.String(),
				}))
				file.Close()
				continue
			}
			tempPath := tempFile.Name()

			_, err = io.Copy(tempFile, file)
			if err != nil {
				tempFile.Close()
				os.Remove(tempPath)
				results = append(results, f.wrapResultAsDict(map[string]any{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     ErrG3FUTempFileWriteFailed.String(),
				}))
				file.Close()
				continue
			}

			err = tempFile.Sync()
			if err != nil {
				tempFile.Close()
				os.Remove(tempPath)
				results = append(results, f.wrapResultAsDict(map[string]any{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     ErrG3FUTempFileSyncFailed.String(),
				}))
				file.Close()
				continue
			}
			tempFile.Close()
			file.Close()

			finalPath := filepath.Join(mappedDir, finalFileName)
			err = moveUploadedFile(tempPath, finalPath)
			if err != nil {
				os.Remove(tempPath)
				results = append(results, f.wrapResultAsDict(map[string]any{
					"IsSuccess":        false,
					"OriginalFileName": fileName,
					"ErrorMessage":     fmt.Sprintf(ErrG3FUFinalMoveFailedDetail.String(), err),
				}))
				continue
			}

			mimeType := mime.TypeByExtension(ext)
			if mimeType == "" {
				mimeType = "application/octet-stream"
			}

			relPath := targetDir + "/" + finalFileName

			results = append(results, f.wrapResultAsDict(map[string]any{
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
		return f.wrapResultAsDict(map[string]any{
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

	return f.wrapResultAsDict(map[string]any{
		"IsSuccess":        true,
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

func (f *G3FileUploader) getFormFields() Value {
	dictVal := f.vm.newDictionaryObject()
	if f.vm.host == nil || f.vm.host.Request() == nil || f.vm.host.Request().HTTPRequest() == nil {
		return dictVal
	}

	req := f.vm.host.Request().HTTPRequest()
	if req.MultipartForm == nil {
		var parseLimit int64 = 32 << 20
		if f.maxFileSize > 0 {
			parseLimit = f.maxFileSize + (5 << 20)
		}
		_ = req.ParseMultipartForm(parseLimit)
	}

	if req.MultipartForm != nil && req.MultipartForm.Value != nil {
		for k, vals := range req.MultipartForm.Value {
			if len(vals) > 0 {
				f.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString(k), NewString(vals[0])})
			}
		}
	}

	return dictVal
}

func (f *G3FileUploader) generateUniqueFileName(ext string) string {
	timestamp := fmt.Sprintf("%04d%02d%02d_%02d%02d%02d",
		time.Now().Year(), time.Now().Month(), time.Now().Day(),
		time.Now().Hour(), time.Now().Minute(), time.Now().Second(),
	)
	randomPart := fmt.Sprintf("%d", os.Getpid())
	return fmt.Sprintf("upload_%s_%s%s", timestamp, randomPart, ext)
}

func (f *G3FileUploader) SendBinary(path string, contentType string) error {
	var physicalPath string
	if filepath.IsAbs(path) {
		physicalPath = filepath.Clean(path)
	} else {
		physicalPath = f.vm.host.Server().MapPath(path)
	}

	data, err := os.ReadFile(physicalPath)
	if err != nil {
		return err
	}

	if contentType != "" {
		f.vm.host.Response().SetContentType(contentType)
	} else {
		ext := strings.ToLower(filepath.Ext(physicalPath))
		mimeType := mime.TypeByExtension(ext)
		if mimeType != "" {
			f.vm.host.Response().SetContentType(mimeType)
		} else {
			f.vm.host.Response().SetContentType("application/octet-stream")
		}
	}

	f.vm.host.Response().BinaryWrite(data)
	return nil
}

func (f *G3FileUploader) saveFile(item *G3FileItem, targetDir string, overwrite bool) error {
	if item.fileHeader == nil {
		return fmt.Errorf("no file header")
	}

	var mappedDir string
	if f.allowAbsolutePaths && (filepath.IsAbs(targetDir) || strings.HasPrefix(targetDir, "\\\\")) {
		mappedDir = filepath.Clean(targetDir)
	} else {
		mappedDir = f.vm.host.Server().MapPath(targetDir)
	}

	if mappedDir == "" {
		return fmt.Errorf("invalid target directory")
	}

	err := os.MkdirAll(mappedDir, 0755)
	if err != nil {
		return err
	}

	fileName := item.fileName
	finalPath := filepath.Join(mappedDir, fileName)

	if _, err := os.Stat(finalPath); err == nil && !overwrite {
		ext := filepath.Ext(fileName)
		base := strings.TrimSuffix(fileName, ext)
		counter := 1
		for {
			newName := fmt.Sprintf("%s(%d)%s", base, counter, ext)
			newPath := filepath.Join(mappedDir, newName)
			if _, err := os.Stat(newPath); os.IsNotExist(err) {
				finalPath = newPath
				item.fileName = newName
				break
			}
			counter++
		}
	}

	srcFile, err := item.fileHeader.Open()
	if err != nil {
		return err
	}
	defer srcFile.Close()

	dstFile, err := os.Create(finalPath)
	if err != nil {
		return err
	}
	defer dstFile.Close()

	_, err = io.Copy(dstFile, srcFile)
	if err != nil {
		return err
	}

	item.savedPath = finalPath
	item.saved = true
	return nil
}

func (f *G3FileUploader) writeItemToPath(item *G3FileItem, destPath string) error {
	if item.fileHeader == nil {
		return fmt.Errorf("no file data")
	}
	dir := filepath.Dir(destPath)
	err := os.MkdirAll(dir, 0755)
	if err != nil {
		return err
	}

	src, err := item.fileHeader.Open()
	if err != nil {
		return err
	}
	defer src.Close()

	dst, err := os.Create(destPath)
	if err != nil {
		return err
	}
	defer dst.Close()

	_, err = io.Copy(dst, src)
	return err
}

func (f *G3FileUploader) getFormItem(key Value) Value {
	_ = f.ensureParsed()
	target := f
	if f.parent != nil {
		target = f.parent
	}

	if key.Type == VTInteger || key.Type == VTDouble {
		idx := int(f.vm.asInt(key)) - 1
		if idx >= 0 && idx < len(target.allKeys) {
			fieldName := target.allKeys[idx]
			return f.getFormItemByName(fieldName)
		}
		return NewEmpty()
	}
	return f.getFormItemByName(key.String())
}

func (f *G3FileUploader) getFormItemByName(name string) Value {
	_ = f.ensureParsed()
	target := f
	if f.parent != nil {
		target = f.parent
	}

	nameLower := strings.ToLower(name)
	if target.mode == ModeSAFileUp {
		if items, exists := target.files[nameLower]; exists && len(items) > 0 {
			return target.newSubObject(fuKindFile, items[0])
		}
	}
	if vals, exists := target.formValues[nameLower]; exists && len(vals) > 0 {
		return NewString(vals[0])
	}
	return NewEmpty()
}

func (f *G3FileUploader) getFileItem(key Value) Value {
	_ = f.ensureParsed()
	target := f
	if f.parent != nil {
		target = f.parent
	}

	if key.Type == VTInteger || key.Type == VTDouble {
		idx := int(f.vm.asInt(key)) - 1
		if idx >= 0 && idx < len(target.allFiles) {
			return target.newSubObject(fuKindFile, target.allFiles[idx])
		}
		return NewEmpty()
	}
	nameLower := strings.ToLower(key.String())
	if items, exists := target.files[nameLower]; exists && len(items) > 0 {
		return target.newSubObject(fuKindFile, items[0])
	}
	return NewEmpty()
}

func (item *G3FileItem) getBinary(vm *VM) Value {
	if item.fileHeader == nil {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}
	f, err := item.fileHeader.Open()
	if err != nil {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}
	defer f.Close()
	data, err := io.ReadAll(f)
	if err != nil {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}

	values := make([]Value, len(data))
	for i, b := range data {
		values[i] = NewInteger(int64(b))
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
}

func (item *G3FileItem) ensureImageDimensions() {
	if item.imageWidth > 0 || item.imageHeight > 0 {
		return
	}
	if item.fileHeader == nil {
		return
	}
	f, err := item.fileHeader.Open()
	if err != nil {
		return
	}
	defer f.Close()

	cfg, _, err := image.DecodeConfig(f)
	if err == nil {
		item.imageWidth = cfg.Width
		item.imageHeight = cfg.Height
	}
}
