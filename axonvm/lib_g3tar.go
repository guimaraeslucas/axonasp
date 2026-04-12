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
	"archive/tar"
	"bytes"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"time"

	"g3pix.com.br/axonasp/vbscript"
)

// g3tarEntry stores one indexed header for read-mode operations.
type g3tarEntry struct {
	Name     string
	Size     int64
	Mode     int64
	ModTime  time.Time
	TypeFlag byte
}

// G3TAR wraps archive/tar for ASP usage.
type G3TAR struct {
	vm        *VM
	lastError string
	file      *os.File
	writer    *tar.Writer
	path      string
	mode      string
	entries   []g3tarEntry
}

// newG3TARObject creates one native G3TAR object.
func (vm *VM) newG3TARObject() Value {
	obj := &G3TAR{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3tarItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet routes readable G3TAR properties.
func (t *G3TAR) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(strings.TrimSpace(propertyName)) {
	case "path":
		return NewString(t.path)
	case "mode":
		return NewString(t.mode)
	case "count":
		return NewInteger(int64(len(t.entries)))
	case "lasterror":
		return NewString(t.lastError)
	}
	return t.DispatchMethod(propertyName, nil)
}

// DispatchMethod routes all G3TAR methods.
func (t *G3TAR) DispatchMethod(methodName string, args []Value) Value {
	switch strings.ToLower(strings.TrimSpace(methodName)) {
	case "create":
		if len(args) < 1 {
			return NewBool(false)
		}
		return NewBool(t.Create(args[0].String()))
	case "open":
		if len(args) < 1 {
			return NewBool(false)
		}
		return NewBool(t.Open(args[0].String()))
	case "addfile":
		if len(args) < 1 {
			return NewBool(false)
		}
		nameInTar := ""
		if len(args) >= 2 {
			nameInTar = args[1].String()
		}
		return NewBool(t.AddFile(args[0].String(), nameInTar))
	case "addfolder":
		if len(args) < 1 {
			return NewBool(false)
		}
		nameInTar := ""
		if len(args) >= 2 {
			nameInTar = args[1].String()
		}
		return NewBool(t.AddFolder(args[0].String(), nameInTar))
	case "addfiles":
		if len(args) < 1 {
			return NewBool(false)
		}
		prefix := ""
		if len(args) >= 2 {
			prefix = args[1].String()
		}
		return NewBool(t.AddFiles(args[0], prefix))
	case "addtext":
		if len(args) < 2 {
			return NewBool(false)
		}
		return NewBool(t.AddText(args[0].String(), args[1].String()))
	case "list":
		return t.List()
	case "extractall", "extract":
		if len(args) < 1 {
			return NewBool(false)
		}
		return NewBool(t.ExtractAll(args[0].String()))
	case "extractfile", "extractsingle":
		if len(args) < 2 {
			return NewBool(false)
		}
		return NewBool(t.ExtractFile(args[0].String(), args[1].String()))
	case "getinfo", "getfileinfo":
		if len(args) < 1 {
			return NewEmpty()
		}
		return t.GetFileInfo(args[0].String())
	case "close", "dispose":
		t.Close()
		return NewBool(true)
	}
	return NewEmpty()
}

// Create initializes TAR write mode.
func (t *G3TAR) Create(relPath string) bool {
	t.Close()
	targetPath, ok := t.vm.fsoResolvePath(relPath)
	if !ok {
		t.raiseError("create tar failed", fmt.Errorf("target path is outside the web root sandbox"))
		return false
	}
	if err := os.MkdirAll(filepath.Dir(targetPath), 0755); err != nil {
		t.raiseError("create tar mkdir failed", err)
		return false
	}
	f, err := os.Create(targetPath)
	if err != nil {
		t.raiseError("create tar file failed", err)
		return false
	}
	t.file = f
	t.writer = tar.NewWriter(f)
	t.path = targetPath
	t.mode = "w"
	t.entries = t.entries[:0]
	t.lastError = ""
	return true
}

// Open initializes TAR read mode and indexes headers for fast list/info operations.
func (t *G3TAR) Open(relPath string) bool {
	t.Close()
	sourcePath, ok := t.vm.fsoResolvePath(relPath)
	if !ok {
		t.raiseError("open tar failed", fmt.Errorf("source path is outside the web root sandbox"))
		return false
	}
	f, err := os.Open(sourcePath)
	if err != nil {
		t.raiseError("open tar file failed", err)
		return false
	}

	r := tar.NewReader(f)
	entries := make([]g3tarEntry, 0, 64)
	for {
		hdr, err := r.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			f.Close()
			t.raiseError("open tar index failed", err)
			return false
		}
		entries = append(entries, g3tarEntry{Name: hdr.Name, Size: hdr.Size, Mode: hdr.Mode, ModTime: hdr.ModTime, TypeFlag: hdr.Typeflag})
	}

	if _, err := f.Seek(0, io.SeekStart); err != nil {
		f.Close()
		t.raiseError("open tar rewind failed", err)
		return false
	}
	t.file = f
	t.path = sourcePath
	t.mode = "r"
	t.entries = entries
	t.lastError = ""
	return true
}

// AddFile streams one file into the current tar archive.
func (t *G3TAR) AddFile(sourceRelPath string, nameInTar string) bool {
	if t.mode != "w" || t.writer == nil {
		return false
	}
	sourcePath, ok := t.vm.fsoResolvePath(sourceRelPath)
	if !ok {
		t.raiseError("add file failed", fmt.Errorf("source path is outside the web root sandbox"))
		return false
	}
	f, err := os.Open(sourcePath)
	if err != nil {
		t.raiseError("add file open failed", err)
		return false
	}
	defer f.Close()

	stat, err := f.Stat()
	if err != nil {
		t.raiseError("add file stat failed", err)
		return false
	}
	if stat.IsDir() {
		t.raiseError("add file failed", fmt.Errorf("source path points to directory"))
		return false
	}

	entryName := nameInTar
	if strings.TrimSpace(entryName) == "" {
		entryName = filepath.Base(sourcePath)
	}
	entryName = filepath.ToSlash(strings.TrimLeft(entryName, "\\/"))
	if entryName == "" {
		t.raiseError("add file failed", fmt.Errorf("entry name is empty"))
		return false
	}

	hdr, err := tar.FileInfoHeader(stat, "")
	if err != nil {
		t.raiseError("add file header failed", err)
		return false
	}
	hdr.Name = entryName
	if err := t.writer.WriteHeader(hdr); err != nil {
		t.raiseError("add file write header failed", err)
		return false
	}
	if _, err := io.Copy(t.writer, f); err != nil {
		t.raiseError("add file write body failed", err)
		return false
	}
	t.entries = append(t.entries, g3tarEntry{Name: hdr.Name, Size: hdr.Size, Mode: hdr.Mode, ModTime: hdr.ModTime, TypeFlag: hdr.Typeflag})
	t.lastError = ""
	return true
}

// AddFolder recursively writes one folder into the current tar archive.
func (t *G3TAR) AddFolder(sourceRelPath string, nameInTar string) bool {
	if t.mode != "w" || t.writer == nil {
		return false
	}
	sourcePath, ok := t.vm.fsoResolvePath(sourceRelPath)
	if !ok {
		t.raiseError("add folder failed", fmt.Errorf("source path is outside the web root sandbox"))
		return false
	}
	baseName := nameInTar
	if strings.TrimSpace(baseName) == "" {
		baseName = filepath.Base(sourcePath)
	}
	baseName = filepath.ToSlash(strings.TrimLeft(baseName, "\\/"))

	err := filepath.Walk(sourcePath, func(path string, info os.FileInfo, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		rel, err := filepath.Rel(sourcePath, path)
		if err != nil {
			return err
		}

		entryName := baseName
		if rel != "." {
			entryName = filepath.ToSlash(filepath.Join(baseName, rel))
		}
		entryName = strings.TrimLeft(entryName, "\\/")
		hdr, err := tar.FileInfoHeader(info, "")
		if err != nil {
			return err
		}
		if info.IsDir() {
			if !strings.HasSuffix(entryName, "/") {
				entryName += "/"
			}
		}
		hdr.Name = entryName
		if err := t.writer.WriteHeader(hdr); err != nil {
			return err
		}
		if info.IsDir() {
			t.entries = append(t.entries, g3tarEntry{Name: hdr.Name, Size: hdr.Size, Mode: hdr.Mode, ModTime: hdr.ModTime, TypeFlag: hdr.Typeflag})
			return nil
		}
		f, err := os.Open(path)
		if err != nil {
			return err
		}
		defer f.Close()
		if _, err := io.Copy(t.writer, f); err != nil {
			return err
		}
		t.entries = append(t.entries, g3tarEntry{Name: hdr.Name, Size: hdr.Size, Mode: hdr.Mode, ModTime: hdr.ModTime, TypeFlag: hdr.Typeflag})
		return nil
	})
	if err != nil {
		t.raiseError("add folder failed", err)
		return false
	}
	t.lastError = ""
	return true
}

// AddFiles adds many files from one ASP array using one optional target prefix.
func (t *G3TAR) AddFiles(paths Value, prefix string) bool {
	items := g3zlibNormalizeBatchInput(paths)
	prefix = filepath.ToSlash(strings.TrimLeft(strings.TrimSpace(prefix), "\\/"))
	for i := 0; i < len(items); i++ {
		source := items[i].String()
		base := filepath.Base(source)
		nameInTar := base
		if prefix != "" {
			nameInTar = filepath.ToSlash(filepath.Join(prefix, base))
		}
		if !t.AddFile(source, nameInTar) {
			return false
		}
	}
	return true
}

// AddText writes one in-memory text entry into the current tar archive.
func (t *G3TAR) AddText(nameInTar string, content string) bool {
	if t.mode != "w" || t.writer == nil {
		return false
	}
	entryName := filepath.ToSlash(strings.TrimLeft(strings.TrimSpace(nameInTar), "\\/"))
	if entryName == "" {
		t.raiseError("add text failed", fmt.Errorf("entry name is empty"))
		return false
	}
	payload := []byte(content)
	hdr := &tar.Header{
		Name:    entryName,
		Mode:    0644,
		Size:    int64(len(payload)),
		ModTime: time.Now().UTC(),
	}
	if err := t.writer.WriteHeader(hdr); err != nil {
		t.raiseError("add text write header failed", err)
		return false
	}
	if len(payload) > 0 {
		if _, err := io.Copy(t.writer, bytes.NewReader(payload)); err != nil {
			t.raiseError("add text write body failed", err)
			return false
		}
	}
	t.entries = append(t.entries, g3tarEntry{Name: hdr.Name, Size: hdr.Size, Mode: hdr.Mode, ModTime: hdr.ModTime, TypeFlag: hdr.Typeflag})
	t.lastError = ""
	return true
}

// List returns one ASP array with all entry names.
func (t *G3TAR) List() Value {
	if len(t.entries) == 0 {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}
	arr := make([]Value, len(t.entries))
	for i := 0; i < len(t.entries); i++ {
		arr[i] = NewString(t.entries[i].Name)
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, arr)}
}

// GetFileInfo returns one dictionary with entry metadata.
func (t *G3TAR) GetFileInfo(name string) Value {
	for i := 0; i < len(t.entries); i++ {
		if strings.EqualFold(t.entries[i].Name, name) {
			entry := t.entries[i]
			dictVal := t.vm.newDictionaryObject()
			t.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Name"), NewString(entry.Name)})
			t.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Size"), NewInteger(entry.Size)})
			t.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Mode"), NewString(fmt.Sprintf("0o%o", entry.Mode))})
			t.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Modified"), NewString(entry.ModTime.Format(time.RFC3339))})
			t.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("IsDir"), NewBool(entry.TypeFlag == tar.TypeDir || strings.HasSuffix(entry.Name, "/"))})
			return dictVal
		}
	}
	return NewEmpty()
}

// ExtractAll extracts all archive entries to one target directory.
func (t *G3TAR) ExtractAll(targetRelPath string) bool {
	if t.mode != "r" || t.file == nil {
		return false
	}
	targetPath, ok := t.vm.fsoResolvePath(targetRelPath)
	if !ok {
		t.raiseError("extract all failed", fmt.Errorf("target path is outside the web root sandbox"))
		return false
	}
	if _, err := t.file.Seek(0, io.SeekStart); err != nil {
		t.raiseError("extract all rewind failed", err)
		return false
	}
	r := tar.NewReader(t.file)
	for {
		hdr, err := r.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.raiseError("extract all read header failed", err)
			return false
		}
		if err := t.extractHeader(r, hdr, targetPath); err != nil {
			t.raiseError("extract all write failed", err)
			return false
		}
	}
	t.lastError = ""
	return true
}

// ExtractFile extracts one specific entry by name.
func (t *G3TAR) ExtractFile(entryName string, targetRelPath string) bool {
	if t.mode != "r" || t.file == nil {
		return false
	}
	targetPath, ok := t.vm.fsoResolvePath(targetRelPath)
	if !ok {
		t.raiseError("extract file failed", fmt.Errorf("target path is outside the web root sandbox"))
		return false
	}
	if _, err := t.file.Seek(0, io.SeekStart); err != nil {
		t.raiseError("extract file rewind failed", err)
		return false
	}
	r := tar.NewReader(t.file)
	for {
		hdr, err := r.Next()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.raiseError("extract file read header failed", err)
			return false
		}
		if strings.EqualFold(hdr.Name, entryName) {
			if err := t.extractHeader(r, hdr, targetPath); err != nil {
				t.raiseError("extract file write failed", err)
				return false
			}
			t.lastError = ""
			return true
		}
	}
	t.raiseError("extract file failed", fmt.Errorf("entry not found: %s", entryName))
	return false
}

// extractHeader writes one tar header payload to disk with path traversal protection.
func (t *G3TAR) extractHeader(reader io.Reader, hdr *tar.Header, targetRoot string) error {
	cleanRoot := filepath.Clean(targetRoot)
	cleanName := filepath.Clean(filepath.FromSlash(hdr.Name))
	destination := filepath.Join(cleanRoot, cleanName)
	if !strings.HasPrefix(strings.ToLower(destination), strings.ToLower(cleanRoot)+string(os.PathSeparator)) && !strings.EqualFold(destination, cleanRoot) {
		return fmt.Errorf("path traversal attempt detected: %s", hdr.Name)
	}

	switch hdr.Typeflag {
	case tar.TypeDir:
		return os.MkdirAll(destination, os.FileMode(hdr.Mode))
	case tar.TypeReg, tar.TypeRegA:
		if err := os.MkdirAll(filepath.Dir(destination), 0755); err != nil {
			return err
		}
		f, err := os.OpenFile(destination, os.O_CREATE|os.O_WRONLY|os.O_TRUNC, os.FileMode(hdr.Mode))
		if err != nil {
			return err
		}
		defer f.Close()
		_, err = io.Copy(f, reader)
		return err
	default:
		return fmt.Errorf("unsupported tar entry type: %d", hdr.Typeflag)
	}
}

// Close releases all open resources.
func (t *G3TAR) Close() {
	if t.writer != nil {
		t.writer.Close()
		t.writer = nil
	}
	if t.file != nil {
		t.file.Close()
		t.file = nil
	}
	t.path = ""
	t.mode = ""
	t.entries = t.entries[:0]
}

// raiseError stores one error string and raises one VM runtime error.
func (t *G3TAR) raiseError(context string, err error) {
	if err == nil {
		return
	}
	message := fmt.Sprintf("%s: %v", context, err)
	t.lastError = message
	t.vm.raise(vbscript.InternalError, message)
}
