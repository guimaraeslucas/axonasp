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
	"archive/zip"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"
)

// G3ZIP implements Component interface for ZIP operations
type G3ZIP struct {
	ctx     *ExecutionContext
	zipFile *os.File
	writer  *zip.Writer
	reader  *zip.ReadCloser
	path    string
	mode    string // "r" for read, "w" for write
}

func (z *G3ZIP) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "path":
		return z.path
	case "mode":
		return z.mode
	case "count":
		if z.reader != nil {
			return len(z.reader.File)
		}
		return 0
	}
	return nil
}

func (z *G3ZIP) SetProperty(name string, value interface{}) {}

func (z *G3ZIP) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	// Handle explicit CallMethod call (e.g. obj.CallMethod("MethodName", args))
	if method == "callmethod" && len(args) > 0 {
		actualMethod := fmt.Sprintf("%v", args[0])
		return z.CallMethod(actualMethod, args[1:]...)
	}

	switch method {
	case "open":
		if len(args) < 1 {
			return false
		}
		return z.Open(fmt.Sprintf("%v", args[0]))

	case "create":
		if len(args) < 1 {
			return false
		}
		return z.Create(fmt.Sprintf("%v", args[0]))

	case "addfile":
		if len(args) < 1 {
			return false
		}
		source := fmt.Sprintf("%v", args[0])
		nameInZip := ""
		if len(args) >= 2 {
			nameInZip = fmt.Sprintf("%v", args[1])
		}
		return z.AddFile(source, nameInZip)

	case "addfolder":
		if len(args) < 1 {
			return false
		}
		source := fmt.Sprintf("%v", args[0])
		nameInZip := ""
		if len(args) >= 2 {
			nameInZip = fmt.Sprintf("%v", args[1])
		}
		return z.AddFolder(source, nameInZip)

	case "addtext":
		if len(args) < 2 {
			return false
		}
		nameInZip := fmt.Sprintf("%v", args[0])
		content := fmt.Sprintf("%v", args[1])
		return z.AddText(nameInZip, content)

	case "extractall", "extract":
		if len(args) < 1 {
			return false
		}
		return z.ExtractAll(fmt.Sprintf("%v", args[0]))

	case "extractfile":
		if len(args) < 2 {
			return false
		}
		fileName := fmt.Sprintf("%v", args[0])
		targetPath := fmt.Sprintf("%v", args[1])
		return z.ExtractFile(fileName, targetPath)

	case "list":
		return z.List()

	case "getinfo":
		if len(args) < 1 {
			return nil
		}
		return z.GetFileInfo(fmt.Sprintf("%v", args[0]))

	case "close":
		z.Close()
		return true
	}

	return nil
}

// Open opens a zip file for reading
func (z *G3ZIP) Open(relPath string) bool {
	z.Close()
	fullPath := z.ctx.Server_MapPath(relPath)
	if !z.validatePath(fullPath) {
		return false
	}

	r, err := zip.OpenReader(fullPath)
	if err != nil {
		log.Printf("G3ZIP Error opening %s: %v\n", fullPath, err)
		return false
	}

	z.reader = r
	z.path = fullPath
	z.mode = "r"
	return true
}

// Create creates a new zip file for writing
func (z *G3ZIP) Create(relPath string) bool {
	z.Close()
	fullPath := z.ctx.Server_MapPath(relPath)
	if !z.validatePath(fullPath) {
		return false
	}

	// Ensure directory exists
	dir := filepath.Dir(fullPath)
	os.MkdirAll(dir, 0755)

	f, err := os.Create(fullPath)
	if err != nil {
		log.Printf("G3ZIP Error creating %s: %v\n", fullPath, err)
		return false
	}

	z.zipFile = f
	z.writer = zip.NewWriter(f)
	z.path = fullPath
	z.mode = "w"
	return true
}

// AddFile adds a file from disk to the zip
func (z *G3ZIP) AddFile(sourceRelPath, nameInZip string) bool {
	if z.mode != "w" || z.writer == nil {
		return false
	}

	sourceFullPath := z.ctx.Server_MapPath(sourceRelPath)
	if !z.validatePath(sourceFullPath) {
		return false
	}

	fileToZip, err := os.Open(sourceFullPath)
	if err != nil {
		return false
	}
	defer fileToZip.Close()

	info, err := fileToZip.Stat()
	if err != nil {
		return false
	}

	if nameInZip == "" {
		nameInZip = filepath.Base(sourceFullPath)
	}

	header, err := zip.FileInfoHeader(info)
	if err != nil {
		return false
	}

	header.Name = nameInZip
	header.Method = zip.Deflate

	writer, err := z.writer.CreateHeader(header)
	if err != nil {
		return false
	}

	_, err = io.Copy(writer, fileToZip)
	return err == nil
}

// AddText adds a string as a file in the zip
func (z *G3ZIP) AddText(nameInZip, content string) bool {
	if z.mode != "w" || z.writer == nil {
		return false
	}

	writer, err := z.writer.Create(nameInZip)
	if err != nil {
		return false
	}

	_, err = io.WriteString(writer, content)
	return err == nil
}

// AddFolder adds a folder recursively to the zip
func (z *G3ZIP) AddFolder(sourceRelPath, nameInZip string) bool {
	if z.mode != "w" || z.writer == nil {
		return false
	}

	sourceFullPath := z.ctx.Server_MapPath(sourceRelPath)
	if !z.validatePath(sourceFullPath) {
		return false
	}

	if nameInZip == "" {
		nameInZip = filepath.Base(sourceFullPath)
	}

	err := filepath.Walk(sourceFullPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

		// Create relative path within zip
		rel, err := filepath.Rel(sourceFullPath, path)
		if err != nil {
			return err
		}

		zipPath := filepath.ToSlash(filepath.Join(nameInZip, rel))

		if info.IsDir() {
			if zipPath != "" {
				_, err = z.writer.Create(zipPath + "/")
			}
			return err
		}

		header, err := zip.FileInfoHeader(info)
		if err != nil {
			return err
		}
		header.Name = zipPath
		header.Method = zip.Deflate

		writer, err := z.writer.CreateHeader(header)
		if err != nil {
			return err
		}

		file, err := os.Open(path)
		if err != nil {
			return err
		}
		defer file.Close()

		_, err = io.Copy(writer, file)
		return err
	})

	return err == nil
}

// ExtractAll extracts all files to a directory
func (z *G3ZIP) ExtractAll(targetRelPath string) bool {
	if z.mode != "r" || z.reader == nil {
		return false
	}

	targetFullPath := z.ctx.Server_MapPath(targetRelPath)
	if !z.validatePath(targetFullPath) {
		return false
	}

	for _, f := range z.reader.File {
		if !z.extractFileTo(f, targetFullPath) {
			return false
		}
	}

	return true
}

// ExtractFile extracts a single file from the zip
func (z *G3ZIP) ExtractFile(fileName, targetRelPath string) bool {
	if z.mode != "r" || z.reader == nil {
		return false
	}

	targetFullPath := z.ctx.Server_MapPath(targetRelPath)
	if !z.validatePath(targetFullPath) {
		return false
	}

	for _, f := range z.reader.File {
		if f.Name == fileName {
			return z.extractFileTo(f, targetFullPath)
		}
	}

	return false
}

func (z *G3ZIP) extractFileTo(f *zip.File, targetDir string) bool {
	path := filepath.Join(targetDir, f.Name)

	// Check for ZipSlip vulnerability
	if !strings.HasPrefix(path, filepath.Clean(targetDir)+string(os.PathSeparator)) && path != filepath.Clean(targetDir) {
		return false
	}

	if f.FileInfo().IsDir() {
		os.MkdirAll(path, os.ModePerm)
		return true
	}

	if err := os.MkdirAll(filepath.Dir(path), os.ModePerm); err != nil {
		return false
	}

	dstFile, err := os.OpenFile(path, os.O_WRONLY|os.O_CREATE|os.O_TRUNC, f.Mode())
	if err != nil {
		return false
	}
	defer dstFile.Close()

	fileInZip, err := f.Open()
	if err != nil {
		return false
	}
	defer fileInZip.Close()

	_, err = io.Copy(dstFile, fileInZip)
	return err == nil
}

// List returns a slice of file names
func (z *G3ZIP) List() []interface{} {
	if z.mode != "r" || z.reader == nil {
		return []interface{}{}
	}

	var list []interface{}
	for _, f := range z.reader.File {
		list = append(list, f.Name)
	}
	return list
}

// GetFileInfo returns info about a file in the zip
func (z *G3ZIP) GetFileInfo(fileName string) interface{} {
	if z.mode != "r" || z.reader == nil {
		return nil
	}

	for _, f := range z.reader.File {
		if f.Name == fileName {
			dict := NewDictionary(z.ctx)
			dict.CallMethod("Add", "Name", f.Name)
			dict.CallMethod("Size", f.UncompressedSize64)
			dict.CallMethod("PackedSize", f.CompressedSize64)
			dict.CallMethod("Modified", f.Modified.Format(time.RFC3339))
			dict.CallMethod("IsDir", f.FileInfo().IsDir())
			return dict
		}
	}

	return nil
}

// Close closes the zip file
func (z *G3ZIP) Close() {
	if z.writer != nil {
		z.writer.Close()
		z.writer = nil
	}
	if z.zipFile != nil {
		z.zipFile.Close()
		z.zipFile = nil
	}
	if z.reader != nil {
		z.reader.Close()
		z.reader = nil
	}
	z.path = ""
	z.mode = ""
}

func (z *G3ZIP) validatePath(fullPath string) bool {
	if fullPath == "" {
		return false
	}

	rootDir, _ := filepath.Abs(z.ctx.RootDir)
	absPath, _ := filepath.Abs(fullPath)

	if !strings.HasPrefix(strings.ToLower(absPath), strings.ToLower(rootDir)) {
		log.Printf("Security Warning: G3ZIP tried to access %s (Root: %s)\n", absPath, rootDir)
		return false
	}

	return true
}
