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
	"archive/zip"
	"io"
	"os"
	"path/filepath"
	"strings"
	"time"
)

type G3Zip struct {
	vm      *VM
	zipFile *os.File
	writer  *zip.Writer
	reader  *zip.ReadCloser
	path    string
	mode    string // "r" for read, "w" for write
}

// newG3ZipObject instantiates the G3ZIP custom functions library.
func (vm *VM) newG3ZipObject() Value {
	obj := &G3Zip{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.g3zipItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter.
func (z *G3Zip) DispatchPropertyGet(propertyName string) Value {
	switch strings.ToLower(propertyName) {
	case "path":
		return NewString(z.path)
	case "mode":
		return NewString(z.mode)
	case "count":
		if z.reader != nil {
			return NewInteger(int64(len(z.reader.File)))
		}
		return NewInteger(0)
	}
	return z.DispatchMethod(propertyName, nil)
}

// DispatchMethod provides O(1) string matching resolution.
func (z *G3Zip) DispatchMethod(methodName string, args []Value) Value {
	method := strings.ToLower(methodName)

	switch method {
	case "open":
		if len(args) < 1 {
			return NewBool(false)
		}
		return NewBool(z.Open(args[0].String()))

	case "create":
		if len(args) < 1 {
			return NewBool(false)
		}
		return NewBool(z.Create(args[0].String()))

	case "addfile":
		if len(args) < 1 {
			return NewBool(false)
		}
		source := args[0].String()
		nameInZip := ""
		if len(args) >= 2 {
			nameInZip = args[1].String()
		}
		return NewBool(z.AddFile(source, nameInZip))

	case "addfolder":
		if len(args) < 1 {
			return NewBool(false)
		}
		source := args[0].String()
		nameInZip := ""
		if len(args) >= 2 {
			nameInZip = args[1].String()
		}
		return NewBool(z.AddFolder(source, nameInZip))

	case "addtext":
		if len(args) < 2 {
			return NewBool(false)
		}
		return NewBool(z.AddText(args[0].String(), args[1].String()))

	case "extractall", "extract":
		if len(args) < 1 {
			return NewBool(false)
		}
		return NewBool(z.ExtractAll(args[0].String()))

	case "extractfile":
		if len(args) < 2 {
			return NewBool(false)
		}
		return NewBool(z.ExtractFile(args[0].String(), args[1].String()))

	case "list":
		return z.List()

	case "getinfo", "getfileinfo":
		if len(args) < 1 {
			return NewEmpty()
		}
		return z.GetFileInfo(args[0].String())

	case "close":
		z.Close()
		return NewBool(true)
	}

	return NewEmpty()
}

func (z *G3Zip) getFullPath(relPath string) string {
	trimmed := strings.TrimSpace(relPath)
	if trimmed == "" {
		return ""
	}

	if filepath.IsAbs(trimmed) || strings.HasPrefix(trimmed, "\\\\") || filepath.VolumeName(trimmed) != "" {
		return filepath.Clean(trimmed)
	}

	if z.vm.host != nil && z.vm.host.Server() != nil {
		return z.vm.host.Server().MapPath(trimmed)
	}
	abs, _ := filepath.Abs(trimmed)
	return abs
}

func (z *G3Zip) Open(relPath string) bool {
	z.Close()
	fullPath := z.getFullPath(relPath)

	r, err := zip.OpenReader(fullPath)
	if err != nil {
		return false
	}

	z.reader = r
	z.path = fullPath
	z.mode = "r"
	return true
}

func (z *G3Zip) Create(relPath string) bool {
	z.Close()
	fullPath := z.getFullPath(relPath)

	dir := filepath.Dir(fullPath)
	os.MkdirAll(dir, 0755)

	f, err := os.Create(fullPath)
	if err != nil {
		return false
	}

	z.zipFile = f
	z.writer = zip.NewWriter(f)
	z.path = fullPath
	z.mode = "w"
	return true
}

func (z *G3Zip) AddFile(sourceRelPath, nameInZip string) bool {
	if z.mode != "w" || z.writer == nil {
		return false
	}

	sourceFullPath := z.getFullPath(sourceRelPath)
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
	nameInZip = filepath.ToSlash(strings.TrimLeft(nameInZip, "\\/"))
	if nameInZip == "" {
		return false
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

func (z *G3Zip) AddText(nameInZip, content string) bool {
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

func (z *G3Zip) AddFolder(sourceRelPath, nameInZip string) bool {
	if z.mode != "w" || z.writer == nil {
		return false
	}

	sourceFullPath := z.getFullPath(sourceRelPath)
	if nameInZip == "" {
		nameInZip = filepath.Base(sourceFullPath)
	}

	err := filepath.Walk(sourceFullPath, func(path string, info os.FileInfo, err error) error {
		if err != nil {
			return err
		}

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

func (z *G3Zip) ExtractAll(targetRelPath string) bool {
	if z.mode != "r" || z.reader == nil {
		return false
	}

	targetFullPath := z.getFullPath(targetRelPath)
	for _, f := range z.reader.File {
		if !z.extractFileTo(f, targetFullPath) {
			return false
		}
	}
	return true
}

func (z *G3Zip) ExtractFile(fileName, targetRelPath string) bool {
	if z.mode != "r" || z.reader == nil {
		return false
	}

	targetFullPath := z.getFullPath(targetRelPath)
	for _, f := range z.reader.File {
		if f.Name == fileName {
			return z.extractFileTo(f, targetFullPath)
		}
	}
	return false
}

func (z *G3Zip) extractFileTo(f *zip.File, targetDir string) bool {
	path := filepath.Join(targetDir, f.Name)

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

func (z *G3Zip) List() Value {
	if z.mode != "r" || z.reader == nil {
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
	}

	arr := make([]Value, len(z.reader.File))
	for i, f := range z.reader.File {
		arr[i] = NewString(f.Name)
	}
	return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, arr)}
}

func (z *G3Zip) GetFileInfo(fileName string) Value {
	if z.mode != "r" || z.reader == nil {
		return NewEmpty()
	}

	for _, f := range z.reader.File {
		if strings.EqualFold(f.Name, fileName) {
			dictVal := z.vm.newDictionaryObject()
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Name"), NewString(f.Name)})
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Size"), NewInteger(int64(f.UncompressedSize64))})
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("PackedSize"), NewInteger(int64(f.CompressedSize64))})
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("Modified"), NewString(f.Modified.Format(time.RFC3339))})
			z.vm.dispatchDictionaryMethod(dictVal.Num, "Add", []Value{NewString("IsDir"), NewBool(f.FileInfo().IsDir())})
			return dictVal
		}
	}
	return NewEmpty()
}

func (z *G3Zip) Close() {
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
