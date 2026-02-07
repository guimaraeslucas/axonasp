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
	"bufio"
	"bytes"
	"encoding/binary"
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"
	"unicode/utf16"
)

// G3FILES implements Component interface for File operations
type G3FILES struct {
	ctx *ExecutionContext
}

func (f *G3FILES) GetProperty(name string) interface{} {
	return nil
}

func (f *G3FILES) SetProperty(name string, value interface{}) {}

func (f *G3FILES) CallMethod(name string, args ...interface{}) interface{} {
	if len(args) < 1 || args[0] == nil {
		log.Println("Error: G3FILES method requires a valid path argument")
		return nil
	}

	getStr := func(i int) string {
		if i >= len(args) || args[i] == nil {
			return ""
		}
		return fmt.Sprintf("%v", args[i])
	}

	relPath := getStr(0)

	// Validate path is not empty or nil
	if relPath == "" || relPath == "<nil>" {
		log.Println("Error: G3FILES received empty or nil path")
		return nil
	}

	fullPath := f.ctx.Server_MapPath(relPath)

	// Validate mapped path
	if fullPath == "" || fullPath == "<nil>" {
		log.Printf("Error: Server_MapPath returned invalid path for %s\n", relPath)
		return nil
	}

	rootDir, _ := filepath.Abs(f.ctx.RootDir)
	absPath, _ := filepath.Abs(fullPath)

	if !strings.HasPrefix(strings.ToLower(absPath), strings.ToLower(rootDir)) {
		log.Printf("Security Warning: Script tried to access %s (Root: %s)\n", absPath, rootDir)
		return nil
	}

	method := strings.ToLower(name)

	switch method {
	case "exists":
		_, err := os.Stat(fullPath)
		return err == nil || !os.IsNotExist(err)

	case "read", "readtext":
		content, err := os.ReadFile(fullPath)
		if err != nil {
			return ""
		}
		return string(decodeTextWithCharset(content, "utf-8"))

	case "write", "writetext":
		if len(args) < 2 {
			return false
		}
		content := getStr(1)
		err := os.WriteFile(fullPath, []byte(content), 0644)
		return err == nil

	case "append", "appendtext":
		if len(args) < 2 {
			return false
		}
		content := getStr(1)

		f, err := os.OpenFile(fullPath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
		if err != nil {
			return false
		}
		defer f.Close()

		_, err = f.WriteString(content)
		return err == nil

	case "delete", "remove":
		err := os.Remove(fullPath)
		return err == nil

	case "copy":
		if len(args) < 2 {
			return false
		}
		// fullPath is source (args[0])
		// dest needs resolution
		destRel := getStr(1)
		destPath := f.ctx.Server_MapPath(destRel)

		data, err := os.ReadFile(fullPath)
		if err != nil {
			return false
		}
		err = os.WriteFile(destPath, data, 0644)
		return err == nil

	case "size":
		info, err := os.Stat(fullPath)
		if err != nil {
			return 0
		}
		return int(info.Size())

	case "mkdir", "makedir":
		err := os.MkdirAll(fullPath, 0755)
		return err == nil

	case "list", "listfiles":
		names := make([]interface{}, 0)
		files, err := os.ReadDir(fullPath)
		if err != nil {
			fmt.Printf("AxonASP File Error: %v\n", err)
			return names
		}
		for _, f := range files {
			if !f.IsDir() {
				names = append(names, f.Name())
			}
		}
		return names
	}

	return nil
}

// decodeTextWithCharset decodes bytes to string based on charset
func decodeTextWithCharset(data []byte, charset string) []byte {
	// Handle BOM
	if len(data) >= 3 && bytes.Equal(data[:3], []byte{0xEF, 0xBB, 0xBF}) {
		return data[3:]
	}

	if len(data) >= 2 {
		switch {
		case data[0] == 0xFF && data[1] == 0xFE:
			return []byte(decodeUTF16(data[2:], binary.LittleEndian))
		case data[0] == 0xFE && data[1] == 0xFF:
			return []byte(decodeUTF16(data[2:], binary.BigEndian))
		}
	}

	cs := strings.ToLower(charset)
	switch cs {
	case "unicode", "utf-16", "utf-16le":
		return []byte(decodeUTF16(data, binary.LittleEndian))
	case "utf-16be":
		return []byte(decodeUTF16(data, binary.BigEndian))
	case "iso-8859-1", "windows-1252", "ascii", "us-ascii":
		// Binary-safe string conversion (1 byte = 1 char)
		// This is critical for binary file uploads where ADODB.Stream is used
		// to convert binary data to a string for parsing boundaries.
		// Go's string([]byte) assumes UTF-8 and replaces invalid bytes with \uFFFD.
		// We must map bytes 0-255 directly to runes 0-255.
		runes := make([]rune, len(data))
		for i, b := range data {
			runes[i] = rune(b)
		}
		// Convert runes to string (UTF-8 encoded internally in Go)
		// But logically preserving the 1-to-1 mapping
		return []byte(string(runes))
	default:
		return data
	}
}

// decodeSingleByteString converts bytes to string treating each byte as a rune (ISO-8859-1)
func decodeSingleByteString(data []byte) string {
	runes := make([]rune, len(data))
	for i, b := range data {
		runes[i] = rune(b)
	}
	return string(runes)
}

// encodeSingleByteString converts string to bytes treating each char as a byte (ISO-8859-1)
func encodeSingleByteString(s string) []byte {
	data := make([]byte, 0, len(s))
	for _, r := range s {
		if r <= 0xFF {
			data = append(data, byte(r))
		} else {
			// Replace non-ISO-8859-1 chars with '?'
			data = append(data, '?')
		}
	}
	return data
}

func decodeUTF16(data []byte, order binary.ByteOrder) string {
	if len(data) == 0 {
		return ""
	}
	if len(data)%2 != 0 {
		data = data[:len(data)-1]
	}

	u16 := make([]uint16, len(data)/2)
	for i := range u16 {
		u16[i] = order.Uint16(data[i*2:])
	}
	return string(utf16.Decode(u16))
}

func FileSystemAPI(method string, args []string, ctx *ExecutionContext) interface{} {
	lib := &G3FILES{ctx: ctx}
	var ifaceArgs []interface{}
	for _, a := range args {
		val := EvaluateExpression(a, ctx)
		ifaceArgs = append(ifaceArgs, val)
	}
	return lib.CallMethod(method, ifaceArgs)
}

// --- FSO Implementation (Scripting.FileSystemObject) ---

type FSOObject struct {
	ctx *ExecutionContext
}

func (f *FSOObject) GetProperty(name string) interface{} {
	if strings.EqualFold(name, "Drives") {
		return &FSODrives{ctx: f.ctx}
	}
	return nil
}

func (f *FSOObject) SetProperty(name string, value interface{}) {}

func (f *FSOObject) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	getStr := func(i int) string {
		if i >= len(args) || args[i] == nil {
			return ""
		}
		return fmt.Sprintf("%v", args[i])
	}

	// Helper to resolve path (simple MapPath wrapper)
	resolve := func(p string) string {
		// Validate path is not empty or nil
		if p == "" || p == "<nil>" {
			// Return empty silently - this is normal when path property is uninitialized
			return ""
		}
		// If p is absolute (e.g. C:\), use it?
		// VBScript FSO usually takes absolute paths or relative.
		// If it has drive letter, it's absolute.
		// But here we are sandboxed usually.
		// However, user asked for Drives support.
		// Let's try to be smart.
		// If starts with X:, assume absolute.
		if len(p) > 1 && p[1] == ':' {
			return p
		}
		mapped := f.ctx.Server_MapPath(p)
		// Validate mapped result
		if mapped == "" || mapped == "<nil>" {
			// Return empty silently - path resolution failed
			return ""
		}
		return mapped
	}

	switch method {
	case "buildpath":
		if len(args) < 2 {
			return ""
		}
		return filepath.Join(getStr(0), getStr(1))
	case "copyfile":
		// CopyFile source, dest, [overwrite]
		if len(args) < 2 {
			return nil
		}
		src := resolve(getStr(0))
		dst := resolve(getStr(1))
		// naive copy
		data, err := os.ReadFile(src)
		if err != nil {
			return nil
		}
		os.WriteFile(dst, data, 0644)
		return nil
	case "copyfolder":
		//TODO - CREATE COPYFOLDER PRIORITY
		return nil
	case "createfolder":
		path := resolve(getStr(0))
		os.MkdirAll(path, 0755)
		return &FSOFolder{ctx: f.ctx, path: path}
	case "createtextfile":
		// CreateTextFile(filename, [overwrite], [unicode])
		path := resolve(getStr(0))
		fp, err := os.Create(path)
		if err != nil {
			return nil
		}
		return &TextStream{f: fp, mode: 2} // 2=ForWriting
	case "deletefile":
		path := resolve(getStr(0))
		if path == "" {
			return nil // Silently ignore delete of empty path
		}
		os.Remove(path)
		return nil
	case "deletefolder":
		path := resolve(getStr(0))
		os.RemoveAll(path)
		return nil
	case "driveexists":
		// Check if drive letter exists
		d := getStr(0)
		if len(d) > 0 {
			// Simple check if we can stat the root?
			// On Windows, d + "\" or just d.
			if len(d) == 1 {
				d = d + ":"
			}
			_, err := os.Stat(d + string(os.PathSeparator))
			return err == nil
		}
		return false
	case "fileexists":
		path := resolve(getStr(0))
		if path == "" {
			return false
		}
		info, err := os.Stat(path)
		return err == nil && !info.IsDir()
	case "folderexists":
		path := resolve(getStr(0))
		info, err := os.Stat(path)
		return err == nil && info.IsDir()
	case "getabsolutepathname":
		path := resolve(getStr(0))
		abs, _ := filepath.Abs(path)
		return abs
	case "getbasename":
		return filepath.Base(strings.TrimSuffix(getStr(0), filepath.Ext(getStr(0))))
	case "getdrive":
		// Returns Drive object
		dName := getStr(0)
		return &FSODrive{letter: dName}
	case "getdrivename":
		return filepath.VolumeName(getStr(0))
	case "getextensionname":
		return strings.TrimPrefix(filepath.Ext(getStr(0)), ".")
	case "getfile":
		path := resolve(getStr(0))
		_, err := os.Stat(path)
		if err != nil {
			return nil
		}
		return &FSOFile{ctx: f.ctx, path: path}
	case "getfilename":
		return filepath.Base(getStr(0))
	case "getfolder":
		path := resolve(getStr(0))
		_, err := os.Stat(path)
		if err != nil {
			return nil
		}
		return &FSOFolder{ctx: f.ctx, path: path}
	case "getparentfoldername":
		return filepath.Dir(getStr(0))
	case "getspecialfolder":
		// 0=WindowsFolder, 1=SystemFolder, 2=TemporaryFolder
		i := 0
		if len(args) > 0 {
			i, _ = args[0].(int)
		}
		if i == 2 {
			return os.TempDir()
		}
		return "C:\\Windows" // Mock
	case "gettempname":
		return fmt.Sprintf("rad%X.tmp", time.Now().UnixNano())
	case "movefile":
		src := resolve(getStr(0))
		dst := resolve(getStr(1))
		os.Rename(src, dst)
		return nil
	case "movefolder":
		src := resolve(getStr(0))
		dst := resolve(getStr(1))
		os.Rename(src, dst)
		return nil
	case "opentextfile":
		// OpenTextFile(filename, [iomode], [create], [format])
		// iomode: 1=ForReading, 2=ForWriting, 8=ForAppending
		path := resolve(getStr(0))
		mode := 1
		create := false
		if len(args) > 1 {
			// Check type of arg[1]
			if i, ok := args[1].(int); ok {
				mode = i
			}
			if s, ok := args[1].(string); ok {
				fmt.Sscanf(s, "%d", &mode)
			}
		}
		if len(args) > 2 {
			if b, ok := args[2].(bool); ok {
				create = b
			}
		}

		flag := os.O_RDONLY
		if mode == 2 {
			flag = os.O_WRONLY | os.O_CREATE | os.O_TRUNC
		}
		if mode == 8 {
			flag = os.O_WRONLY | os.O_CREATE | os.O_APPEND
		}

		if !create && (flag&os.O_CREATE) != 0 {
			// If create is false, ensure we check existence?
			// Go OpenFile with O_CREATE creates it.
			// But if Create=false, we should probably fail if not exists?
			// But O_CREATE is needed for Write/Append usually?
			// Logic: If Create=False, check stat first?
			if _, err := os.Stat(path); os.IsNotExist(err) {
				return nil
			}
		}

		fp, err := os.OpenFile(path, flag, 0666)
		if err != nil {
			if create {
				fp, err = os.Create(path)
				if err != nil {
					return nil
				}
			} else {
				return nil
			}
		}
		return &TextStream{f: fp, mode: mode}
	}
	return nil
}

// --- TextStream ---
type TextStream struct {
	f       *os.File
	mode    int // 1=Read, 2=Write, 8=Append
	scanner *bufio.Scanner
	reader  *bufio.Reader
}

func (ts *TextStream) GetProperty(name string) interface{} {
	if strings.EqualFold(name, "AtEndOfStream") {
		// Check EOF
		// Hard without reading.
		// For now, return false unless we hit EOF in reading?
		// We can peek?
		return false
	}
	return nil
}
func (ts *TextStream) SetProperty(name string, value interface{}) {}
func (ts *TextStream) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)
	switch method {
	case "close":
		if ts.f != nil {
			ts.f.Close()
			ts.f = nil
		}
	case "read":
		n := 1
		if len(args) > 0 {
			n, _ = args[0].(int)
		}
		b := make([]byte, n)
		if ts.f != nil {
			ts.f.Read(b)
		}
		return string(b)
	case "readall":
		if ts.f == nil {
			return ""
		}
		// Seek start?
		ts.f.Seek(0, 0)
		b, _ := io.ReadAll(ts.f)
		return string(b)
	case "readline":
		if ts.f == nil {
			return ""
		}
		if ts.reader == nil {
			ts.reader = bufio.NewReader(ts.f)
		}
		line, _, _ := ts.reader.ReadLine()
		return string(line)
	case "write":
		if len(args) > 0 && ts.f != nil {
			ts.f.WriteString(fmt.Sprintf("%v", args[0]))
		}
	case "writeline":
		if ts.f != nil {
			s := ""
			if len(args) > 0 {
				s = fmt.Sprintf("%v", args[0])
			}
			ts.f.WriteString(s + "\r\n")
		}
	}
	return nil
}

// --- FSO Drive ---
type FSODrive struct {
	letter string
}

func (d *FSODrive) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "driveletter":
		return d.letter
	case "path":
		return d.letter + ":"
	case "drivetype":
		return 2 // Fixed
	case "filesystem":
		return "NTFS"
	case "isready":
		return true
	case "serialnumber":
		return 12345
	case "rootfolder":
		return &FSOFolder{path: d.letter + ":\\"}
	case "totalsize":
		return 10000
		//return 1024 * 1024 * 1024 * 100 // Mock 100GB
	case "freespace", "availablespace":
		return 10000
		//return 1024 * 1024 * 1024 * 50 // Mock 50GB
	}
	return nil
}
func (d *FSODrive) SetProperty(name string, value interface{})              {}
func (d *FSODrive) CallMethod(name string, args ...interface{}) interface{} { return nil }

// --- FSO Drives Collection ---
type FSODrives struct {
	ctx *ExecutionContext
}

func (ds *FSODrives) GetProperty(name string) interface{} {
	if strings.EqualFold(name, "Count") {
		return 1
	} // Mock
	if strings.EqualFold(name, "Item") {
		return nil
	} // Key needed
	return nil
}
func (ds *FSODrives) SetProperty(name string, value interface{}) {}
func (ds *FSODrives) CallMethod(name string, args ...interface{}) interface{} {
	if strings.EqualFold(name, "Item") && len(args) > 0 {
		return &FSODrive{letter: fmt.Sprintf("%v", args[0])}
	}
	return nil
}
func (ds *FSODrives) Enumeration() []interface{} {
	// Return list of FSODrive
	// On Windows, iterate C..Z?
	var list []interface{}
	// Mock C:
	list = append(list, &FSODrive{letter: "C"})
	return list
}

// --- FSO File ---
type FSOFile struct {
	ctx  *ExecutionContext
	path string
}

func (f *FSOFile) GetProperty(name string) interface{} {
	info, err := os.Stat(f.path)
	if err != nil {
		return nil
	}
	switch strings.ToLower(name) {
	case "attributes":
		return 0
	case "datecreated":
		return info.ModTime() // Unix doesn't have create time easily
	case "datelastaccessed":
		return info.ModTime()
	case "datelastmodified":
		return info.ModTime()
	case "drive":
		return &FSODrive{letter: filepath.VolumeName(f.path)}
	case "name":
		return info.Name()
	case "parentfolder":
		return &FSOFolder{ctx: f.ctx, path: filepath.Dir(f.path)}
	case "path":
		return f.path
	case "shortname":
		return info.Name() // Short names are hard
	case "shortpath":
		return f.path
	case "size":
		return int(info.Size())
	case "type":
		return "File"
	}
	return nil
}
func (f *FSOFile) SetProperty(name string, value interface{}) {}
func (f *FSOFile) CallMethod(name string, args ...interface{}) interface{} {
	switch strings.ToLower(name) {
	case "copy":
		// Copy(dest, [overwrite])
		if len(args) < 1 {
			return nil
		}
		dest := fmt.Sprintf("%v", args[0])
		data, _ := os.ReadFile(f.path)
		os.WriteFile(dest, data, 0666)
	case "delete":
		os.Remove(f.path)
	case "move":
		if len(args) < 1 {
			return nil
		}
		dest := fmt.Sprintf("%v", args[0])
		os.Rename(f.path, dest)
	case "openastextstream":
		// OpenAsTextStream([iomode], [format])
		mode := 1
		if len(args) > 0 {
			if i, ok := args[0].(int); ok {
				mode = i
			}
		}
		flag := os.O_RDONLY
		if mode == 2 {
			flag = os.O_WRONLY | os.O_CREATE
		}
		if mode == 8 {
			flag = os.O_WRONLY | os.O_CREATE | os.O_APPEND
		}
		fp, err := os.OpenFile(f.path, flag, 0666)
		if err != nil {
			return nil
		}
		return &TextStream{f: fp, mode: mode}
	}
	return nil
}

// --- FSO Files Collection ---
type FSOFiles struct {
	ctx        *ExecutionContext
	folderPath string
}

func (fs *FSOFiles) GetProperty(name string) interface{} {
	if strings.EqualFold(name, "Count") {
		entries, _ := os.ReadDir(fs.folderPath)
		count := 0
		for _, e := range entries {
			if !e.IsDir() {
				count++
			}
		}
		return count
	}
	return nil
}
func (fs *FSOFiles) SetProperty(name string, value interface{}) {}
func (fs *FSOFiles) CallMethod(name string, args ...interface{}) interface{} {
	if strings.EqualFold(name, "Item") && len(args) > 0 {
		// Key is name
		key := fmt.Sprintf("%v", args[0])
		p := filepath.Join(fs.folderPath, key)
		return &FSOFile{ctx: fs.ctx, path: p}
	}
	return nil
}
func (fs *FSOFiles) Enumeration() []interface{} {
	var list []interface{}
	entries, _ := os.ReadDir(fs.folderPath)
	for _, e := range entries {
		if !e.IsDir() {
			list = append(list, &FSOFile{ctx: fs.ctx, path: filepath.Join(fs.folderPath, e.Name())})
		}
	}
	return list
}

// --- FSO Folder ---
type FSOFolder struct {
	ctx  *ExecutionContext
	path string
}

func (f *FSOFolder) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "files":
		return &FSOFiles{ctx: f.ctx, folderPath: f.path}
	case "subfolders":
		return &FSOSubFolders{ctx: f.ctx, folderPath: f.path}
	case "name":
		return filepath.Base(f.path)
	case "path":
		return f.path
	case "parentfolder":
		return &FSOFolder{ctx: f.ctx, path: filepath.Dir(f.path)}
	case "isrootfolder":
		return false // Simplify
	case "size":
		return 0 // Recursive size calc?
	case "datecreated", "datelastmodified":
		info, _ := os.Stat(f.path)
		if info != nil {
			return info.ModTime()
		}
		return time.Now()
	}
	return nil
}
func (f *FSOFolder) SetProperty(name string, value interface{}) {}
func (f *FSOFolder) CallMethod(name string, args ...interface{}) interface{} {
	// Copy, Delete, Move, CreateTextFile
	switch strings.ToLower(name) {
	case "copy":
		if len(args) < 1 {
			return nil
		}
		// dest := fmt.Sprintf("%v", args[0])
		// TODO IMPLEMENT Recursive copy not implemented Use filepath.Walk to copy.
		return nil
	case "delete":
		os.RemoveAll(f.path)
		return nil
	case "move":
		if len(args) < 1 {
			return nil
		}
		dest := fmt.Sprintf("%v", args[0])
		os.Rename(f.path, dest)
		return nil
	case "createtextfile":
		if len(args) < 1 {
			return nil
		}
		fname := fmt.Sprintf("%v", args[0])
		fpath := filepath.Join(f.path, fname)
		fp, _ := os.Create(fpath)
		return &TextStream{f: fp, mode: 2}
	}
	return nil
}

// --- FSO SubFolders Collection ---
type FSOSubFolders struct {
	ctx        *ExecutionContext
	folderPath string
}

func (fs *FSOSubFolders) GetProperty(name string) interface{} {
	if strings.EqualFold(name, "Count") {
		entries, _ := os.ReadDir(fs.folderPath)
		count := 0
		for _, e := range entries {
			if e.IsDir() {
				count++
			}
		}
		return count
	}
	return nil
}
func (fs *FSOSubFolders) SetProperty(name string, value interface{}) {}
func (fs *FSOSubFolders) CallMethod(name string, args ...interface{}) interface{} {
	if strings.EqualFold(name, "Item") && len(args) > 0 {
		key := fmt.Sprintf("%v", args[0])
		p := filepath.Join(fs.folderPath, key)
		return &FSOFolder{ctx: fs.ctx, path: p}
	}
	if strings.EqualFold(name, "Add") && len(args) > 0 {
		name := fmt.Sprintf("%v", args[0])
		p := filepath.Join(fs.folderPath, name)
		os.MkdirAll(p, 0755)
		return &FSOFolder{ctx: fs.ctx, path: p}
	}
	return nil
}
func (fs *FSOSubFolders) Enumeration() []interface{} {
	var list []interface{}
	entries, _ := os.ReadDir(fs.folderPath)
	for _, e := range entries {
		if e.IsDir() {
			list = append(list, &FSOFolder{ctx: fs.ctx, path: filepath.Join(fs.folderPath, e.Name())})
		}
	}
	return list
}

// --- ADODB.Stream ---

// ADODBStream implements ADODB.Stream for binary and text file operations
type ADODBStream struct {
	Type          int    // 1 = adTypeBinary, 2 = adTypeText
	Mode          int    // Read/Write mode
	State         int    // 0 = closed, 1 = open
	Position      int64  // Current position in stream
	Size          int64  // Size of stream
	Charset       string // Character set (default "utf-8")
	LineSeparator int    // Line separator type
	buffer        []byte // Internal buffer
	lastFile      string // Last file loaded via LoadFromFile
	ctx           *ExecutionContext
}

// NewADODBStream creates a new ADODB.Stream object
func NewADODBStream(ctx *ExecutionContext) *ADODBStream {
	return &ADODBStream{
		Type:          2,            // adTypeText by default (2 = text, 1 = binary)
		Mode:          3,            // adModeReadWrite
		State:         0,            // closed
		Charset:       "iso-8859-1", // Default to binary-safe charset (was utf-8, but that corrupts binary uploads)
		LineSeparator: 13,           // adCRLF
		buffer:        make([]byte, 0),
		ctx:           ctx,
	}
}

func (s *ADODBStream) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "type":
		return s.Type
	case "mode":
		return s.Mode
	case "state":
		return s.State
	case "position":
		return int(s.Position)
	case "size":
		return int(s.Size)
	case "charset":
		return s.Charset
	case "lineseparator":
		return s.LineSeparator
	case "eos": // End of stream
		return s.Position >= s.Size
	case "readtext":
		// Support VBScript calling ReadText without parentheses (as property)
		// This is common: strTmp = objStream.ReadText
		return s.CallMethod("readtext")
	}
	return nil
}

func (s *ADODBStream) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "type":
		s.Type = toInt(value)
	case "mode":
		s.Mode = toInt(value)
	case "charset":
		if v, ok := value.(string); ok {
			s.Charset = v
		}
	case "lineseparator":
		s.LineSeparator = toInt(value)
	case "position":
		var newPos int64
		switch v := value.(type) {
		case int:
			newPos = int64(v)
		case int64:
			newPos = v
		case float64:
			newPos = int64(v)
		default:
			log.Printf("ADODBStream.SetProperty position: could not convert type %T\n", value)
			return
		}
		if newPos < 0 {
			newPos = 0
		}
		//fmt.Printf("[DEBUG] ADODBStream.SetProperty Position: setting to %d (was %d), Size=%d\n", newPos, s.Position, s.Size)
		s.Position = newPos
	}
}

func (s *ADODBStream) CallMethod(name string, args ...interface{}) interface{} {
	method := strings.ToLower(name)

	switch method {
	case "open":
		// Open([Source], [Mode], [Options], [UserName], [Password])
		s.State = 1
		s.Position = 0
		s.buffer = make([]byte, 0)
		s.Size = 0
		return nil

	case "close":
		s.State = 0
		s.buffer = nil
		return nil

	case "read":
		// Read([NumBytes]) - reads binary data
		// Returns all data from current position if no argument specified
		// Does NOT check State - allows reading even if stream wasn't explicitly opened
		numBytes := int64(-1) // -1 = read all
		if len(args) > 0 {
			if v, ok := args[0].(int); ok {
				numBytes = int64(v)
			}
		}

		if numBytes == -1 {
			// Read all from current position to end
			if s.Position >= s.Size || len(s.buffer) == 0 {
				return ""
			}
			data := s.buffer[s.Position:]
			s.Position = s.Size
			return string(data)
		}

		if s.Position+numBytes > s.Size {
			numBytes = s.Size - s.Position
		}

		if numBytes <= 0 {
			return ""
		}

		data := s.buffer[s.Position : s.Position+numBytes]
		s.Position += numBytes
		return string(data)

	case "readtext":
		// ReadText([NumChars]) - reads text data
		// Returns all text from current position if no argument specified
		if s.State != 1 {
			if len(s.buffer) == 0 && s.lastFile != "" {
				data, err := os.ReadFile(s.lastFile)
				if err == nil {
					s.buffer = data
					s.Size = int64(len(data))
					s.Position = 0
					s.State = 1
				}
			}
			if s.State != 1 {
				log.Printf("ADODBStream ReadText error: Stream not open (State=%d)\n", s.State)
				return ""
			}
		}

		numChars := int64(-1) // -1 = read all
		if len(args) > 0 {
			if v, ok := args[0].(int); ok {
				numChars = int64(v)
			}
		}

		// Calculate bytes to read
		remaining := s.Size - s.Position
		if remaining <= 0 {
			// fmt.Println("[DEBUG] ADODBStream.ReadText: End of stream reached.")
			return ""
		}

		bytesToRead := remaining
		if numChars != -1 && numChars < remaining {
			// Note: For multi-byte charsets (UTF-8), numChars might convert to more bytes
			// But for simplicity/performance we treat it as bytes here unless it's strict
			// For ISO-8859-1 it is exactly bytes.
			bytesToRead = numChars
		}

		data := s.buffer[s.Position : s.Position+bytesToRead]
		s.Position += bytesToRead

		// Decode based on charset
		cs := strings.ToLower(s.Charset)

		retStr := ""
		if cs == "iso-8859-1" || cs == "windows-1252" || cs == "ascii" || cs == "us-ascii" {
			retStr = decodeSingleByteString(data)
		} else if cs == "unicode" || cs == "utf-16" {
			retStr = decodeUTF16(data, binary.LittleEndian)
		} else if cs == "utf-8" || cs == "utf8" {
			retStr = string(bytes.TrimPrefix(data, []byte{0xEF, 0xBB, 0xBF}))
		} else {
			// Default UTF-8
			retStr = string(data)
		}

		preview := retStr
		if len(preview) > 50 {
			preview = preview[:50] + "..."
		}
		preview = strings.ReplaceAll(preview, "\n", "\\n")
		preview = strings.ReplaceAll(preview, "\r", "\\r")

		//fmt.Printf("[DEBUG] ADODBStream.ReadText: Reading %d bytes (requested=%d), Charset=%s, Pos=%d/%d, Content=%q\n", bytesToRead, numChars, cs, s.Position, s.Size, preview)
		return retStr

	case "write":
		// Write(Buffer) - writes binary data
		if len(args) < 1 || s.State != 1 {
			return nil
		}

		var data []byte
		switch v := args[0].(type) {
		case []byte:
			data = v
		case string:
			// If string passed to Write (Binary), what encoding?
			// Usually Write takes binary (Variant array of bytes).
			// If we get string, maybe we should encode it using current Charset?
			// Or just raw bytes of UTF-8 string?
			// Standard behavior: Write expects Variant (Array of Bytes).
			// If string is passed, it might be coerced.
			// Let's assume input string is already correct bytes or UTF-8.
			data = []byte(v)
		default:
			data = []byte(fmt.Sprintf("%v", v))
		}

		// Insert data at current position
		if s.Position >= int64(len(s.buffer)) {
			s.buffer = append(s.buffer, data...)
		} else {
			// Overwrite at position
			// If we need to expand buffer
			needed := s.Position + int64(len(data))
			if needed > int64(len(s.buffer)) {
				// Extend buffer
				newBuf := make([]byte, needed)
				copy(newBuf, s.buffer)
				s.buffer = newBuf
			}
			copy(s.buffer[s.Position:], data)
		}

		s.Position += int64(len(data))
		if s.Position > s.Size {
			s.Size = s.Position
		}

		//fmt.Printf("[DEBUG] ADODBStream.Write: Wrote %d bytes. Pos=%d/%d\n", len(data), s.Position, s.Size)
		return nil

	case "writetext":
		// WriteText(Data, [Options]) - writes text data
		// For binary data ([]byte), we store bytes directly without any charset conversion
		// This is critical for multipart upload parsing where bytes must be preserved exactly
		if len(args) < 1 || s.State != 1 {
			return nil
		}

		var data []byte

		switch v := args[0].(type) {
		case []byte:
			// CRITICAL: Store bytes directly without any conversion
			// This preserves binary data exactly as received from BinaryRead
			data = make([]byte, len(v))
			copy(data, v)
			//if len(v) > 50 {
			//	fmt.Printf("[DEBUG] ADODBStream.WriteText: Input []byte, len=%d, first50=%x\n", len(v), v[:50])
			//} else {
			//	fmt.Printf("[DEBUG] ADODBStream.WriteText: Input []byte, len=%d, data=%x\n", len(v), v)
			//}
		case string:
			// For string input, encode based on charset
			cs := strings.ToLower(s.Charset)
			if cs == "iso-8859-1" || cs == "windows-1252" || cs == "ascii" || cs == "us-ascii" {
				data = encodeSingleByteString(v)
			} else if cs == "unicode" || cs == "utf-16" {
				runes := []rune(v)
				data = make([]byte, len(runes)*2)
				for i, r := range runes {
					binary.LittleEndian.PutUint16(data[i*2:], uint16(r))
				}
			} else {
				data = []byte(v)
			}
			//if len(v) > 50 {
			//	fmt.Printf("[DEBUG] ADODBStream.WriteText: Input string, len=%d, charset=%s\n", len(v), cs)
			//}
		default:
			data = []byte(fmt.Sprintf("%v", args[0]))
		}

		// Append or write based on options
		options := 0 // 0 = default, 1 = adWriteLine
		if len(args) > 1 {
			if v, ok := args[1].(int); ok {
				options = v
			}
		}

		// Add line separator if adWriteLine (1)
		if options == 1 {
			switch s.LineSeparator {
			case 10: // adLF
				data = append(data, '\n')
			case 13: // adCR
				data = append(data, '\r')
			case -1, 0: // adCRLF
				data = append(data, '\r', '\n')
			}
		}

		// Write at current position
		if s.Position >= int64(len(s.buffer)) {
			s.buffer = append(s.buffer, data...)
		} else {
			// Overwrite at position
			needed := s.Position + int64(len(data))
			if needed > int64(len(s.buffer)) {
				newBuf := make([]byte, needed)
				copy(newBuf, s.buffer)
				s.buffer = newBuf
			}
			copy(s.buffer[s.Position:], data)
		}

		s.Position += int64(len(data))
		if s.Position > s.Size {
			s.Size = s.Position
		}

		//fmt.Printf("[DEBUG] ADODBStream.WriteText: Wrote %d bytes, Pos=%d/%d\n", len(data), s.Position, s.Size)
		return nil

	case "loadfromfile":
		// LoadFromFile(FileName) - loads file into stream
		// FileName should be a full path or relative to the current working directory
		// (It's typically already mapped via Server.MapPath in the ASP call)
		if len(args) < 1 || args[0] == nil {
			errMsg := "LoadFromFile requires a valid filename argument"
			log.Printf("ADODBStream Error: %s\n", errMsg)
			s.ctx.Err.SetError(fmt.Errorf("%s", errMsg))
			return nil
		}

		filename := fmt.Sprintf("%v", args[0])
		//fmt.Printf("[DEBUG] ADODBStream.LoadFromFile called with: %s\n", filename)

		// Validate filename is not empty or nil
		if filename == "" || filename == "<nil>" {
			errMsg := "LoadFromFile received empty or nil filename"
			log.Printf("ADODBStream Error: %s\n", errMsg)
			s.ctx.Err.SetError(fmt.Errorf("%s", errMsg))
			return nil
		}

		// Don't call Server_MapPath here - the ASP script should have already done it
		fullPath := filename

		// Security check
		rootDir, _ := filepath.Abs(s.ctx.RootDir)
		absPath, _ := filepath.Abs(fullPath)
		if !strings.HasPrefix(strings.ToLower(absPath), strings.ToLower(rootDir)) {
			errMsg := fmt.Sprintf("access denied: %s (outside root %s)", absPath, rootDir)
			log.Printf("ADODBStream Security Error: %s\n", errMsg)
			s.ctx.Err.SetError(fmt.Errorf("access denied: %s", filename))
			return nil
		}

		data, err := os.ReadFile(fullPath)
		if err != nil {
			log.Printf("ADODBStream LoadFromFile error reading %s: %v\n", fullPath, err)
			s.ctx.Err.SetError(err)
			return nil
		}

		// Note: We do NOT decode here even if Type is Text (2).
		// We store raw bytes in the buffer. ReadText will decode them based on Charset.
		// This ensures that switching between Binary and Text (with specific Charset) works correctly.

		// Ensure stream is open
		s.State = 1
		s.buffer = data
		s.Size = int64(len(data))
		s.Position = 0
		s.lastFile = fullPath

		//fmt.Printf("[DEBUG] ADODBStream.LoadFromFile: Loaded %d bytes from %s. Charset=%s\n", s.Size, filepath.Base(fullPath), s.Charset)

		// Success - no logging needed for normal operation
		return nil

	case "savetofile":
		// SaveToFile(FileName, [Options]) - saves stream to file
		if len(args) < 1 || args[0] == nil || s.State != 1 {
			if len(args) < 1 || args[0] == nil {
				log.Println("Error: SaveToFile requires a valid filename argument")
			}
			return nil
		}

		filename := fmt.Sprintf("%v", args[0])

		// Validate filename is not empty or nil
		if filename == "" || filename == "<nil>" {
			log.Println("Error: SaveToFile received empty or nil filename")
			return nil
		}

		// Check if path is already absolute (has drive letter on Windows or starts with / on Unix)
		fullPath := filename
		if !filepath.IsAbs(filename) {
			fullPath = s.ctx.Server_MapPath(filename)
		}
		// log.Printf("[DEBUG] SaveToFile: fullPath=%q\n", fullPath)

		// Validate mapped path
		if fullPath == "" || fullPath == "<nil>" {
			log.Printf("Error: Server_MapPath returned invalid path for %s\n", filename)
			return nil
		}

		// Security check
		rootDir, _ := filepath.Abs(s.ctx.RootDir)
		absPath, _ := filepath.Abs(fullPath)
		if !strings.HasPrefix(strings.ToLower(absPath), strings.ToLower(rootDir)) {
			log.Printf("Security Warning: Script tried to access %s (Root: %s)\n", absPath, rootDir)
			return nil
		}

		// Options: 1 = adSaveCreateNotExist (default), 2 = adSaveCreateOverWrite
		options := 2 // Default to overwrite
		if len(args) > 1 {
			if v, ok := args[1].(int); ok {
				options = v
			}
		}

		// Check if file exists
		if options == 1 {
			if _, err := os.Stat(fullPath); err == nil {
				// File exists, don't overwrite
				fmt.Printf("[DEBUG] ADODBStream.SaveToFile: File exists and Overwrite=False. Skipping %s\n", filepath.Base(fullPath))
				return nil
			}
		}

		err := os.WriteFile(fullPath, s.buffer, 0644)
		if err != nil {
			fmt.Printf("[DEBUG] ADODBStream.SaveToFile: Write failed: %v\n", err)
			return nil
		}
		//fmt.Printf("[DEBUG] ADODBStream.SaveToFile: Saved %d bytes to %s\n", len(s.buffer), filepath.Base(fullPath))
		return nil

	case "copyto":
		// CopyTo(DestStream, [CharNumber]) - copies data to another stream
		if len(args) < 1 {
			return nil
		}

		// Accept both *ADODBStream and *ADOStream (wrapper)
		var destStream *ADODBStream
		switch v := args[0].(type) {
		case *ADODBStream:
			destStream = v
		case *ADOStream:
			destStream = v.lib
		default:
			log.Printf("ADODBStream.CopyTo: arg[0] is not a stream type, got %T\n", args[0])
			return nil
		}

		numChars := int64(-1)
		if len(args) > 1 {
			switch v := args[1].(type) {
			case int:
				numChars = int64(v)
			case int64:
				numChars = v
			case float64:
				numChars = int64(v)
			}
		}

		if s.Position < 0 {
			return nil
		}

		if numChars == -1 {
			numChars = s.Size - s.Position
		}

		if s.Position+numChars > s.Size {
			numChars = s.Size - s.Position
		}

		if numChars <= 0 {
			return nil
		}

		// Ensure we don't go out of bounds of actual buffer
		if s.Position >= int64(len(s.buffer)) {
			return nil
		}
		if s.Position+numChars > int64(len(s.buffer)) {
			numChars = int64(len(s.buffer)) - s.Position
		}

		data := s.buffer[s.Position : s.Position+numChars]
		// Debug: show first few bytes of data being copied
		//preview := ""
		//if len(data) > 30 {
		//	preview = string(data[:30]) + "..."
		//} else {
		//	preview = string(data)
		//}
		//fmt.Printf("[DEBUG] ADODBStream.CopyTo: Position=%d, numChars=%d, data[0:30]=%q\n", s.Position, numChars, preview)
		destStream.buffer = append(destStream.buffer, data...)
		destStream.Size = int64(len(destStream.buffer))
		s.Position += numChars

		return nil

	case "flush":
		// Flush() - writes buffer to underlying storage
		// In our implementation, buffer is already in memory
		return nil

	case "seteos":
		// SetEOS() - sets size to current position
		s.Size = s.Position
		if int64(len(s.buffer)) > s.Size {
			s.buffer = s.buffer[:s.Size]
		}
		return nil

	case "skipline":
		// SkipLine() - skips to next line
		if s.State != 1 {
			return nil
		}

		// Find next line separator
		for s.Position < s.Size {
			if s.buffer[s.Position] == '\n' {
				s.Position++
				break
			}
			if s.buffer[s.Position] == '\r' {
				s.Position++
				if s.Position < s.Size && s.buffer[s.Position] == '\n' {
					s.Position++
				}
				break
			}
			s.Position++
		}
		return nil
	}

	return nil
}
