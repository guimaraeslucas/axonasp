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
	"bufio"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strings"
	"time"

	"g3pix.com.br/axonasp/vbscript"
	"github.com/ricochet2200/go-disk-usage/du"
)

const (
	fsoKindRoot int = iota + 1
	fsoKindFile
	fsoKindFolder
	fsoKindTextStream
	fsoKindFilesCollection
	fsoKindSubFoldersCollection
	fsoKindDrive
	fsoKindDrivesCollection
)

// fsoNativeObject stores one dynamic FileSystemObject runtime instance.
type fsoNativeObject struct {
	kind   int
	path   string
	drive  string
	stream *fsoTextStreamState
}

// fsoTextStreamState stores TextStream runtime state.
type fsoTextStreamState struct {
	file     *os.File
	reader   *bufio.Reader
	mode     int
	atEnd    bool
	line     int
	column   int
	standard bool
}

// newFSORootObject creates a root Scripting.FileSystemObject instance.
func (vm *VM) newFSORootObject() Value {
	return vm.newFSONativeObject(fsoKindRoot, "", nil)
}

// newFSONativeObject stores one FSO runtime object and returns its native VM value.
func (vm *VM) newFSONativeObject(kind int, path string, stream *fsoTextStreamState) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.fsoItems[objID] = &fsoNativeObject{kind: kind, path: path, stream: stream}
	return Value{Type: VTNativeObject, Num: objID}
}

// newFSODriveObject creates one FSO Drive runtime object.
func (vm *VM) newFSODriveObject(drive string) Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.fsoItems[objID] = &fsoNativeObject{kind: fsoKindDrive, drive: drive}
	return Value{Type: VTNativeObject, Num: objID}
}

// newFSODrivesCollectionObject creates the FSO Drives runtime collection object.
func (vm *VM) newFSODrivesCollectionObject() Value {
	objID := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.fsoItems[objID] = &fsoNativeObject{kind: fsoKindDrivesCollection}
	return Value{Type: VTNativeObject, Num: objID}
}

// dispatchFSOMethod executes methods and property Let/Set calls for all FSO runtime objects.
func (vm *VM) dispatchFSOMethod(objID int64, member string, args []Value) (Value, bool) {
	obj, exists := vm.fsoItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	switch obj.kind {
	case fsoKindRoot:
		return vm.dispatchFSORootMethod(obj, member, args), true
	case fsoKindFile:
		return vm.dispatchFSOFileMethod(obj, member, args), true
	case fsoKindFolder:
		return vm.dispatchFSOFolderMethod(obj, member, args), true
	case fsoKindTextStream:
		return vm.dispatchFSOTextStreamMethod(obj, member, args), true
	case fsoKindFilesCollection:
		return vm.dispatchFSOFilesCollectionMethod(obj, member, args), true
	case fsoKindSubFoldersCollection:
		return vm.dispatchFSOSubFoldersCollectionMethod(obj, member, args), true
	case fsoKindDrive:
		return vm.dispatchFSODriveMethod(obj, member, args), true
	case fsoKindDrivesCollection:
		return vm.dispatchFSODrivesCollectionMethod(obj, member, args), true
	default:
		return Value{Type: VTEmpty}, true
	}
}

// dispatchFSOPropertySet handles property Let assignments for FSO runtime objects.
// Most FSO properties are read-only. Name is writable on File and Folder objects,
// so this handler routes those assignments through the same rename logic used by
// the method-dispatch path and ignores the remaining properties for compatibility.
func (vm *VM) dispatchFSOPropertySet(objID int64, member string, val Value) bool {
	obj, exists := vm.fsoItems[objID]
	if !exists {
		return false
	}
	switch obj.kind {
	case fsoKindFile:
		if strings.EqualFold(member, "Name") {
			vm.dispatchFSOFileMethod(obj, "Name", []Value{val})
		}
	case fsoKindFolder:
		if strings.EqualFold(member, "Name") {
			vm.dispatchFSOFolderMethod(obj, "Name", []Value{val})
		}
	}
	return true
}

// dispatchFSOPropertyGet resolves property reads for all FSO runtime objects.
func (vm *VM) dispatchFSOPropertyGet(objID int64, member string) (Value, bool) {
	obj, exists := vm.fsoItems[objID]
	if !exists {
		return Value{Type: VTEmpty}, false
	}

	switch obj.kind {
	case fsoKindRoot:
		return vm.dispatchFSORootPropertyGet(obj, member), true
	case fsoKindFile:
		return vm.dispatchFSOFilePropertyGet(obj, member), true
	case fsoKindFolder:
		return vm.dispatchFSOFolderPropertyGet(obj, member), true
	case fsoKindTextStream:
		return vm.dispatchFSOTextStreamPropertyGet(obj, member), true
	case fsoKindFilesCollection:
		return vm.dispatchFSOFilesCollectionPropertyGet(obj, member), true
	case fsoKindSubFoldersCollection:
		return vm.dispatchFSOSubFoldersCollectionPropertyGet(obj, member), true
	case fsoKindDrive:
		return vm.dispatchFSODrivePropertyGet(obj, member), true
	case fsoKindDrivesCollection:
		return vm.dispatchFSODrivesCollectionPropertyGet(obj, member), true
	default:
		return Value{Type: VTEmpty}, true
	}
}

// dispatchFSORootMethod handles Scripting.FileSystemObject method calls.
func (vm *VM) dispatchFSORootMethod(_ *fsoNativeObject, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "FileExists"):
		if len(args) < 1 {
			return NewBool(false)
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			return NewBool(false)
		}
		info, err := os.Stat(path)
		return NewBool(err == nil && !info.IsDir())
	case strings.EqualFold(member, "FolderExists"):
		if len(args) < 1 {
			return NewBool(false)
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			return NewBool(false)
		}
		info, err := os.Stat(path)
		return NewBool(err == nil && info.IsDir())
	case strings.EqualFold(member, "CreateFolder"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok || os.MkdirAll(path, 0755) != nil {
			return Value{Type: VTEmpty}
		}
		return vm.newFSONativeObject(fsoKindFolder, path, nil)
	case strings.EqualFold(member, "DeleteFolder"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			vm.raise(vbscript.PathNotFound, "Path not found")
			return Value{Type: VTEmpty}
		}
		info, err := os.Stat(path)
		if err != nil {
			if os.IsNotExist(err) {
				vm.raise(vbscript.PathNotFound, "Path not found")
			} else if os.IsPermission(err) {
				vm.raise(vbscript.PermissionDenied, "Permission denied")
			} else {
				vm.raise(vbscript.PathNotFound, "Path not found")
			}
			return Value{Type: VTEmpty}
		}
		if !info.IsDir() {
			vm.raise(vbscript.PathNotFound, "Path not found")
			return Value{Type: VTEmpty}
		}
		if err := vm.fsoDeletePath(path, true); err != nil {
			if os.IsPermission(err) {
				vm.raise(vbscript.PermissionDenied, "Permission denied")
			} else if os.IsNotExist(err) {
				vm.raise(vbscript.PathNotFound, "Path not found")
			} else {
				vm.raise(vbscript.PathNotFound, "Path not found")
			}
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "DeleteFile"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			vm.raise(vbscript.PathNotFound, "Path not found")
			return Value{Type: VTEmpty}
		}
		info, err := os.Stat(path)
		if err != nil {
			if os.IsNotExist(err) {
				vm.raise(vbscript.FileNotFound, "File not found")
			} else if os.IsPermission(err) {
				vm.raise(vbscript.PermissionDenied, "Permission denied")
			} else {
				vm.raise(vbscript.FileNotFound, "File not found")
			}
			return Value{Type: VTEmpty}
		}
		if info.IsDir() {
			vm.raise(vbscript.FileNotFound, "File not found")
			return Value{Type: VTEmpty}
		}
		if err := vm.fsoDeletePath(path, false); err != nil {
			if os.IsPermission(err) {
				vm.raise(vbscript.PermissionDenied, "Permission denied")
			} else if os.IsNotExist(err) {
				vm.raise(vbscript.FileNotFound, "File not found")
			} else {
				vm.raise(vbscript.FileNotFound, "File not found")
			}
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "GetFile"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			vm.raise(vbscript.PathNotFound, "Path not found")
			return Value{Type: VTEmpty}
		}
		info, err := os.Stat(path)
		if err != nil {
			if os.IsNotExist(err) {
				vm.raise(vbscript.FileNotFound, "File not found")
			} else if os.IsPermission(err) {
				vm.raise(vbscript.PermissionDenied, "Permission denied")
			} else {
				vm.raise(vbscript.FileNotFound, "File not found")
			}
			return Value{Type: VTEmpty}
		}
		if info.IsDir() {
			vm.raise(vbscript.FileNotFound, "File not found")
			return Value{Type: VTEmpty}
		}
		return vm.newFSONativeObject(fsoKindFile, path, nil)
	case strings.EqualFold(member, "GetFolder"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			vm.raise(vbscript.PathNotFound, "Path not found")
			return Value{Type: VTEmpty}
		}
		info, err := os.Stat(path)
		if err != nil {
			if os.IsNotExist(err) {
				vm.raise(vbscript.PathNotFound, "Path not found")
			} else if os.IsPermission(err) {
				vm.raise(vbscript.PermissionDenied, "Permission denied")
			} else {
				vm.raise(vbscript.PathNotFound, "Path not found")
			}
			return Value{Type: VTEmpty}
		}
		if !info.IsDir() {
			vm.raise(vbscript.PathNotFound, "Path not found")
			return Value{Type: VTEmpty}
		}
		return vm.newFSONativeObject(fsoKindFolder, path, nil)
	case strings.EqualFold(member, "BuildPath"):
		if len(args) < 2 {
			return NewString("")
		}
		return NewString(filepath.Join(args[0].String(), args[1].String()))
	case strings.EqualFold(member, "GetFileName"):
		if len(args) < 1 {
			return NewString("")
		}
		return NewString(filepath.Base(args[0].String()))
	case strings.EqualFold(member, "GetBaseName"):
		if len(args) < 1 {
			return NewString("")
		}
		name := filepath.Base(args[0].String())
		return NewString(strings.TrimSuffix(name, filepath.Ext(name)))
	case strings.EqualFold(member, "GetExtensionName"):
		if len(args) < 1 {
			return NewString("")
		}
		ext := filepath.Ext(args[0].String())
		return NewString(strings.TrimPrefix(ext, "."))
	case strings.EqualFold(member, "GetParentFolderName"):
		if len(args) < 1 {
			return NewString("")
		}
		return NewString(filepath.Dir(args[0].String()))
	case strings.EqualFold(member, "GetAbsolutePathName"):
		if len(args) < 1 {
			return NewString("")
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			return NewString("")
		}
		return NewString(path)
	case strings.EqualFold(member, "GetTempName"):
		return NewString(fmt.Sprintf("rad%X.axon.tmp", time.Now().UnixNano()))
	case strings.EqualFold(member, "GetDriveName"):
		if len(args) < 1 {
			return NewString("")
		}
		return NewString(vm.fsoDriveNameFromPath(args[0].String()))
	case strings.EqualFold(member, "GetDrive"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		driveName := vm.fsoDriveNameFromPath(args[0].String())
		if strings.TrimSpace(driveName) == "" {
			return Value{Type: VTEmpty}
		}
		return vm.newFSODriveObject(driveName)
	case strings.EqualFold(member, "GetStandardStream"):
		streamType := 0
		if len(args) >= 1 {
			streamType = vm.asInt(args[0])
		}
		return vm.fsoGetStandardStream(streamType)
	case strings.EqualFold(member, "GetFileVersion"):
		if len(args) < 1 {
			return NewString("1.0.0.0")
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			return NewString("1.0.0.0")
		}
		return NewString(fsoGetFileVersion(path))
	case strings.EqualFold(member, "DriveExists"):
		if len(args) < 1 {
			return NewBool(false)
		}
		drivePath := vm.fsoDriveRootFromName(args[0].String())
		if strings.TrimSpace(drivePath) == "" {
			return NewBool(false)
		}
		_, err := os.Stat(drivePath)
		return NewBool(err == nil)
	case strings.EqualFold(member, "MoveFile"):
		if len(args) < 2 {
			return Value{Type: VTEmpty}
		}
		sourcePath, sourceOK := vm.fsoResolvePath(args[0].String())
		destPath, destOK := vm.fsoResolvePath(args[1].String())
		if sourceOK && destOK {
			_ = vm.fsoMovePath(sourcePath, destPath)
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "MoveFolder"):
		if len(args) < 2 {
			return Value{Type: VTEmpty}
		}
		sourcePath, sourceOK := vm.fsoResolvePath(args[0].String())
		destPath, destOK := vm.fsoResolvePath(args[1].String())
		if sourceOK && destOK {
			_ = vm.fsoMovePath(sourcePath, destPath)
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "CopyFile"):
		if len(args) < 2 {
			return Value{Type: VTEmpty}
		}
		sourcePath, sourceOK := vm.fsoResolvePath(args[0].String())
		destPath, destOK := vm.fsoResolvePath(args[1].String())
		if sourceOK && destOK {
			_, _ = vm.fsoCopyFile(sourcePath, destPath, vm.fsoOverwriteDefault(args, 2, true))
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "CopyFolder"):
		if len(args) < 2 {
			return Value{Type: VTEmpty}
		}
		sourcePath, sourceOK := vm.fsoResolvePath(args[0].String())
		destPath, destOK := vm.fsoResolvePath(args[1].String())
		if sourceOK && destOK {
			_ = vm.fsoCopyFolder(sourcePath, destPath, vm.fsoOverwriteDefault(args, 2, true))
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "CreateTextFile"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			return Value{Type: VTEmpty}
		}
		overwrite := vm.fsoOverwriteDefault(args, 1, true)
		flags := os.O_CREATE | os.O_WRONLY
		if overwrite {
			flags |= os.O_TRUNC
		} else {
			flags |= os.O_EXCL
		}
		fileHandle, err := os.OpenFile(path, flags, 0644)
		if err != nil {
			return Value{Type: VTEmpty}
		}
		state := &fsoTextStreamState{file: fileHandle, mode: 2, atEnd: false, line: 1, column: 1}
		return vm.newFSONativeObject(fsoKindTextStream, path, state)
	case strings.EqualFold(member, "OpenTextFile"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		path, ok := vm.fsoResolvePath(args[0].String())
		if !ok {
			return Value{Type: VTEmpty}
		}
		mode := 1
		if len(args) >= 2 {
			mode = vm.asInt(args[1])
		}
		create := false
		if len(args) >= 3 {
			create = vm.asBool(args[2])
		}
		stream := vm.fsoOpenTextStream(path, mode, create)
		if stream == nil {
			return Value{Type: VTEmpty}
		}
		return vm.newFSONativeObject(fsoKindTextStream, path, stream)
	case strings.EqualFold(member, "GetSpecialFolder"):
		folderType := 0
		if len(args) >= 1 {
			folderType = vm.asInt(args[0])
		}
		switch folderType {
		case 2:
			return NewString(os.TempDir())
		case 1:
			if runtime.GOOS == "windows" {
				systemRoot := strings.TrimSpace(os.Getenv("SystemRoot"))
				if systemRoot == "" {
					systemRoot = strings.TrimSpace(os.Getenv("WINDIR"))
				}
				if systemRoot == "" {
					systemRoot = `C:\Windows`
				}
				return NewString(filepath.Join(systemRoot, "System32"))
			}
			return NewString("/usr/bin")
		default:
			if runtime.GOOS == "windows" {
				windowsDir := strings.TrimSpace(os.Getenv("WINDIR"))
				if windowsDir == "" {
					windowsDir = strings.TrimSpace(os.Getenv("SystemRoot"))
				}
				if windowsDir == "" {
					windowsDir = `C:\Windows`
				}
				return NewString(windowsDir)
			}
			return NewString(string(os.PathSeparator))
		}
	default:
		return Value{Type: VTEmpty}
	}
}

// dispatchFSORootPropertyGet resolves Scripting.FileSystemObject root properties.
func (vm *VM) dispatchFSORootPropertyGet(_ *fsoNativeObject, member string) Value {
	if strings.EqualFold(member, "Drives") {
		return vm.newFSODrivesCollectionObject()
	}
	return Value{Type: VTEmpty}
}

// dispatchFSODriveMethod handles FSODrive methods.
func (vm *VM) dispatchFSODriveMethod(_ *fsoNativeObject, _ string, _ []Value) Value {
	return Value{Type: VTEmpty}
}

// dispatchFSODrivesCollectionMethod handles Drives collection calls.
func (vm *VM) dispatchFSODrivesCollectionMethod(_ *fsoNativeObject, member string, args []Value) Value {
	if member == "" || strings.EqualFold(member, "Item") {
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		driveName := vm.fsoDriveNameFromPath(args[0].String())
		if strings.TrimSpace(driveName) == "" {
			return Value{Type: VTEmpty}
		}
		return vm.newFSODriveObject(driveName)
	}
	return Value{Type: VTEmpty}
}

// dispatchFSODrivePropertyGet resolves Drive property reads.
func (vm *VM) dispatchFSODrivePropertyGet(obj *fsoNativeObject, member string) Value {
	driveName := vm.fsoDriveNameFromPath(obj.drive)
	rootPath := vm.fsoDriveRootFromName(driveName)
	usage := vm.fsoDiskUsage(rootPath)

	switch {
	case strings.EqualFold(member, "DriveLetter"):
		return NewString(driveName)
	case strings.EqualFold(member, "Path"):
		if runtime.GOOS == "windows" {
			if strings.HasSuffix(driveName, ":") {
				return NewString(driveName)
			}
			return NewString(driveName + ":")
		}
		return NewString(rootPath)
	case strings.EqualFold(member, "DriveType"):
		return NewInteger(2)
	case strings.EqualFold(member, "SerialNumber"):
		return NewString(fsoGetDriveSerialNumber(rootPath))
	case strings.EqualFold(member, "ShareName"):
		return NewString(fsoGetDriveShareName(rootPath))
	case strings.EqualFold(member, "VolumeName"):
		return NewString(fsoGetDriveVolumeName(rootPath))
	case strings.EqualFold(member, "FileSystem"):
		if runtime.GOOS == "windows" {
			return NewString("NTFS")
		}
		return NewString("UnixFS")
	case strings.EqualFold(member, "IsReady"):
		_, err := os.Stat(rootPath)
		return NewBool(err == nil)
	case strings.EqualFold(member, "RootFolder"):
		return vm.newFSONativeObject(fsoKindFolder, rootPath, nil)
	case strings.EqualFold(member, "TotalSize"):
		if usage == nil {
			return NewInteger(0)
		}
		return NewInteger(int64(usage.Size()))
	case strings.EqualFold(member, "FreeSpace"):
		if usage == nil {
			return NewInteger(0)
		}
		return NewInteger(int64(usage.Free()))
	case strings.EqualFold(member, "AvailableSpace"):
		if usage == nil {
			return NewInteger(0)
		}
		return NewInteger(int64(usage.Available()))
	default:
		return Value{Type: VTEmpty}
	}
}

// dispatchFSODrivesCollectionPropertyGet resolves Drives collection property reads.
func (vm *VM) dispatchFSODrivesCollectionPropertyGet(_ *fsoNativeObject, member string) Value {
	if strings.EqualFold(member, "Count") {
		return NewInteger(int64(len(vm.fsoEnumerateDriveNames())))
	}
	return Value{Type: VTEmpty}
}

// fsoEnumerateDriveNames returns available drive identifiers for the active OS.
func (vm *VM) fsoEnumerateDriveNames() []string {
	if runtime.GOOS != "windows" {
		return []string{"C"}
	}
	drives := make([]string, 0, 8)
	for letter := 'A'; letter <= 'Z'; letter++ {
		candidate := fmt.Sprintf("%c:\\", letter)
		if _, err := os.Stat(candidate); err == nil {
			drives = append(drives, string(letter))
		}
	}
	if len(drives) == 0 {
		cwd, err := os.Getwd()
		if err == nil {
			vol := filepath.VolumeName(cwd)
			if len(vol) >= 1 {
				drives = append(drives, strings.TrimSuffix(vol, ":"))
			}
		}
	}
	sort.Strings(drives)
	return drives
}

// fsoDriveNameFromPath resolves a normalized drive name from any path input.
func (vm *VM) fsoDriveNameFromPath(path string) string {
	trimmed := strings.TrimSpace(path)
	if runtime.GOOS != "windows" {
		if strings.HasSuffix(trimmed, ":") && len(trimmed) == 2 {
			return strings.ToUpper(strings.TrimSuffix(trimmed, ":"))
		}
		if len(trimmed) >= 2 && trimmed[1] == ':' {
			return strings.ToUpper(trimmed[:1])
		}
		return "C"
	}

	if strings.HasSuffix(trimmed, ":") && len(trimmed) == 2 {
		return strings.ToUpper(strings.TrimSuffix(trimmed, ":"))
	}
	if len(trimmed) >= 2 && trimmed[1] == ':' {
		return strings.ToUpper(trimmed[:1])
	}
	if resolved, ok := vm.fsoResolvePath(trimmed); ok {
		vol := filepath.VolumeName(resolved)
		if len(vol) >= 1 {
			return strings.ToUpper(strings.TrimSuffix(vol, ":"))
		}
	}
	return ""
}

// fsoDriveRootFromName builds a root path from one normalized drive name.
func (vm *VM) fsoDriveRootFromName(driveName string) string {
	if runtime.GOOS != "windows" {
		return string(os.PathSeparator)
	}
	trimmed := strings.TrimSpace(strings.TrimSuffix(driveName, ":"))
	if trimmed == "" {
		return ""
	}
	if len(trimmed) == 1 {
		return strings.ToUpper(trimmed) + ":\\"
	}
	return trimmed
}

// fsoDiskUsage returns disk usage information for one root path.
func (vm *VM) fsoDiskUsage(rootPath string) *du.DiskUsage {
	if strings.TrimSpace(rootPath) == "" {
		return nil
	}
	if _, err := os.Stat(rootPath); err != nil {
		return nil
	}
	return du.NewDiskUsage(rootPath)
}

// dispatchFSOFileMethod handles FSOFile methods and property Let behavior.
func (vm *VM) dispatchFSOFileMethod(obj *fsoNativeObject, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "Delete"):
		_ = vm.fsoDeletePath(obj.path, false)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Move"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		destination, ok := vm.fsoResolvePath(args[0].String())
		if ok && vm.fsoMovePath(obj.path, destination) == nil {
			obj.path = destination
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Copy"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		destination, ok := vm.fsoResolvePath(args[0].String())
		if ok {
			_, _ = vm.fsoCopyFile(obj.path, destination, vm.fsoOverwriteDefault(args, 1, true))
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "OpenAsTextStream"):
		mode := 1
		if len(args) >= 1 {
			mode = vm.asInt(args[0])
		}
		stream := vm.fsoOpenTextStream(obj.path, mode, mode != 1)
		if stream == nil {
			return Value{Type: VTEmpty}
		}
		return vm.newFSONativeObject(fsoKindTextStream, obj.path, stream)
	case strings.EqualFold(member, "Name"):
		if len(args) < 1 {
			return vm.dispatchFSOFilePropertyGet(obj, "Name")
		}
		newName := strings.TrimSpace(args[0].String())
		if newName == "" {
			return Value{Type: VTEmpty}
		}
		parentPath := filepath.Dir(obj.path)
		destination := filepath.Join(parentPath, newName)
		if vm.fsoMovePath(obj.path, destination) == nil {
			obj.path = destination
		}
		return Value{Type: VTEmpty}
	default:
		return Value{Type: VTEmpty}
	}
}

// dispatchFSOFolderMethod handles FSOFolder methods and property Let behavior.
func (vm *VM) dispatchFSOFolderMethod(obj *fsoNativeObject, member string, args []Value) Value {
	switch {
	case strings.EqualFold(member, "Delete"):
		_ = vm.fsoDeletePath(obj.path, true)
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Move"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		destination, ok := vm.fsoResolvePath(args[0].String())
		if ok && vm.fsoMovePath(obj.path, destination) == nil {
			obj.path = destination
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Copy"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		destination, ok := vm.fsoResolvePath(args[0].String())
		if ok {
			_ = vm.fsoCopyFolder(obj.path, destination, vm.fsoOverwriteDefault(args, 1, true))
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "CreateTextFile"):
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		targetPath := filepath.Join(obj.path, args[0].String())
		resolvedPath, ok := vm.fsoResolvePath(targetPath)
		if !ok {
			return Value{Type: VTEmpty}
		}
		overwrite := vm.fsoOverwriteDefault(args, 1, true)
		flags := os.O_CREATE | os.O_WRONLY
		if overwrite {
			flags |= os.O_TRUNC
		} else {
			flags |= os.O_EXCL
		}
		fileHandle, err := os.OpenFile(resolvedPath, flags, 0644)
		if err != nil {
			return Value{Type: VTEmpty}
		}
		state := &fsoTextStreamState{file: fileHandle, mode: 2, atEnd: false, line: 1, column: 1}
		return vm.newFSONativeObject(fsoKindTextStream, resolvedPath, state)
	case strings.EqualFold(member, "Name"):
		if len(args) < 1 {
			return vm.dispatchFSOFolderPropertyGet(obj, "Name")
		}
		newName := strings.TrimSpace(args[0].String())
		if newName == "" {
			return Value{Type: VTEmpty}
		}
		parentPath := filepath.Dir(obj.path)
		destination := filepath.Join(parentPath, newName)
		if vm.fsoMovePath(obj.path, destination) == nil {
			obj.path = destination
		}
		return Value{Type: VTEmpty}
	default:
		return Value{Type: VTEmpty}
	}
}

// fsoDeletePath removes one file or folder tree and normalizes attributes first
// so Windows cleanup succeeds for test sandboxes and moved/copied paths.
func (vm *VM) fsoDeletePath(path string, recursive bool) error {
	info, err := os.Stat(path)
	if err != nil {
		if os.IsNotExist(err) {
			return nil
		}
		return err
	}

	vm.fsoReleasePathObjects(path)

	if !recursive || !info.IsDir() {
		return vm.fsoRemoveWithRetry(path, false)
	}

	_ = filepath.Walk(path, func(walkPath string, walkInfo os.FileInfo, walkErr error) error {
		if walkErr != nil {
			return nil
		}
		mode := os.FileMode(0666)
		if walkInfo.IsDir() {
			mode = 0777
		}
		_ = os.Chmod(walkPath, mode)
		return nil
	})

	return vm.fsoRemoveWithRetry(path, true)
}

// fsoReleasePathObjects closes and drops tracked FSO runtime objects rooted at
// one file or folder path so recursive deletes are not blocked by live streams.
func (vm *VM) fsoReleasePathObjects(path string) {
	cleanTarget := strings.ToLower(filepath.Clean(path))
	prefix := cleanTarget + string(os.PathSeparator)

	for objID, obj := range vm.fsoItems {
		if obj == nil || strings.TrimSpace(obj.path) == "" {
			continue
		}
		cleanObjectPath := strings.ToLower(filepath.Clean(obj.path))
		if cleanObjectPath != cleanTarget && !strings.HasPrefix(cleanObjectPath, prefix) {
			continue
		}
		if obj.stream != nil && obj.stream.file != nil {
			_ = obj.stream.file.Close()
			obj.stream.file = nil
			obj.stream.reader = nil
			obj.stream.atEnd = true
		}
		delete(vm.fsoItems, objID)
	}
}

// fsoRemoveWithRetry removes one path and retries briefly to absorb transient
// Windows sharing delays after recent stream close and rename operations.
func (vm *VM) fsoRemoveWithRetry(path string, recursive bool) error {
	var lastErr error
	for attempt := 0; attempt < 5; attempt++ {
		_ = os.Chmod(path, 0666)
		if recursive {
			lastErr = os.RemoveAll(path)
		} else {
			lastErr = os.Remove(path)
		}
		if lastErr == nil {
			if _, statErr := os.Stat(path); os.IsNotExist(statErr) {
				return nil
			}
		}
		time.Sleep(15 * time.Millisecond)
	}
	return lastErr
}

// fsoMovePath moves one file or folder path and replaces an existing destination
// when needed. If rename cannot complete directly, it falls back to copy+delete.
func (vm *VM) fsoMovePath(sourcePath, destinationPath string) error {
	if strings.TrimSpace(sourcePath) == "" || strings.TrimSpace(destinationPath) == "" {
		return os.ErrInvalid
	}
	if strings.EqualFold(filepath.Clean(sourcePath), filepath.Clean(destinationPath)) {
		return nil
	}
	if err := os.MkdirAll(filepath.Dir(destinationPath), 0755); err != nil {
		return err
	}
	if err := os.Rename(sourcePath, destinationPath); err == nil {
		return nil
	}

	if destInfo, err := os.Stat(destinationPath); err == nil {
		if delErr := vm.fsoDeletePath(destinationPath, destInfo.IsDir()); delErr == nil {
			if renameErr := os.Rename(sourcePath, destinationPath); renameErr == nil {
				return nil
			}
		}
	}

	sourceInfo, err := os.Stat(sourcePath)
	if err != nil {
		return err
	}
	if sourceInfo.IsDir() {
		if err := vm.fsoCopyFolder(sourcePath, destinationPath, true); err != nil {
			return err
		}
		return vm.fsoDeletePath(sourcePath, true)
	}
	if _, err := vm.fsoCopyFile(sourcePath, destinationPath, true); err != nil {
		return err
	}
	return vm.fsoDeletePath(sourcePath, false)
}

// dispatchFSOTextStreamMethod handles TextStream methods.
func (vm *VM) dispatchFSOTextStreamMethod(obj *fsoNativeObject, member string, args []Value) Value {
	if obj.stream == nil {
		return Value{Type: VTEmpty}
	}

	switch {
	case strings.EqualFold(member, "Close"):
		if obj.stream.file != nil && !obj.stream.standard {
			_ = obj.stream.file.Close()
		}
		obj.stream.file = nil
		obj.stream.reader = nil
		obj.stream.atEnd = true
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "ReadAll"):
		if obj.stream.file == nil {
			return NewString("")
		}
		if !obj.stream.standard {
			_, _ = obj.stream.file.Seek(0, io.SeekStart)
		}
		obj.stream.reader = nil
		content, err := io.ReadAll(obj.stream.file)
		if err != nil {
			return NewString("")
		}
		obj.stream.atEnd = true
		return NewString(string(content))
	case strings.EqualFold(member, "ReadLine"):
		if obj.stream.file == nil {
			return NewString("")
		}
		if obj.stream.reader == nil {
			obj.stream.reader = bufio.NewReader(obj.stream.file)
		}
		line, err := obj.stream.reader.ReadString('\n')
		if err == io.EOF {
			obj.stream.atEnd = true
			obj.stream.line++
			obj.stream.column = 1
			return NewString(strings.TrimRight(line, "\r\n"))
		}
		if err != nil {
			obj.stream.atEnd = true
			return NewString("")
		}
		obj.stream.line++
		obj.stream.column = 1
		return NewString(strings.TrimRight(line, "\r\n"))
	case strings.EqualFold(member, "Read"):
		if obj.stream.file == nil {
			return NewString("")
		}
		charCount := 1
		if len(args) >= 1 {
			charCount = vm.asInt(args[0])
		}
		if charCount < 0 {
			charCount = 0
		}
		buffer := make([]byte, charCount)
		readCount, err := obj.stream.file.Read(buffer)
		if err == io.EOF {
			obj.stream.atEnd = true
		}
		if readCount <= 0 {
			return NewString("")
		}
		obj.stream.column += readCount
		return NewString(string(buffer[:readCount]))
	case strings.EqualFold(member, "Skip"):
		if obj.stream.file == nil {
			return Value{Type: VTEmpty}
		}
		charCount := 1
		if len(args) >= 1 {
			charCount = vm.asInt(args[0])
		}
		if charCount < 0 {
			charCount = 0
		}
		if charCount == 0 {
			return Value{Type: VTEmpty}
		}
		buffer := make([]byte, charCount)
		readCount, err := obj.stream.file.Read(buffer)
		if err == io.EOF {
			obj.stream.atEnd = true
		}
		if readCount > 0 {
			obj.stream.column += readCount
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "SkipLine"):
		if obj.stream.file == nil {
			return Value{Type: VTEmpty}
		}
		if obj.stream.reader == nil {
			obj.stream.reader = bufio.NewReader(obj.stream.file)
		}
		line, err := obj.stream.reader.ReadString('\n')
		if err == io.EOF {
			obj.stream.atEnd = true
			if len(line) > 0 {
				obj.stream.line++
				obj.stream.column = 1
			}
			return Value{Type: VTEmpty}
		}
		if err != nil {
			obj.stream.atEnd = true
			return Value{Type: VTEmpty}
		}
		obj.stream.line++
		obj.stream.column = 1
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "Write"):
		if obj.stream.file == nil || obj.stream.mode == 1 {
			return Value{Type: VTEmpty}
		}
		if len(args) >= 1 {
			text := args[0].String()
			_, _ = obj.stream.file.WriteString(text)
			obj.stream.column += len(text)
		}
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "WriteLine"):
		if obj.stream.file == nil || obj.stream.mode == 1 {
			return Value{Type: VTEmpty}
		}
		text := ""
		if len(args) >= 1 {
			text = args[0].String()
		}
		_, _ = obj.stream.file.WriteString(text + "\r\n")
		obj.stream.line++
		obj.stream.column = 1
		return Value{Type: VTEmpty}
	case strings.EqualFold(member, "WriteBlankLines"):
		if obj.stream.file == nil || obj.stream.mode == 1 {
			return Value{Type: VTEmpty}
		}
		lineCount := 1
		if len(args) >= 1 {
			lineCount = vm.asInt(args[0])
		}
		if lineCount < 0 {
			lineCount = 0
		}
		for i := 0; i < lineCount; i++ {
			_, _ = obj.stream.file.WriteString("\r\n")
		}
		if lineCount > 0 {
			obj.stream.line += lineCount
			obj.stream.column = 1
		}
		return Value{Type: VTEmpty}
	default:
		return Value{Type: VTEmpty}
	}
}

// fsoGetStandardStream returns a TextStream wrapper for stdin, stdout, or stderr.
func (vm *VM) fsoGetStandardStream(streamType int) Value {
	var fileHandle *os.File
	mode := 1

	switch streamType {
	case 0:
		fileHandle = os.Stdin
		mode = 1
	case 1:
		fileHandle = os.Stdout
		mode = 2
	case 2:
		fileHandle = os.Stderr
		mode = 2
	default:
		vm.raise(vbscript.InvalidProcedureCallOrArgument, "Scripting.FileSystemObject.GetStandardStream requires stream type 0, 1, or 2")
		return Value{Type: VTEmpty}
	}

	state := &fsoTextStreamState{
		file:     fileHandle,
		mode:     mode,
		atEnd:    false,
		line:     1,
		column:   1,
		standard: true,
	}
	return vm.newFSONativeObject(fsoKindTextStream, "", state)
}

// dispatchFSOFilesCollectionMethod handles Files collection calls.
func (vm *VM) dispatchFSOFilesCollectionMethod(obj *fsoNativeObject, member string, args []Value) Value {
	if member == "" || strings.EqualFold(member, "Item") {
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		targetPath, ok := vm.fsoResolvePath(filepath.Join(obj.path, args[0].String()))
		if !ok {
			return Value{Type: VTEmpty}
		}
		info, err := os.Stat(targetPath)
		if err != nil || info.IsDir() {
			return Value{Type: VTEmpty}
		}
		return vm.newFSONativeObject(fsoKindFile, targetPath, nil)
	}
	return Value{Type: VTEmpty}
}

// dispatchFSOSubFoldersCollectionMethod handles SubFolders collection calls.
func (vm *VM) dispatchFSOSubFoldersCollectionMethod(obj *fsoNativeObject, member string, args []Value) Value {
	if member == "" || strings.EqualFold(member, "Item") {
		if len(args) < 1 {
			return Value{Type: VTEmpty}
		}
		targetPath, ok := vm.fsoResolvePath(filepath.Join(obj.path, args[0].String()))
		if !ok {
			return Value{Type: VTEmpty}
		}
		info, err := os.Stat(targetPath)
		if err != nil || !info.IsDir() {
			return Value{Type: VTEmpty}
		}
		return vm.newFSONativeObject(fsoKindFolder, targetPath, nil)
	}
	return Value{Type: VTEmpty}
}

// fsoResolveFileTimes resolves created/accessed/modified timestamps with a Windows-specific path and cross-platform fallback.
func fsoResolveFileTimes(info os.FileInfo) (time.Time, time.Time, time.Time) {
	modified := info.ModTime()
	created := modified
	accessed := modified
	if runtime.GOOS == "windows" {
		return fsoResolveWindowsFileTimes(info, modified)
	}
	return created, accessed, modified
}

// dispatchFSOFilePropertyGet resolves FSOFile property reads.
func (vm *VM) dispatchFSOFilePropertyGet(obj *fsoNativeObject, member string) Value {
	info, err := os.Stat(obj.path)
	if err != nil || info.IsDir() {
		return Value{Type: VTEmpty}
	}
	created, accessed, modified := fsoResolveFileTimes(info)

	switch {
	case strings.EqualFold(member, "Name"):
		return NewString(info.Name())
	case strings.EqualFold(member, "ShortName"):
		return NewString(fsoGetShortName(obj.path, info.Name()))
	case strings.EqualFold(member, "Path"):
		return NewString(obj.path)
	case strings.EqualFold(member, "ShortPath"):
		return NewString(fsoGetShortPath(obj.path))
	case strings.EqualFold(member, "Attributes"):
		return NewInteger(int64(vm.fsoAttributesForPath(obj.path, info)))
	case strings.EqualFold(member, "DateCreated"):
		return NewDate(created)
	case strings.EqualFold(member, "DateLastAccessed"):
		return NewDate(accessed)
	case strings.EqualFold(member, "DateLastModified"):
		return NewDate(modified)
	case strings.EqualFold(member, "Drive"):
		return vm.newFSODriveObject(vm.fsoDriveNameFromPath(obj.path))
	case strings.EqualFold(member, "Size"):
		return NewInteger(info.Size())
	case strings.EqualFold(member, "Type"):
		ext := strings.TrimPrefix(strings.ToLower(filepath.Ext(info.Name())), ".")
		if ext == "" {
			return NewString("File")
		}
		return NewString(strings.ToUpper(ext) + " File")
	case strings.EqualFold(member, "IsRootFolder"):
		return NewBool(false)
	case strings.EqualFold(member, "ParentFolder"):
		return vm.newFSONativeObject(fsoKindFolder, filepath.Dir(obj.path), nil)
	default:
		return Value{Type: VTEmpty}
	}
}

// dispatchFSOFolderPropertyGet resolves FSOFolder property reads.
func (vm *VM) dispatchFSOFolderPropertyGet(obj *fsoNativeObject, member string) Value {
	info, err := os.Stat(obj.path)
	if err != nil || !info.IsDir() {
		return Value{Type: VTEmpty}
	}
	created, accessed, modified := fsoResolveFileTimes(info)

	switch {
	case strings.EqualFold(member, "Name"):
		return NewString(info.Name())
	case strings.EqualFold(member, "ShortName"):
		name := info.Name()
		if name == "" {
			name = obj.path
		}
		return NewString(fsoGetShortName(obj.path, name))
	case strings.EqualFold(member, "Path"):
		return NewString(obj.path)
	case strings.EqualFold(member, "ShortPath"):
		return NewString(fsoGetShortPath(obj.path))
	case strings.EqualFold(member, "Attributes"):
		return NewInteger(int64(vm.fsoAttributesForPath(obj.path, info)))
	case strings.EqualFold(member, "DateCreated"):
		return NewDate(created)
	case strings.EqualFold(member, "DateLastAccessed"):
		return NewDate(accessed)
	case strings.EqualFold(member, "DateLastModified"):
		return NewDate(modified)
	case strings.EqualFold(member, "Drive"):
		return vm.newFSODriveObject(vm.fsoDriveNameFromPath(obj.path))
	case strings.EqualFold(member, "Size"):
		return NewInteger(info.Size())
	case strings.EqualFold(member, "Type"):
		return NewString("Folder")
	case strings.EqualFold(member, "IsRootFolder"):
		cleanPath := filepath.Clean(obj.path)
		rootPath := filepath.Clean(vm.fsoDriveRootFromName(vm.fsoDriveNameFromPath(obj.path)))
		if rootPath == "" {
			rootPath = string(os.PathSeparator)
		}
		return NewBool(strings.EqualFold(cleanPath, rootPath))
	case strings.EqualFold(member, "ParentFolder"):
		return vm.newFSONativeObject(fsoKindFolder, filepath.Dir(obj.path), nil)
	case strings.EqualFold(member, "Files"):
		return vm.newFSONativeObject(fsoKindFilesCollection, obj.path, nil)
	case strings.EqualFold(member, "SubFolders"):
		return vm.newFSONativeObject(fsoKindSubFoldersCollection, obj.path, nil)
	default:
		return Value{Type: VTEmpty}
	}
}

// dispatchFSOTextStreamPropertyGet resolves TextStream property reads.
func (vm *VM) dispatchFSOTextStreamPropertyGet(obj *fsoNativeObject, member string) Value {
	if obj.stream == nil {
		return Value{Type: VTEmpty}
	}

	switch {
	case strings.EqualFold(member, "ReadAll"):
		// VBScript compatibility: TextStream.ReadAll can be consumed in expression
		// context without explicit parentheses (e.g., value = ts.ReadAll).
		return vm.dispatchFSOTextStreamMethod(obj, "ReadAll", nil)
	case strings.EqualFold(member, "ReadLine"):
		// VBScript compatibility: TextStream.ReadLine can also be used without
		// parentheses in expression/member-get context.
		return vm.dispatchFSOTextStreamMethod(obj, "ReadLine", nil)
	case strings.EqualFold(member, "AtEndOfStream"):
		return NewBool(obj.stream.atEnd)
	case strings.EqualFold(member, "Line"):
		if obj.stream.line <= 0 {
			return NewInteger(1)
		}
		return NewInteger(int64(obj.stream.line))
	case strings.EqualFold(member, "Column"):
		if obj.stream.column <= 0 {
			return NewInteger(1)
		}
		return NewInteger(int64(obj.stream.column))
	default:
		return Value{Type: VTEmpty}
	}
}

// dispatchFSOFilesCollectionPropertyGet resolves Files collection property reads.
func (vm *VM) dispatchFSOFilesCollectionPropertyGet(obj *fsoNativeObject, member string) Value {
	if !strings.EqualFold(member, "Count") {
		return Value{Type: VTEmpty}
	}

	entries, err := os.ReadDir(obj.path)
	if err != nil {
		return NewInteger(0)
	}

	count := int64(0)
	for _, entry := range entries {
		if !entry.IsDir() {
			count++
		}
	}

	return NewInteger(count)
}

// dispatchFSOSubFoldersCollectionPropertyGet resolves SubFolders collection property reads.
func (vm *VM) dispatchFSOSubFoldersCollectionPropertyGet(obj *fsoNativeObject, member string) Value {
	if !strings.EqualFold(member, "Count") {
		return Value{Type: VTEmpty}
	}

	entries, err := os.ReadDir(obj.path)
	if err != nil {
		return NewInteger(0)
	}

	count := int64(0)
	for _, entry := range entries {
		if entry.IsDir() {
			count++
		}
	}

	return NewInteger(count)
}

// fsoOpenTextStream opens a text stream with classic mode semantics.
func (vm *VM) fsoOpenTextStream(path string, mode int, create bool) *fsoTextStreamState {
	flags := os.O_RDONLY
	if mode == 2 {
		flags = os.O_WRONLY | os.O_CREATE | os.O_TRUNC
	}
	if mode == 8 {
		flags = os.O_WRONLY | os.O_CREATE | os.O_APPEND
	}
	if create && mode == 1 {
		flags = os.O_CREATE | os.O_RDONLY
	}

	if mode == 2 || mode == 8 {
		_ = os.MkdirAll(filepath.Dir(path), 0755)
	}

	fileHandle, err := os.OpenFile(path, flags, 0644)
	if err != nil {
		return nil
	}

	return &fsoTextStreamState{file: fileHandle, mode: mode, atEnd: false, line: 1, column: 1}
}

// fsoAttributesForPath computes classic FSO attributes from stat metadata.
func (vm *VM) fsoAttributesForPath(path string, info os.FileInfo) int {
	attrs := 0
	if info.IsDir() {
		attrs |= 16
	} else {
		attrs |= 32
	}
	if info.Mode()&0222 == 0 {
		attrs |= 1
	}
	base := filepath.Base(path)
	if strings.HasPrefix(base, ".") && base != "." && base != ".." {
		attrs |= 2
	}
	return attrs
}

// fsoGetShortName returns the short name segment for one path.
func fsoGetShortName(path string, fallback string) string {
	shortPath := fsoGetShortPath(path)
	shortName := filepath.Base(shortPath)
	if strings.TrimSpace(shortName) == "" || shortName == "." || shortName == string(os.PathSeparator) {
		return fallback
	}
	return shortName
}

// fsoResolvePath resolves and validates paths inside the web root sandbox.
func (vm *VM) fsoResolvePath(path string) (string, bool) {
	trimmed := strings.TrimSpace(path)
	if trimmed == "" {
		return "", false
	}

	rootPath := vm.host.Server().MapPath("/")
	resolvedPath := trimmed
	if !filepath.IsAbs(trimmed) {
		resolvedPath = vm.host.Server().MapPath(trimmed)
	}

	absRoot, rootErr := filepath.Abs(rootPath)
	absResolved, resolvedErr := filepath.Abs(resolvedPath)
	if rootErr != nil || resolvedErr != nil {
		return "", false
	}

	cleanRoot := strings.ToLower(filepath.Clean(absRoot))
	cleanResolved := strings.ToLower(filepath.Clean(absResolved))

	if cleanResolved == cleanRoot {
		return absResolved, true
	}

	prefix := cleanRoot + string(os.PathSeparator)
	if !strings.HasPrefix(cleanResolved, prefix) {
		return "", false
	}

	return absResolved, true
}

// fsoOverwriteDefault resolves optional overwrite arguments for copy and create methods.
func (vm *VM) fsoOverwriteDefault(args []Value, index int, fallback bool) bool {
	if len(args) <= index {
		return fallback
	}
	return vm.asBool(args[index])
}

// fsoCopyFile copies one file with overwrite handling.
func (vm *VM) fsoCopyFile(sourcePath, destinationPath string, overwrite bool) (int64, error) {
	if !overwrite {
		if _, err := os.Stat(destinationPath); err == nil {
			return 0, os.ErrExist
		}
	}

	sourceFile, err := os.Open(sourcePath)
	if err != nil {
		return 0, err
	}
	defer sourceFile.Close()

	flags := os.O_CREATE | os.O_WRONLY
	if overwrite {
		flags |= os.O_TRUNC
	} else {
		flags |= os.O_EXCL
	}

	destinationFile, err := os.OpenFile(destinationPath, flags, 0644)
	if err != nil {
		return 0, err
	}
	defer destinationFile.Close()

	return io.Copy(destinationFile, sourceFile)
}

// fsoCopyFolder copies one folder recursively.
func (vm *VM) fsoCopyFolder(sourcePath, destinationPath string, overwrite bool) error {
	info, err := os.Stat(sourcePath)
	if err != nil {
		return err
	}
	if !info.IsDir() {
		return os.ErrInvalid
	}

	if _, err := os.Stat(destinationPath); err == nil {
		if !overwrite {
			return os.ErrExist
		}
	}

	if err := os.MkdirAll(destinationPath, 0755); err != nil {
		return err
	}

	entries, err := os.ReadDir(sourcePath)
	if err != nil {
		return err
	}

	for _, entry := range entries {
		sourceEntryPath := filepath.Join(sourcePath, entry.Name())
		destinationEntryPath := filepath.Join(destinationPath, entry.Name())

		if entry.IsDir() {
			if err := vm.fsoCopyFolder(sourceEntryPath, destinationEntryPath, overwrite); err != nil {
				return err
			}
			continue
		}

		if _, err := vm.fsoCopyFile(sourceEntryPath, destinationEntryPath, overwrite); err != nil {
			return err
		}
	}

	return nil
}
