//go:build windows

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
	"encoding/binary"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"syscall"
	"time"
	"unsafe"

	"golang.org/x/sys/windows"
)

// fsoResolveWindowsFileTimes resolves Windows creation/access/write file timestamps.
func fsoResolveWindowsFileTimes(info os.FileInfo, fallback time.Time) (time.Time, time.Time, time.Time) {
	createdTime := fallback
	accessedTime := fallback
	modifiedTime := fallback
	if stat, ok := info.Sys().(*syscall.Win32FileAttributeData); ok {
		createdTime = time.Unix(0, stat.CreationTime.Nanoseconds())
		accessedTime = time.Unix(0, stat.LastAccessTime.Nanoseconds())
		modifiedTime = time.Unix(0, stat.LastWriteTime.Nanoseconds())
	}
	return createdTime, accessedTime, modifiedTime
}

// fsoGetFileVersion reads version resource data using version.dll APIs.
func fsoGetFileVersion(path string) string {
	const fallback = "1.0.0.0"
	absPath, err := filepath.Abs(path)
	if err != nil {
		absPath = path
	}

	versionDLL := windows.NewLazySystemDLL("version.dll")
	procGetFileVersionInfoSizeW := versionDLL.NewProc("GetFileVersionInfoSizeW")
	procGetFileVersionInfoW := versionDLL.NewProc("GetFileVersionInfoW")
	procVerQueryValueW := versionDLL.NewProc("VerQueryValueW")

	filePtr, err := windows.UTF16PtrFromString(absPath)
	if err != nil {
		return fallback
	}

	sizeRet, _, _ := procGetFileVersionInfoSizeW.Call(uintptr(unsafe.Pointer(filePtr)), 0)
	size := uint32(sizeRet)
	if size == 0 {
		return fallback
	}

	data := make([]byte, size)
	okRet, _, _ := procGetFileVersionInfoW.Call(
		uintptr(unsafe.Pointer(filePtr)),
		0,
		uintptr(size),
		uintptr(unsafe.Pointer(&data[0])),
	)
	if okRet == 0 {
		return fallback
	}

	queryPtr, err := windows.UTF16PtrFromString("\\")
	if err != nil {
		return fallback
	}

	var blockPtr uintptr
	var blockLen uint32
	okRet, _, _ = procVerQueryValueW.Call(
		uintptr(unsafe.Pointer(&data[0])),
		uintptr(unsafe.Pointer(queryPtr)),
		uintptr(unsafe.Pointer(&blockPtr)),
		uintptr(unsafe.Pointer(&blockLen)),
	)
	const fixedInfoSize = 52
	if okRet == 0 || blockPtr == 0 || blockLen < fixedInfoSize {
		return fallback
	}
	basePtr := uintptr(unsafe.Pointer(&data[0]))
	if blockPtr < basePtr {
		return fallback
	}
	offset := blockPtr - basePtr
	if offset+fixedInfoSize > uintptr(len(data)) {
		return fallback
	}
	fixed := data[offset : offset+fixedInfoSize]
	signature := binary.LittleEndian.Uint32(fixed[0:4])
	if signature != 0xFEEF04BD {
		return fallback
	}
	fileVersionMS := binary.LittleEndian.Uint32(fixed[8:12])
	fileVersionLS := binary.LittleEndian.Uint32(fixed[12:16])
	major := uint16(fileVersionMS >> 16)
	minor := uint16(fileVersionMS & 0xFFFF)
	build := uint16(fileVersionLS >> 16)
	revision := uint16(fileVersionLS & 0xFFFF)
	return fmt.Sprintf("%d.%d.%d.%d", major, minor, build, revision)
}

// fsoGetShortPath resolves one DOS 8.3 path when available.
func fsoGetShortPath(path string) string {
	absPath, err := filepath.Abs(path)
	if err != nil {
		absPath = path
	}

	kernel32 := windows.NewLazySystemDLL("kernel32.dll")
	procGetShortPathNameW := kernel32.NewProc("GetShortPathNameW")

	inputPtr, err := windows.UTF16PtrFromString(absPath)
	if err != nil {
		return absPath
	}

	buf := make([]uint16, 1024)
	ret, _, _ := procGetShortPathNameW.Call(
		uintptr(unsafe.Pointer(inputPtr)),
		uintptr(unsafe.Pointer(&buf[0])),
		uintptr(len(buf)),
	)
	if ret == 0 {
		return absPath
	}

	return windows.UTF16ToString(buf[:ret])
}

// fsoGetDriveSerialNumber resolves the volume serial number when available.
func fsoGetDriveSerialNumber(rootPath string) string {
	_, serial, _, _, ok := fsoGetVolumeInformation(rootPath)
	if !ok {
		return "00001"
	}
	return fmt.Sprintf("%d", serial)
}

// fsoGetDriveShareName returns a synthetic UNC-like share name for compatibility.
func fsoGetDriveShareName(rootPath string) string {
	trimmed := strings.TrimSpace(rootPath)
	if trimmed == "" {
		return "\\AxonASP"
	}
	trimmed = strings.TrimRight(trimmed, "\\/")
	if len(trimmed) >= 2 && trimmed[1] == ':' {
		return "\\\\" + strings.TrimSuffix(trimmed, ":")
	}
	return "\\AxonASP"
}

// fsoGetDriveVolumeName resolves the volume label when available.
func fsoGetDriveVolumeName(rootPath string) string {
	name, _, _, _, ok := fsoGetVolumeInformation(rootPath)
	if !ok || strings.TrimSpace(name) == "" {
		return "AxonASPVolume"
	}
	return name
}

// fsoGetVolumeInformation wraps GetVolumeInformationW in a compact helper.
func fsoGetVolumeInformation(rootPath string) (string, uint32, uint32, uint32, bool) {
	kernel32 := windows.NewLazySystemDLL("kernel32.dll")
	procGetVolumeInformationW := kernel32.NewProc("GetVolumeInformationW")

	root := strings.TrimSpace(rootPath)
	if root == "" {
		root = `C:\\`
	}
	if !strings.HasSuffix(root, "\\") {
		root += "\\"
	}

	rootPtr, err := windows.UTF16PtrFromString(root)
	if err != nil {
		return "", 0, 0, 0, false
	}

	volumeName := make([]uint16, windows.MAX_PATH+1)
	fileSystemName := make([]uint16, windows.MAX_PATH+1)
	var serialNumber uint32
	var maxComponentLen uint32
	var fileSystemFlags uint32

	ret, _, _ := procGetVolumeInformationW.Call(
		uintptr(unsafe.Pointer(rootPtr)),
		uintptr(unsafe.Pointer(&volumeName[0])),
		uintptr(len(volumeName)),
		uintptr(unsafe.Pointer(&serialNumber)),
		uintptr(unsafe.Pointer(&maxComponentLen)),
		uintptr(unsafe.Pointer(&fileSystemFlags)),
		uintptr(unsafe.Pointer(&fileSystemName[0])),
		uintptr(len(fileSystemName)),
	)
	if ret == 0 {
		return "", 0, 0, 0, false
	}

	return windows.UTF16ToString(volumeName), serialNumber, maxComponentLen, fileSystemFlags, true
}
