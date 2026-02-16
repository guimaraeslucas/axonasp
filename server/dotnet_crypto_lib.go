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
	"crypto/md5"
	"fmt"
	"strings"
)

// DotNetMD5CryptoServiceProvider simulates .NET MD5CryptoServiceProvider.
type DotNetMD5CryptoServiceProvider struct {
	lastHash []byte
}

func NewDotNetMD5CryptoServiceProvider() *DotNetMD5CryptoServiceProvider {
	return &DotNetMD5CryptoServiceProvider{lastHash: nil}
}

func (m *DotNetMD5CryptoServiceProvider) GetProperty(name string) interface{} {
	switch lower := strings.ToLower(name); lower {
	case "hash":
		if len(m.lastHash) == 0 {
			return NewVBArrayFromValues(0, []interface{}{})
		}
		return bytesToVBArray(m.lastHash)
	case "hashsize":
		return 128
	case "canreusetransform":
		return true
	case "cantransformmultipleblocks":
		return true
	}
	return nil
}

func (m *DotNetMD5CryptoServiceProvider) SetProperty(name string, value interface{}) {
}

func (m *DotNetMD5CryptoServiceProvider) CallMethod(name string, args ...interface{}) interface{} {
	switch lower := strings.ToLower(name); lower {
	case "computehash":
		if len(args) < 1 {
			return NewVBArrayFromValues(0, []interface{}{})
		}
		data := m.normalizeInput(args[0])
		hash := md5.Sum(data)
		m.lastHash = hash[:]
		return bytesToVBArray(m.lastHash)
	case "initialize", "clear", "dispose":
		m.lastHash = nil
		return nil
	}
	return nil
}

func (m *DotNetMD5CryptoServiceProvider) normalizeInput(value interface{}) []byte {
	if value == nil {
		return []byte{}
	}
	if vbArr, ok := toVBArray(value); ok {
		buf := make([]byte, 0, len(vbArr.Values))
		for _, item := range vbArr.Values {
			byteVal := toInt(item)
			if byteVal < 0 {
				byteVal = 0
			}
			if byteVal > 255 {
				byteVal = 255
			}
			buf = append(buf, byte(byteVal))
		}
		return buf
	}
	if data, ok := value.([]byte); ok {
		return data
	}
	return []byte(fmt.Sprintf("%v", value))
}
