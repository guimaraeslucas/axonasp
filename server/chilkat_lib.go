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
	"net/http"
	"strings"
	"time"
)

// ChilkatGlobal mimics the Chilkat_9_5_0.Global object.
// It accepts any unlock key and always reports "unlocked".
type ChilkatGlobal struct {
	ctx *ExecutionContext
}

func NewChilkatGlobal(ctx *ExecutionContext) *ChilkatGlobal {
	return &ChilkatGlobal{ctx: ctx}
}

func (g *ChilkatGlobal) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "unlockstatus":
		return 1
	}
	return nil
}

func (g *ChilkatGlobal) SetProperty(name string, value interface{}) {
	// No-op — accept any property silently
}

func (g *ChilkatGlobal) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "unlockbundle":
		return true, nil
	}
	return nil, nil
}

// ChilkatHttp mimics the Chilkat_9_5_0.Http object.
// Uses Go's net/http client to perform HTTP requests.
type ChilkatHttp struct {
	ctx               *ExecutionContext
	lastMethodSuccess int
	lastErrorText     string
	timeout           time.Duration
}

func NewChilkatHttp(ctx *ExecutionContext) *ChilkatHttp {
	return &ChilkatHttp{
		ctx:               ctx,
		lastMethodSuccess: 1,
		timeout:           30 * time.Second,
	}
}

func (h *ChilkatHttp) GetProperty(name string) interface{} {
	switch strings.ToLower(name) {
	case "lastmethodsuccess":
		return h.lastMethodSuccess
	case "lasterrortext":
		return h.lastErrorText
	}
	return nil
}

func (h *ChilkatHttp) SetProperty(name string, value interface{}) {
	switch strings.ToLower(name) {
	case "connecttimeout":
		h.timeout = time.Duration(toInt(value)) * time.Second
	}
}

func (h *ChilkatHttp) CallMethod(name string, args ...interface{}) (interface{}, error) {
	switch strings.ToLower(name) {
	case "quickgetstr":
		if len(args) < 1 {
			h.lastMethodSuccess = 0
			h.lastErrorText = "QuickGetStr requires a URL argument"
			return "", nil
		}
		url := fmt.Sprintf("%v", args[0])
		return h.quickGetStr(url), nil
	}
	return nil, nil
}

func (h *ChilkatHttp) quickGetStr(url string) string {
	client := &http.Client{Timeout: h.timeout}
	req, err := http.NewRequest(http.MethodGet, url, nil)
	if err != nil {
		h.lastMethodSuccess = 0
		h.lastErrorText = err.Error()
		return ""
	}

	req.Header.Set("User-Agent", "Mozilla/5.0 (Windows NT 11.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/145.0.0.0 Safari/537.36 AxonASPServer/1.0")
	req.Header.Set("Accept", "*/*")

	resp, err := client.Do(req)
	if err != nil {
		h.lastMethodSuccess = 0
		h.lastErrorText = err.Error()
		return ""
	}
	defer resp.Body.Close()

	data, err := io.ReadAll(resp.Body)
	if err != nil {
		h.lastMethodSuccess = 0
		h.lastErrorText = err.Error()
		return ""
	}

	if resp.StatusCode >= 400 {
		h.lastMethodSuccess = 0
		h.lastErrorText = fmt.Sprintf("HTTP %d: %s", resp.StatusCode, resp.Status)
		return string(data)
	}

	h.lastMethodSuccess = 1
	h.lastErrorText = ""
	return decodeResponseText(data, resp.Header.Get("Content-Type"))
}
