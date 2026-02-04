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
package server

import (
	"fmt"
	"runtime"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

type comCallKind int

const (
	comCallGetProperty comCallKind = iota
	comCallSetProperty
	comCallCallMethod
	comCallRelease
)

type comCall struct {
	kind     comCallKind
	dispatch *ole.IDispatch
	name     string
	args     []interface{}
	resp     chan comResult
}

type comResult struct {
	value interface{}
	err   error
}

type comHost struct {
	calls chan comCall
}

type COMObject struct {
	progID   string
	host     *comHost
	dispatch *ole.IDispatch
}

// NewCOMObject creates a COM object wrapper for Windows environments.
func NewCOMObject(progID string) (*COMObject, error) {
	host, dispatch, err := newCOMHost(progID)
	fmt.Printf("New COMObject %s\n", progID)
	if err != nil {
		return nil, err
	}

	obj := &COMObject{progID: progID, host: host, dispatch: dispatch}
	runtime.SetFinalizer(obj, func(o *COMObject) {
		o.release()
	})
	return obj, nil
}

func newCOMHost(progID string) (*comHost, *ole.IDispatch, error) {
	host := &comHost{calls: make(chan comCall)}
	initCh := make(chan comResult, 1)

	go func() {
		runtime.LockOSThread()
		defer runtime.UnlockOSThread()

		if err := ole.CoInitialize(0); err != nil {
			initCh <- comResult{err: err}
			return
		}
		defer ole.CoUninitialize()

		unknown, err := oleutil.CreateObject(progID)
		if err != nil {
			initCh <- comResult{err: err}
			return
		}

		dispatch, err := unknown.QueryInterface(ole.IID_IDispatch)
		unknown.Release()
		if err != nil {
			initCh <- comResult{err: err}
			return
		}

		initCh <- comResult{value: dispatch}

		for call := range host.calls {
			result := comResult{}
			switch call.kind {
			case comCallGetProperty:
				variant, err := oleutil.GetProperty(call.dispatch, call.name)
				if err != nil {
					result.err = err
					break
				}
				value, shouldClear := normalizeCOMVariant(variant, host)
				if shouldClear {
					variant.Clear()
				}
				result.value = value
			case comCallSetProperty:
				args := normalizeCOMArgs(call.args)
				_, err := oleutil.PutProperty(call.dispatch, call.name, args...)
				result.err = err
			case comCallCallMethod:
				args := normalizeCOMArgs(call.args)
				variant, err := oleutil.CallMethod(call.dispatch, call.name, args...)
				if err != nil {
					result.err = err
					break
				}
				value, shouldClear := normalizeCOMVariant(variant, host)
				if shouldClear {
					variant.Clear()
				}
				result.value = value
			case comCallRelease:
				if call.dispatch != nil {
					call.dispatch.Release()
				}
			}
			call.resp <- result
		}
	}()

	initResult := <-initCh
	if initResult.err != nil {
		return nil, nil, initResult.err
	}

	dispatch, ok := initResult.value.(*ole.IDispatch)
	if !ok || dispatch == nil {
		return nil, nil, fmt.Errorf("COM initialization failed for %s", progID)
	}

	return host, dispatch, nil
}

func (c *COMObject) GetProperty(name string) interface{} {
	result, err := c.host.call(comCallGetProperty, c.dispatch, name, nil)
	if err != nil {
		return nil
	}
	return result
}

func (c *COMObject) SetProperty(name string, value interface{}) error {
	_, err := c.host.call(comCallSetProperty, c.dispatch, name, []interface{}{value})
	return err
}

func (c *COMObject) CallMethod(name string, args ...interface{}) (interface{}, error) {
	return c.host.call(comCallCallMethod, c.dispatch, name, args)
}

func (c *COMObject) Enumerate() ([]interface{}, error) {
	countValue := c.GetProperty("Count")
	count := toInt(countValue)
	if count < 0 {
		count = 0
	}

	items := make([]interface{}, 0, count)
	if count == 0 {
		return items, nil
	}

	for i := 0; i < count; i++ {
		item, err := c.CallMethod("Item", i)
		if err != nil {
			return c.enumerateFromOne(count)
		}
		items = append(items, item)
	}

	return items, nil
}

func (c *COMObject) enumerateFromOne(count int) ([]interface{}, error) {
	items := make([]interface{}, 0, count)
	for i := 1; i <= count; i++ {
		item, err := c.CallMethod("Item", i)
		if err != nil {
			return nil, err
		}
		items = append(items, item)
	}
	return items, nil
}

func (c *COMObject) release() {
	if c == nil || c.host == nil || c.dispatch == nil {
		return
	}
	_, _ = c.host.call(comCallRelease, c.dispatch, "", nil)
	c.dispatch = nil
}

func (h *comHost) call(kind comCallKind, dispatch *ole.IDispatch, name string, args []interface{}) (interface{}, error) {
	if h == nil {
		return nil, fmt.Errorf("COM host is not available")
	}
	resp := make(chan comResult, 1)
	h.calls <- comCall{kind: kind, dispatch: dispatch, name: name, args: args, resp: resp}
	result := <-resp
	return result.value, result.err
}

func normalizeCOMVariant(variant *ole.VARIANT, host *comHost) (interface{}, bool) {
	if variant == nil {
		return nil, true
	}

	value := variant.Value()
	if value == nil {
		return nil, true
	}

	switch v := value.(type) {
	case *ole.IDispatch:
		return newCOMObjectFromDispatch(host, v), false
	case *ole.IUnknown:
		dispatch, err := v.QueryInterface(ole.IID_IDispatch)
		if err != nil {
			return nil, true
		}
		return newCOMObjectFromDispatch(host, dispatch), false
	case []interface{}:
		items := make([]interface{}, len(v))
		for i, item := range v {
			items[i] = normalizeCOMValue(item, host)
		}
		return NewVBArrayFromValues(0, items), true
	default:
		return value, true
	}
}

func normalizeCOMValue(value interface{}, host *comHost) interface{} {
	switch v := value.(type) {
	case *ole.IDispatch:
		return newCOMObjectFromDispatch(host, v)
	case *ole.IUnknown:
		dispatch, err := v.QueryInterface(ole.IID_IDispatch)
		if err != nil {
			return nil
		}
		return newCOMObjectFromDispatch(host, dispatch)
	case []interface{}:
		items := make([]interface{}, len(v))
		for i, item := range v {
			items[i] = normalizeCOMValue(item, host)
		}
		return NewVBArrayFromValues(0, items)
	default:
		return value
	}
}

func normalizeCOMArgs(args []interface{}) []interface{} {
	normalized := make([]interface{}, 0, len(args))
	for _, arg := range args {
		switch v := arg.(type) {
		case *COMObject:
			if v != nil && v.dispatch != nil {
				normalized = append(normalized, v.dispatch)
			} else {
				normalized = append(normalized, nil)
			}
		case *ADODBConnection:
			if v == nil {
				normalized = append(normalized, nil)
				continue
			}
			if v.oleConnection != nil {
				normalized = append(normalized, v.oleConnection)
				continue
			}
			if v.ConnectionString != "" {
				normalized = append(normalized, v.ConnectionString)
				continue
			}
			normalized = append(normalized, nil)
		case *VBArray:
			if v == nil {
				normalized = append(normalized, nil)
				continue
			}
			normalized = append(normalized, v.Values)
		default:
			normalized = append(normalized, arg)
		}
	}
	return normalized
}

func newCOMObjectFromDispatch(host *comHost, dispatch *ole.IDispatch) *COMObject {
	if host == nil || dispatch == nil {
		return nil
	}
	obj := &COMObject{progID: "", host: host, dispatch: dispatch}
	runtime.SetFinalizer(obj, func(o *COMObject) {
		o.release()
	})
	return obj
}
