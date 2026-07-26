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
	"os"
	"sync"
	"sync/atomic"
	"time"
)

// fsoCacheItem stores cached metadata for one file or directory.
type fsoCacheItem struct {
	stat    os.FileInfo
	entries []os.DirEntry
	expires time.Time
}

// fsoCacheManager handles thread-safe caching of FSO metadata with TTL.
type fsoCacheManager struct {
	mu    sync.RWMutex
	items map[string]*fsoCacheItem

	// pathCache memoizes resolved absolute paths to avoid redundant filepath.Abs calls.
	pathCache sync.Map

	// inflightMu and inflight map implement a simple Singleflight-like mechanism.
	// This ensures that concurrent requests for the same path only trigger one disk I/O.
	inflightMu sync.Mutex
	inflight   map[string]*sync.WaitGroup

	// cacheDisabled, when true, causes GetStat and GetReadDir to always read
	// directly from disk, bypassing the TTL cache. This is used by trusted
	// desktop hosts (e.g. AxonHTA) that need real-time filesystem access.
	cacheDisabled atomic.Bool
}

// globalFSOCache provides process-wide caching for FSO operations.
var globalFSOCache = &fsoCacheManager{
	items:    make(map[string]*fsoCacheItem),
	inflight: make(map[string]*sync.WaitGroup),
}

// SetFSOCacheDisabled toggles the FSO metadata cache. When disabled, every
// GetStat/GetReadDir call reads directly from disk. Trusted desktop hosts
// (AxonHTA) call this at startup so that FSO operations always reflect the
// current filesystem state.
func SetFSOCacheDisabled(disabled bool) {
	globalFSOCache.cacheDisabled.Store(disabled)
}

// GetStat returns cached FileInfo for a path or performs os.Stat and caches it.
func (c *fsoCacheManager) GetStat(path string) (os.FileInfo, error) {
	if c.cacheDisabled.Load() {
		return os.Stat(path)
	}

	c.mu.RLock()
	item, ok := c.items[path]
	c.mu.RUnlock()

	now := time.Now()
	if ok && item.stat != nil && now.Before(item.expires) {
		return item.stat, nil
	}

	// Simple Singleflight for Stat
	c.inflightMu.Lock()
	wg, inProgress := c.inflight["stat:"+path]
	if inProgress {
		c.inflightMu.Unlock()
		wg.Wait()
		return c.GetStat(path) // Recurse to get the value cached by the winner.
	}
	wg = &sync.WaitGroup{}
	wg.Add(1)
	c.inflight["stat:"+path] = wg
	c.inflightMu.Unlock()

	defer func() {
		c.inflightMu.Lock()
		delete(c.inflight, "stat:"+path)
		c.inflightMu.Unlock()
		wg.Done()
	}()

	info, err := os.Stat(path)
	if err != nil {
		return nil, err
	}

	c.mu.Lock()
	defer c.mu.Unlock()

	// Passive cleanup during writes
	if len(c.items) > 1024 {
		for p, it := range c.items {
			if now.After(it.expires) {
				delete(c.items, p)
			}
		}
	}

	if item, ok = c.items[path]; ok {
		item.stat = info
		item.expires = now.Add(5 * time.Second) // Shorter TTL for burst absorption
	} else {
		c.items[path] = &fsoCacheItem{
			stat:    info,
			expires: now.Add(5 * time.Second),
		}
	}

	return info, nil
}

// GetReadDir returns cached directory entries for a path or performs os.ReadDir and caches it.
func (c *fsoCacheManager) GetReadDir(path string) ([]os.DirEntry, error) {
	if c.cacheDisabled.Load() {
		return os.ReadDir(path)
	}

	c.mu.RLock()
	item, ok := c.items[path]
	c.mu.RUnlock()

	now := time.Now()
	if ok && item.entries != nil && now.Before(item.expires) {
		return item.entries, nil
	}

	// Singleflight: Only allow one os.ReadDir for this path at a time.
	c.inflightMu.Lock()
	wg, inProgress := c.inflight["readdir:"+path]
	if inProgress {
		c.inflightMu.Unlock()
		wg.Wait()
		return c.GetReadDir(path)
	}
	wg = &sync.WaitGroup{}
	wg.Add(1)
	c.inflight["readdir:"+path] = wg
	c.inflightMu.Unlock()

	defer func() {
		c.inflightMu.Lock()
		delete(c.inflight, "readdir:"+path)
		c.inflightMu.Unlock()
		wg.Done()
	}()

	entries, err := os.ReadDir(path)
	if err != nil {
		return nil, err
	}

	c.mu.Lock()
	defer c.mu.Unlock()

	// Passive cleanup during writes
	if len(c.items) > 1024 {
		for p, it := range c.items {
			if now.After(it.expires) {
				delete(c.items, p)
			}
		}
	}

	if item, ok = c.items[path]; ok {
		item.entries = entries
		item.expires = now.Add(5 * time.Second) // 5s TTL for micro-burst absorption
	} else {
		c.items[path] = &fsoCacheItem{
			entries: entries,
			expires: now.Add(5 * time.Second),
		}
	}

	return entries, nil
}
