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
// Package axonconfig provides configuration management for AxonASP Server using Viper and also the AboutG3pixAxonASP function to provide information about the project.
package axonconfig

import (
	"os"
	"path/filepath"
	"strings"
	"sync"

	"github.com/fsnotify/fsnotify"
	"github.com/spf13/viper"
)

var (
	customConfigPath string
	activeViper      *viper.Viper
	activeViperMu    sync.Mutex
)

// SetCustomConfigPath sets a custom configuration file path to use instead of the defaults.
func SetCustomConfigPath(path string) {
	activeViperMu.Lock()
	defer activeViperMu.Unlock()
	customConfigPath = path
	activeViper = nil // Invalidate singleton to reload from new path on next NewViper call
}

// NewViper returns the active globally shared Viper instance. If not initialized, it loads it.
func NewViper() *viper.Viper {
	activeViperMu.Lock()
	defer activeViperMu.Unlock()

	if activeViper == nil {
		activeViper = initViper()
	}
	return activeViper
}

func initViper() *viper.Viper {
	v := viper.New()
	v.SetConfigType("toml")

	if customConfigPath != "" {
		v.SetConfigFile(customConfigPath)
		_ = v.ReadInConfig()
	} else {
		configCandidates := []string{
			filepath.Join("config", "axonasp.toml"),
			filepath.Join("..", "config", "axonasp.toml"),
			filepath.Join("..", "..", "config", "axonasp.toml"),
		}
		if executablePath, err := os.Executable(); err == nil {
			configCandidates = append(configCandidates, filepath.Join(filepath.Dir(executablePath), "config", "axonasp.toml"))
		}

		for _, candidate := range configCandidates {
			v.SetConfigFile(candidate)
			if err := v.ReadInConfig(); err == nil {
				break
			}
		}
	}

	if v.GetBool("global.viper_automatic_env") {
		v.SetEnvKeyReplacer(strings.NewReplacer(".", "_"))
		v.AutomaticEnv()
	}

	return v
}

// EnableWatchIfConfigured enables Viper file watching only when global.viper_watch_config is true.
func EnableWatchIfConfigured(v *viper.Viper, onChange func(fsnotify.Event)) bool {
	if v == nil || !v.GetBool("global.viper_watch_config") {
		return false
	}

	if onChange != nil {
		v.OnConfigChange(onChange)
	}
	v.WatchConfig()
	return true
}

// AboutG3pixAxonASP returns a string with information about AxonASP, including its license and copyright. It is intended to be displayed in the "About" section of applications using AxonASP, that's the reason it is not in the main package, but in the axonconfig package, so it can be imported and used by other packages.
func AboutG3pixAxonASP() string {
	return `
G3pix ❖ AxonASP
────────────────────────────

The high-performance, cross-platform engine driving VBScript and JavaScript
into the next era. Built on a zero-allocation VM for Web, FastCGI, and CLI,
it bridges core logic with modern APIs across all systems. High-powered,
open-source, and ready for the future.

Contributing to AxonASP:
As an open-source project, AxonASP relies on community support to remain 
active. Maintenance and infrastructure costs are currently funded personally 
by the lead developer. Your support—whether through code contributions, bug 
reports, security patches, or financial donations—is vital to our continued 
growth. Thank you for being part of our journey.

License: 
Mozilla Public License, v. 2.0.

Copyright (C) 2026 G3pix Ltda. All rights reserved.
https://g3pix.com.br/axonasp


`
}
