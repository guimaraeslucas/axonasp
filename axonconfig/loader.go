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
package axonconfig

import (
	"os"
	"path/filepath"
	"strings"

	"github.com/fsnotify/fsnotify"
	"github.com/spf13/viper"
)

// NewViper loads axonasp.toml using cwd-relative and executable-relative fallbacks.
func NewViper() *viper.Viper {
	v := viper.New()
	v.SetConfigType("toml")

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
