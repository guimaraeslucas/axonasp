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
//Use go install github.com/josephspurrier/goversioninfo/cmd/goversioninfo@latest
//Then run "go generate" in the project root to embed version info into the executable
//You need to specify -64=false or -arm=true if you're trying to create an 32-bit or ARM windows binary, this is required by the new version of golang
//go:generate goversioninfo -icon=icon_service.ico -64=true
package main

import (
	"fmt"
	"os"
	"path"
	"path/filepath"
	"runtime"
	"strings"
	"sync"
	"sync/atomic"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/axonvm"
	"github.com/kardianos/service"
)

const (
	defaultServiceName        = "AxonASPServer"
	defaultServiceDisplayName = "G3pix AxonASP Server"
	defaultServiceDesc        = "AxonASP Service wrapper"
	defaultServiceExecPath    = "./axonasp-http"
)

// serviceSettings stores all [service] configuration values used by the wrapper.
type serviceSettings struct {
	Name        string
	DisplayName string
	Description string
	ExecPath    string
	EnvVars     []string
}

// program owns the service lifecycle and child process state.
type program struct {
	logger   service.Logger
	settings serviceSettings

	cmdMu      sync.Mutex
	cmd        *execCmdWrapper
	isStopping atomic.Bool
}

// Start begins the service runtime loop in a background goroutine.
func (p *program) Start(s service.Service) error {
	p.isStopping.Store(false)
	go p.run()
	return nil
}

// run launches and supervises the configured AxonASP child executable.
func (p *program) run() {
	childConfig, errCode, detail := buildChildConfig(p.settings)
	if errCode != 0 {
		p.logError(errCode, detail)
		exitWithCode(errCode)
	}

	cmd := buildOSCommand(childConfig.execPath, childConfig.env)
	cmd.Dir = childConfig.workDir

	p.cmdMu.Lock()
	p.cmd = &execCmdWrapper{Cmd: cmd}
	p.cmdMu.Unlock()

	if err := cmd.Start(); err != nil {
		p.logError(axonvm.ErrServiceStartProcessFailed, err.Error())
		exitWithCode(axonvm.ErrServiceStartProcessFailed)
	}

	p.logger.Info("AxonASP child server process started.")

	waitErr := cmd.Wait()

	p.cmdMu.Lock()
	p.cmd = nil
	p.cmdMu.Unlock()

	if p.isStopping.Load() {
		p.logger.Info("AxonASP child server process stopped by service request.")
		return
	}

	if waitErr != nil {
		p.logError(axonvm.ErrServiceChildExitedUnexpectedly, waitErr.Error())
	} else {
		p.logError(axonvm.ErrServiceChildExitedUnexpectedly, "AxonASP child server process exited unexpectedly")
	}

	exitWithCode(axonvm.ErrServiceChildExitedUnexpectedly)
}

// Stop requests child process termination for service shutdown.
func (p *program) Stop(s service.Service) error {
	p.isStopping.Store(true)

	p.cmdMu.Lock()
	cmd := p.cmd
	p.cmdMu.Unlock()

	if cmd == nil || cmd.Cmd == nil || cmd.Cmd.Process == nil {
		return nil
	}

	p.logger.Info("Stopping AxonASP child server process.")
	if err := stopOSCommand(cmd.Cmd); err != nil {
		p.logError(axonvm.ErrServiceStopProcessFailed, err.Error())
		return err
	}
	return nil
}

// childProcessConfig is the resolved process launch metadata.
type childProcessConfig struct {
	execPath string
	workDir  string
	env      []string
}

// buildChildConfig resolves executable and environment values from service settings.
func buildChildConfig(cfg serviceSettings) (childProcessConfig, axonvm.AxonASPErrorCode, string) {
	wrapperExe, err := os.Executable()
	if err != nil {
		return childProcessConfig{}, axonvm.ErrServiceResolveExecutablePathFailed, err.Error()
	}

	wrapperDir := filepath.Dir(wrapperExe)
	resolvedExec, resolveErr := resolveServiceExecutablePath(wrapperDir, cfg.ExecPath)
	if resolveErr != nil {
		return childProcessConfig{}, axonvm.ErrServiceResolveExecutablePathFailed, resolveErr.Error()
	}

	if _, statErr := os.Stat(resolvedExec); statErr != nil {
		return childProcessConfig{}, axonvm.ErrServiceExecutableNotFound, resolvedExec
	}

	env, envErr := mergeServiceEnvironment(cfg.EnvVars)
	if envErr != nil {
		return childProcessConfig{}, axonvm.ErrServiceInvalidEnvironmentVariable, envErr.Error()
	}

	return childProcessConfig{
		execPath: resolvedExec,
		workDir:  filepath.Dir(resolvedExec),
		env:      env,
	}, 0, ""
}

// loadServiceSettings loads [service] values with defaults from axonasp.toml.
func loadServiceSettings() (serviceSettings, bool) {
	v := axonconfig.NewViper()
	if strings.TrimSpace(v.ConfigFileUsed()) == "" {
		return serviceSettings{}, false
	}

	settings := serviceSettings{
		Name:        defaultServiceName,
		DisplayName: defaultServiceDisplayName,
		Description: defaultServiceDesc,
		ExecPath:    defaultServiceExecPath,
		EnvVars:     nil,
	}

	if value := strings.TrimSpace(v.GetString("service.service_name")); value != "" {
		settings.Name = value
	}
	if value := strings.TrimSpace(v.GetString("service.service_display_name")); value != "" {
		settings.DisplayName = value
	}
	if value := strings.TrimSpace(v.GetString("service.service_description")); value != "" {
		settings.Description = value
	}
	if value := strings.TrimSpace(v.GetString("service.service_executable_path")); value != "" {
		settings.ExecPath = value
	}
	if values := v.GetStringSlice("service.service_environment_variables"); len(values) > 0 {
		settings.EnvVars = values
	}

	return settings, true
}

// resolveServiceExecutablePath resolves an executable path relative to the wrapper directory.
func resolveServiceExecutablePath(wrapperDir string, configuredPath string) (string, error) {
	trimmed := strings.TrimSpace(configuredPath)
	if trimmed == "" {
		trimmed = defaultServiceExecPath
	}

	if !hasFileExtension(trimmed) {
		if runtime.GOOS == "windows" {
			trimmed += ".exe"
		}
	}

	if !filepath.IsAbs(trimmed) {
		trimmed = filepath.Join(wrapperDir, filepath.FromSlash(trimmed))
	}

	absPath, err := filepath.Abs(trimmed)
	if err != nil {
		return "", err
	}
	return filepath.Clean(absPath), nil
}

// hasFileExtension reports whether the provided path already has a file extension.
func hasFileExtension(p string) bool {
	base := path.Base(strings.ReplaceAll(p, "\\", "/"))
	return strings.Contains(base, ".")
}

// mergeServiceEnvironment validates service env entries and appends them to inherited env.
func mergeServiceEnvironment(extra []string) ([]string, error) {
	if len(extra) == 0 {
		return nil, nil
	}

	merged := make([]string, 0, len(extra)+len(os.Environ()))
	merged = append(merged, os.Environ()...)
	for _, item := range extra {
		entry := strings.TrimSpace(item)
		if entry == "" {
			continue
		}
		eq := strings.IndexByte(entry, '=')
		if eq <= 0 {
			return nil, fmt.Errorf("invalid environment variable %q", entry)
		}
		merged = append(merged, entry)
	}

	return merged, nil
}

// logError emits a consistent AxonASP error line to the service logger.
func (p *program) logError(code axonvm.AxonASPErrorCode, detail string) {
	message := fmt.Sprintf("Error %d: %s", code, code.String())
	if detail != "" {
		message += ": " + detail
	}
	p.logger.Error(message)
}

// exitWithCode exits the wrapper process with the given AxonASP error code.
func exitWithCode(code axonvm.AxonASPErrorCode) {
	os.Exit(int(code))
}

// defaultServiceConfig returns a fallback service manager configuration.
func defaultServiceConfig() serviceSettings {
	return serviceSettings{
		Name:        defaultServiceName,
		DisplayName: defaultServiceDisplayName,
		Description: defaultServiceDesc,
		ExecPath:    defaultServiceExecPath,
	}
}

// main initializes service configuration and executes wrapper commands.
func main() {
	settings := defaultServiceConfig()
	if loaded, ok := loadServiceSettings(); ok {
		settings = loaded
	} else {
		fmt.Printf("Warning: %s\n", axonvm.ErrViperReadConfigFailed.String())
	}

	wrapperExe, err := os.Executable()
	if err != nil {
		fmt.Printf("Error %d: %s: %v\n", axonvm.ErrServiceResolveExecutablePathFailed, axonvm.ErrServiceResolveExecutablePathFailed.String(), err)
		exitWithCode(axonvm.ErrServiceResolveExecutablePathFailed)
	}

	resolvedExecPath, resolveErr := resolveServiceExecutablePath(filepath.Dir(wrapperExe), settings.ExecPath)
	if resolveErr != nil {
		fmt.Printf("Error %d: %s: %v\n", axonvm.ErrServiceResolveExecutablePathFailed, axonvm.ErrServiceResolveExecutablePathFailed.String(), resolveErr)
		exitWithCode(axonvm.ErrServiceResolveExecutablePathFailed)
	}

	workingDir := filepath.Dir(resolvedExecPath)

	svcConfig := &service.Config{
		Name:        settings.Name,
		DisplayName: settings.DisplayName,
		Description: settings.Description,
		Option: service.KeyValue{
			"OnFailure":              "restart",
			"OnFailureDelayDuration": "1s",
			"WorkingDirectory":       workingDir,
		},
	}

	prg := &program{settings: settings}
	s, err := service.New(prg, svcConfig)
	if err != nil {
		fmt.Printf("Error %d: %s: %v\n", axonvm.ErrServiceCreateFailed, axonvm.ErrServiceCreateFailed.String(), err)
		exitWithCode(axonvm.ErrServiceCreateFailed)
	}

	logger, err := s.Logger(nil)
	if err != nil {
		fmt.Printf("Error %d: %s: %v\n", axonvm.ErrServiceLoggerCreateFailed, axonvm.ErrServiceLoggerCreateFailed.String(), err)
		exitWithCode(axonvm.ErrServiceLoggerCreateFailed)
	}
	prg.logger = logger

	if len(os.Args) > 1 {
		err = service.Control(s, os.Args[1])
		if err != nil {
			prg.logError(axonvm.ErrServiceControlCommandFailed, err.Error())
			exitWithCode(axonvm.ErrServiceControlCommandFailed)
		}
		return
	}

	err = s.Run()
	if err != nil {
		prg.logError(axonvm.ErrServiceRunFailed, err.Error())
		exitWithCode(axonvm.ErrServiceRunFailed)
	}
}
