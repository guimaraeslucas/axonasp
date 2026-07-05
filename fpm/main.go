//go:build !windows

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

package main

import (
	"context"
	"errors"
	"fmt"
	"log"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
	"strings"
	"sync"
	"syscall"
	"time"

	"github.com/pelletier/go-toml/v2"
)

type PoolConfig struct {
	SiteName      string `toml:"site_name"`
	UID           uint32 `toml:"uid"`
	GID           uint32 `toml:"gid"`
	Socket        string `toml:"socket"`
	ConfigFile    string `toml:"config_file"`
	AppPath       string `toml:"app_path"`
	MemoryLimitMB int    `toml:"memory_limit_mb"`
	MaxRestarts   int    `toml:"max_restarts"`
	TmpDir        string `toml:"tmp_dir"`
}

const (
	ConfigDir  = "/opt/axonasp/fpm/fpm.d/"
	WorkerExec = "/opt/axonasp/axonasp-fastcgi"
)

// Global state to track running pools and prevent duplicates during reload
var (
	activePools = make(map[string]context.CancelFunc)
	poolsMutex  sync.Mutex
)

// normalizePoolSocketEndpoint normalizes pool socket configuration and returns
// the FastCGI listen endpoint plus a filesystem path when using unix sockets.
func normalizePoolSocketEndpoint(raw string) (listenEndpoint string, socketPath string, isUnix bool, err error) {
	value := strings.TrimSpace(raw)
	if value == "" {
		return "", "", false, fmt.Errorf("socket is required in pool config")
	}

	lower := strings.ToLower(value)
	if strings.HasPrefix(lower, "unix:") {
		path := strings.TrimSpace(value[len("unix:"):])
		if path == "" {
			return "", "", false, fmt.Errorf("unix socket path cannot be empty")
		}
		return "unix:" + path, path, true, nil
	}

	if strings.HasPrefix(value, "/") || strings.HasPrefix(value, "./") || strings.HasPrefix(value, "../") {
		return "unix:" + value, value, true, nil
	}

	return value, "", false, nil
}

func main() {
	//Require root privileges to run the FPM manager, as it needs to manage worker processes and potentially bind to privileged ports or create sockets in system directories.
	if os.Geteuid() != 0 {
		log.Fatal("Fatal Error: G3pix ❖ AxonASP FPM must be run as root.")
	}

	// 1. Initial Load of Configurations
	scanAndLoadConfigs()

	// 2. Setup Signal Handling for Graceful Reload (SIGHUP) and shutdown (SIGINT/SIGTERM)
	sigChan := make(chan os.Signal, 1)
	signal.Notify(sigChan, syscall.SIGHUP, syscall.SIGINT, syscall.SIGTERM)

	// 3. Periodic scan allows new pool files to be detected without requiring SIGHUP.
	//rescanTicker := time.NewTicker(5 * time.Second)
	//defer rescanTicker.Stop()

	log.Println("G3pix ❖ AxonASP FPM Ready")

	// 4. Main Event Loop
	for sig := range sigChan {
		switch sig {
		case syscall.SIGHUP:
			log.Println("SIGHUP received. Rescanning configuration directory for new applications...")
			scanAndLoadConfigs()
		case syscall.SIGINT, syscall.SIGTERM:
			log.Printf("%s received.\n Shutting down G3pix ❖ AxonASP FPM manager...", sig.String())
			shutdownAllPools()
			return
		}
	}
}

// scanAndLoadConfigs reads the config directory and starts supervisors for NEW files only.
func scanAndLoadConfigs() {
	poolsMutex.Lock()
	defer poolsMutex.Unlock()

	files, err := os.ReadDir(ConfigDir)
	if err != nil {
		//We use Fatalf here because if we can't read the config directory, we can't proceed.
		//we also don't want to try to create the directory automatically, as that could lead to unexpected behavior. The user should ensure the directory exists and is readable.
		log.Fatalf("Failed to read configuration directory: %v\nExiting...", err)
	}

	for _, file := range files {
		//Look for .conf files only, ignoring other files or directories.
		if filepath.Ext(file.Name()) == ".conf" {
			configPath := filepath.Join(ConfigDir, file.Name())

			// Check if this config is already being supervised
			if _, exists := activePools[configPath]; !exists {
				log.Printf("❖ New configuration detected: %s. Starting worker pool...", file.Name())

				// Create a cancellable context for this specific worker pool
				ctx, cancel := context.WithCancel(context.Background())
				activePools[configPath] = cancel

				go superviseWorker(ctx, configPath)
			}
		}
	}
}

func shutdownAllPools() {
	poolsMutex.Lock()
	defer poolsMutex.Unlock()

	for configPath, cancel := range activePools {
		cancel()
		delete(activePools, configPath)
	}
}

func superviseWorker(ctx context.Context, configPath string) {
	data, err := os.ReadFile(configPath)
	if err != nil {
		log.Printf("Error reading %s: %v", configPath, err)
		return
	}

	var conf PoolConfig
	err = toml.Unmarshal(data, &conf)
	if err != nil {
		log.Printf("Error parsing TOML %s: %v", configPath, err)
		return
	}

	if conf.TmpDir == "" {
		conf.TmpDir = "/opt/axonasp/temp/"
	}

	if conf.AppPath == "" {
		log.Printf("[%s] Error: app_path is required in pool config", conf.SiteName)
		return
	}
	appPathInfo, err := os.Stat(conf.AppPath)
	if err != nil {
		log.Printf("[%s] Error validating app_path %q: %v", conf.SiteName, conf.AppPath, err)
		return
	}
	if !appPathInfo.IsDir() {
		log.Printf("[%s] Error: app_path %q is not a directory", conf.SiteName, conf.AppPath)
		return
	}

	listenEndpoint, socketPath, isUnixSocket, err := normalizePoolSocketEndpoint(conf.Socket)
	if err != nil {
		log.Printf("[%s] Error: invalid socket value %q: %v", conf.SiteName, conf.Socket, err)
		return
	}

	// Create Directories
	if err := os.MkdirAll(conf.TmpDir, 0755); err != nil {
		log.Printf("[%s] Error creating temp directory: %v", conf.SiteName, err)
		return
	}
	if err := os.Chown(conf.TmpDir, int(conf.UID), int(conf.GID)); err != nil {
		log.Printf("[%s] Error setting permissions on temp directory: %v", conf.SiteName, err)
		return
	}

	if isUnixSocket {
		socketDir := filepath.Dir(socketPath)
		if err := os.MkdirAll(socketDir, 0755); err != nil {
			log.Printf("[%s] Error creating socket directory: %v", conf.SiteName, err)
			return
		}
		if err := os.Chown(socketDir, int(conf.UID), int(conf.GID)); err != nil {
			log.Printf("[%s] Error setting permissions on socket directory: %v", conf.SiteName, err)
			return
		}

		_ = os.Remove(socketPath)
	}

	restarts := 0

	for {
		log.Printf("[%s] Starting Worker (Attempt: %d) with %dMB of RAM", conf.SiteName, restarts, conf.MemoryLimitMB)

		// Use CommandContext so the process can be killed cleanly if the context is cancelled, we still need to support it in the fastcgi implementation

		cmd := exec.CommandContext(ctx, WorkerExec, "--fastcgi.server_port", listenEndpoint, "--config.config_file", conf.ConfigFile)

		cmd.Dir = conf.AppPath
		log.Printf("[%s] Executing: %s", conf.SiteName, strings.Join(cmd.Args, " "))
		log.Printf("[%s] Current Directory: %s", conf.SiteName, cmd.Dir)

		cmd.SysProcAttr = &syscall.SysProcAttr{
			Credential: &syscall.Credential{
				Uid: conf.UID,
				Gid: conf.GID,
			},
		}

		env := os.Environ()
		env = append(env, fmt.Sprintf("GOMEMLIMIT=%dMiB", conf.MemoryLimitMB))
		env = append(env, fmt.Sprintf("GLOBAL_GOLANG_MEMORY_LIMIT_MB=%dMiB", conf.MemoryLimitMB))
		env = append(env, fmt.Sprintf("GLOBAL_TEMP_DIR=%s", conf.TmpDir))
		env = append(env, fmt.Sprintf("TMPDIR=%s", conf.TmpDir))
		cmd.Env = env

		cmd.Stdout = os.Stdout
		cmd.Stderr = os.Stderr

		if err := cmd.Start(); err != nil {
			log.Printf("[%s] Failed to start worker: %v", conf.SiteName, err)
		} else {
			if err := enforceCgroupMemoryLimit(conf.SiteName, cmd.Process.Pid, conf.MemoryLimitMB); err != nil {
				log.Printf("[%s] Warning: Failed to apply cgroup limit: %v", conf.SiteName, err)
			}
			err = cmd.Wait()
			log.Printf("[%s] Worker terminated. Error/Exit State: %v", conf.SiteName, err)
		}

		// Check if the worker stopped because we cancelled the context
		select {
		case <-ctx.Done():
			log.Printf("[%s] Pool supervisor shutting down (Context Cancelled).", conf.SiteName)
			return
		default:
		}

		if conf.MaxRestarts != 0 && restarts >= conf.MaxRestarts {
			log.Printf("[%s] Maximum restart limit reached (%d). Abandoning pool.", conf.SiteName, conf.MaxRestarts)
			break
		}

		restarts++
		// Wait a bit before restarting to avoid rapid restart loops
		time.Sleep(2 * time.Second)
	}
}

// enforceCgroupMemoryLimit remains exactly the same as previously defined
func enforceCgroupMemoryLimit(siteName string, pid int, memoryLimitMB int) error {
	cgroupBase := "/sys/fs/cgroup/axonasp"
	if err := ensureMemoryControllerDelegated(cgroupBase); err != nil {
		return err
	}

	poolCgroup := filepath.Join(cgroupBase, siteName)

	if err := os.MkdirAll(poolCgroup, 0755); err != nil {
		return fmt.Errorf("failed to create cgroup directory: %w", err)
	}

	limitBytes := fmt.Sprintf("%d", memoryLimitMB*1024*1024)
	limitFile := filepath.Join(poolCgroup, "memory.max")
	if err := writeCgroupControl(limitFile, limitBytes); err != nil {
		if isPermissionErr(err) {
			return fmt.Errorf("failed to write memory.max: %w (cgroup permission/delegation issue; verify systemd Delegate=yes and writable cgroup subtree)", err)
		}
		return fmt.Errorf("failed to write memory.max: %w", err)
	}

	procsFile := filepath.Join(poolCgroup, "cgroup.procs")
	pidStr := fmt.Sprintf("%d", pid)
	if err := writeCgroupControl(procsFile, pidStr); err != nil {
		return fmt.Errorf("failed to attach PID to cgroup: %w", err)
	}

	return nil
}

// ensureMemoryControllerDelegated verifies memory controller availability and
// tries to enable it for child cgroups if not already active.
func ensureMemoryControllerDelegated(cgroupBase string) error {
	controllersData, err := os.ReadFile(filepath.Join(cgroupBase, "cgroup.controllers"))
	if err != nil {
		if isPermissionErr(err) {
			return fmt.Errorf("failed to read cgroup controllers: %w (missing permission to inspect cgroup delegation)", err)
		}
		return fmt.Errorf("failed to read cgroup controllers: %w", err)
	}

	controllers := string(controllersData)
	if !hasController(controllers, "memory") {
		return fmt.Errorf("memory controller is not available in %s/cgroup.controllers", cgroupBase)
	}

	subtreePath := filepath.Join(cgroupBase, "cgroup.subtree_control")
	subtreeData, err := os.ReadFile(subtreePath)
	if err != nil {
		if isPermissionErr(err) {
			return fmt.Errorf("failed to read cgroup.subtree_control: %w (missing permission to inspect cgroup delegation)", err)
		}
		return fmt.Errorf("failed to read cgroup.subtree_control: %w", err)
	}

	if hasController(string(subtreeData), "memory") {
		return nil
	}

	if err := writeCgroupControl(subtreePath, "+memory"); err != nil {
		if isPermissionErr(err) {
			return fmt.Errorf("memory controller not delegated for child cgroups in %s: %w (set Delegate=yes in the service unit and enable +memory in cgroup.subtree_control)", cgroupBase, err)
		}
		return fmt.Errorf("failed to enable memory controller in cgroup.subtree_control: %w", err)
	}

	return nil
}

func hasController(list string, controller string) bool {
	for _, item := range strings.Fields(list) {
		if item == controller {
			return true
		}
	}
	return false
}

func writeCgroupControl(path string, value string) error {
	f, err := os.OpenFile(path, os.O_WRONLY|os.O_TRUNC, 0)
	if err != nil {
		return err
	}
	defer f.Close()

	if _, err := f.WriteString(value); err != nil {
		return err
	}

	return nil
}

func isPermissionErr(err error) bool {
	return errors.Is(err, os.ErrPermission) ||
		errors.Is(err, syscall.EPERM) ||
		errors.Is(err, syscall.EACCES)
}
