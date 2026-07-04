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
	"fmt"
	"log"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
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
	ConfigFile	  string `toml:"config_file"`
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

func main() {
	if os.Geteuid() != 0 {
		log.Fatal("Error: The AxonASP Manager must be run as root.")
	}

	log.Println("Starting AxonASP FPM Manager...")

	// 1. Initial Load of Configurations
	scanAndLoadConfigs()

	// 2. Setup Signal Handling for Graceful Reload (SIGHUP)
	sigChan := make(chan os.Signal, 1)
	signal.Notify(sigChan, syscall.SIGHUP)

	log.Println("AxonASP FPM Manager is running...")

	// 3. Main Event Loop
	for sig := range sigChan {
		_ = sig
		log.Println("Received Sighup. Rescanning configuration directory for new applications...")
		scanAndLoadConfigs()
	}
}

// scanAndLoadConfigs reads the config directory and starts supervisors for NEW files only.
func scanAndLoadConfigs() {
	poolsMutex.Lock()
	defer poolsMutex.Unlock()

	files, err := os.ReadDir(ConfigDir)
	if err != nil {
		log.Printf("Failed to read configuration directory: %v", err)
		return
	}

	for _, file := range files {
		if filepath.Ext(file.Name()) == ".conf" {
			configPath := filepath.Join(ConfigDir, file.Name())

			// Check if this config is already being supervised
			if _, exists := activePools[configPath]; !exists {
				log.Printf("New configuration detected: %s. Starting worker pool...", file.Name())

				// Create a cancellable context for this specific worker pool
				ctx, cancel := context.WithCancel(context.Background())
				activePools[configPath] = cancel

				go superviseWorker(ctx, configPath)
			}
		}
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

	// Create Directories
	if err := os.MkdirAll(conf.TmpDir, 0755); err != nil {
		log.Printf("[%s] Error creating temp directory: %v", conf.SiteName, err)
		return
	}
	if err := os.Chown(conf.TmpDir, int(conf.UID), int(conf.GID)); err != nil {
		log.Printf("[%s] Error setting permissions on temp directory: %v", conf.SiteName, err)
		return
	}

	socketDir := filepath.Dir(conf.Socket)
	if err := os.MkdirAll(socketDir, 0755); err != nil {
		log.Printf("[%s] Error creating socket directory: %v", conf.SiteName, err)
		return
	}
	if err := os.Chown(socketDir, int(conf.UID), int(conf.GID)); err != nil {
		log.Printf("[%s] Error setting permissions on socket directory: %v", conf.SiteName, err)
		return
	}

	os.Remove(conf.Socket)
	restarts := 0

	for {
		log.Printf("[%s] Starting Worker (Attempt: %d) with %dMB of RAM", conf.SiteName, restarts, conf.MemoryLimitMB)

		// Use CommandContext so the process can be killed cleanly if the context is cancelled, we still need to support it in the fastcgi implementation
		cmd := exec.CommandContext(ctx, WorkerExec, "--fastcgi.server_port", conf.Socket, "--config.config_file", conf.ConfigFile)

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
	poolCgroup := filepath.Join(cgroupBase, siteName)

	if err := os.MkdirAll(poolCgroup, 0755); err != nil {
		return fmt.Errorf("failed to create cgroup directory: %w", err)
	}

	limitBytes := fmt.Sprintf("%d", memoryLimitMB*1024*1024)
	limitFile := filepath.Join(poolCgroup, "memory.max")
	if err := os.WriteFile(limitFile, []byte(limitBytes), 0644); err != nil {
		return fmt.Errorf("failed to write memory.max: %w", err)
	}

	procsFile := filepath.Join(poolCgroup, "cgroup.procs")
	pidStr := fmt.Sprintf("%d", pid)
	if err := os.WriteFile(procsFile, []byte(pidStr), 0644); err != nil {
		return fmt.Errorf("failed to attach PID to cgroup: %w", err)
	}

	return nil
}