/*
 * AxonASP Server
 * Copyright (C) 2026 G3pix Ltda. All rights reserved.
 *
 * Developed by Jeffrey He (@jeffreyheping)
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
	"fmt"
	"log"
	"os"
	"os/exec"
	"os/signal"
	"path/filepath"
	"runtime"
	"strings"
	"sync/atomic"
	"syscall"
	"time"
)

// lastHeartbeat stores the Unix timestamp of the most recent /__heartbeat__
// request from the browser page. Updated by the heartbeat handler and
// monitored by monitorHeartbeat to detect browser window closure.
var lastHeartbeat atomic.Int64

// openWindow opens the application URL in a desktop window. It first tries
// to launch Chrome, Edge, or Chromium in app mode (no browser chrome, no
// tabs) for a native desktop feel. If none is found, it falls back to the
// system default browser. In app mode, the function blocks until the window
// is closed; in fallback mode, it blocks until Ctrl+C or SIGTERM.
func openWindow(url string) {
	if browserPath := findChromiumBrowser(); browserPath != "" {
		openAppWindow(browserPath, url)
		return
	}
	log.Println("No Chrome/Edge/Chromium found, falling back to default browser")
	openDefaultBrowser(url)
}

// findChromiumBrowser searches for a Chromium-based browser on the system.
// Returns the executable path, or empty string if none is found.
func findChromiumBrowser() string {
	var candidates []string

	switch runtime.GOOS {
	case "windows":
		candidates = []string{
			// Microsoft Edge (built into Windows 10/11)
			`C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`,
			`C:\Program Files\Microsoft\Edge\Application\msedge.exe`,
			// Google Chrome
			`C:\Program Files\Google\Chrome\Application\chrome.exe`,
			`C:\Program Files (x86)\Google\Chrome\Application\chrome.exe`,
		}
	case "darwin":
		candidates = []string{
			"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
			"/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
			"/Applications/Chromium.app/Contents/MacOS/Chromium",
			"/Applications/Brave Browser.app/Contents/MacOS/Brave Browser",
		}
	default:
		// Linux and other Unix-like systems: search via PATH.
		for _, name := range []string{
			"google-chrome",
			"google-chrome-stable",
			"chromium",
			"chromium-browser",
			"microsoft-edge",
			"microsoft-edge-stable",
			"brave-browser",
		} {
			if path, err := exec.LookPath(name); err == nil {
				return path
			}
		}
		return ""
	}

	for _, c := range candidates {
		if _, err := os.Stat(c); err == nil {
			return c
		}
	}
	return ""
}

// appUserDataDir returns a persistent browser profile directory unique to
// the current application, so each app gets its own isolated profile and
// always opens in a dedicated window instead of a tab in an existing
// browser session.
func appUserDataDir() string {
	name := filepath.Base(appDir)
	// Sanitize for use as a directory name across platforms.
	name = strings.NewReplacer(
		" ", "-", ".", "_", "/", "-", "\\", "-",
		":", "-", "*", "-", "?", "-", "\"", "-",
		"<", "-", ">", "-", "|", "-",
	).Replace(name)
	if name == "" || name == "." || name == "-" {
		name = "default"
	}
	return filepath.Join(os.TempDir(), "axonhta-"+name)
}

// waitForInterrupt blocks until the process receives an interrupt or
// SIGTERM signal.
func waitForInterrupt() {
	c := make(chan os.Signal, 1)
	signal.Notify(c, os.Interrupt, syscall.SIGTERM)
	<-c
}

// openAppWindow launches the browser in app mode and blocks until the
// window is closed. A heartbeat mechanism detects window closure
// regardless of whether Chrome forks (launcher exits early) or runs
// directly: injected JS pings /__heartbeat__ every 5 seconds, and
// monitorHeartbeat calls os.Exit when pings stop arriving.
func openAppWindow(browserPath, url string) {
	// Pre-create the user data directory so Chrome does not need to
	// initialize it on first launch, which can cause a delay or error.
	userDataDir := appUserDataDir()
	_ = os.MkdirAll(userDataDir, 0755)

	args := []string{
		"--app=" + url,
		fmt.Sprintf("--window-size=%d,%d", width, height),
		"--no-first-run",
		"--no-default-browser-check",
		"--user-data-dir=" + userDataDir,

		// Minimal desktop-app experience: disable browser features
		// that are unnecessary in app mode.
		"--disable-extensions",                       // No extensions
		"--disable-translate",                        // No translate bar
		"--disable-sync",                             // No account sync
		"--disable-default-apps",                     // No default apps prompt
		"--disable-component-update",                 // No background component updates
		"--disable-background-networking",            // Minimize background network activity
		"--autoplay-policy=no-user-gesture-required", // Allow media autoplay
	}

	// Apply HTA tag attributes if available.
	if htaConfig != nil {
		// windowstate="maximize" → start maximized
		if htaConfig.WindowState == "maximize" {
			args = append(args, "--start-maximized")
		}

		// caption="no" → borderless window.
		// Chrome app mode does not support borderless windows directly.
		if htaConfig.Caption != "" && !htaConfig.BoolAttr(htaConfig.Caption) {
			log.Printf("[HTA] caption=no is not supported in app mode")
		}

		// icon: Chrome uses the page's favicon as the window icon.
		// A custom icon per-app is not supported via command-line flags.
		if htaConfig.Icon != "" {
			log.Printf("[HTA] per-app icon via HTA tag is not supported in app mode (icon=%q)", htaConfig.Icon)
		}
	}

	cmd := exec.Command(browserPath, args...)
	setNoConsole(cmd)
	if err := cmd.Start(); err != nil {
		log.Printf("Failed to launch app window: %v, falling back to default browser", err)
		openDefaultBrowser(url)
		return
	}

	log.Printf("App window opened (browser: %s)", filepath.Base(browserPath))
	log.Println("Close the window to exit.")

	// monitorHeartbeat waits for the first heartbeat, then calls os.Exit
	// when heartbeats stop (browser window closed).
	go monitorHeartbeat()

	// Wait for the Chrome process to exit.
	// If Chrome doesn't fork, this blocks until the window is closed.
	// If Chrome forks, this returns immediately (launcher exited), but
	// the heartbeat monitor keeps the process alive.
	_ = cmd.Wait()

	// If heartbeats were received, Chrome forked — the launcher exited
	// but the browser window is still open. Block here; the heartbeat
	// monitor will call os.Exit when the window is closed.
	if lastHeartbeat.Load() > 0 {
		select {} // block forever; heartbeat monitor handles exit
	}

	// No heartbeats ever received — browser didn't start or crashed.
	log.Println("Shutting down AxonHTA...")
}

// monitorHeartbeat waits for the first heartbeat from the browser page,
// then polls until heartbeats stop, indicating the window was closed.
// When that happens it calls os.Exit(0) to shut down axonhta.exe.
func monitorHeartbeat() {
	// Wait for first heartbeat (60s timeout for slow page loads).
	deadline := time.Now().Add(60 * time.Second)
	for time.Now().Before(deadline) {
		if lastHeartbeat.Load() > 0 {
			break
		}
		time.Sleep(500 * time.Millisecond)
	}

	if lastHeartbeat.Load() == 0 {
		log.Println("No heartbeat received from browser, shutting down...")
		os.Exit(1)
	}

	// Poll until heartbeats stop (15s grace period covers form-submission
	// navigation gaps: when the browser submits a POST and follows a 302
	// redirect, the old page's JS stops and the new page's JS hasn't started
	// yet. With a 5-second heartbeat interval, the worst-case gap before the
	// new page sends its first heartbeat can approach 10 seconds.
	// 15 seconds provides a safe margin without delaying window-close
	// detection excessively.
	for {
		last := lastHeartbeat.Load()
		if time.Now().Unix()-last > 15 {
			log.Println("Browser window closed, shutting down...")
			os.Exit(0)
		}
		time.Sleep(2 * time.Second)
	}
}

// openDefaultBrowser opens the URL in the system default browser and blocks
// until the process receives an interrupt or SIGTERM signal.
func openDefaultBrowser(url string) {
	log.Printf("Opening browser: %s", url)

	var cmd string
	var args []string

	switch runtime.GOOS {
	case "windows":
		cmd = "cmd"
		args = []string{"/c", "start", url}
	case "darwin":
		cmd = "open"
		args = []string{url}
	default:
		cmd = "xdg-open"
		args = []string{url}
	}

	c := exec.Command(cmd, args...)
	setNoConsole(c)
	if err := c.Start(); err != nil {
		fmt.Printf("Failed to open browser. Please navigate to: %s\n", url)
	}

	fmt.Println("Press Ctrl+C to stop")

	waitForInterrupt()

	log.Println("Shutting down AxonHTA...")
}
