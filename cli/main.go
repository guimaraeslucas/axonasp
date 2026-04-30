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
//You need to specify -64=false/-arm=true if you're trying to create an 32-bit or ARM windows binary, this is required by the new version of golang
//go:generate goversioninfo -icon=icon_cli.ico -64=true
package main

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"runtime/debug"
	"slices"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
	"github.com/gdamore/tcell/v2"
	"github.com/joho/godotenv"
	"github.com/rivo/tview"
)

// CLI configuration values.
var (
	Version                     = "0.0.0.0"
	EnableCLI                   = true
	EnableCLIRunFromCommandLine = false
	DebugASP                    = false
	DefaultTimezone             = "UTC"
	MemoryLimitMB               = 128
	VMPoolSize                  = 50
	BytecodeCachingMode         = "enabled"
	CacheMaxSizeMB              = 128
	CLIForceFreshCompile        = true
	ScriptTimeout               = 60
	ResponseBufferLimitBytes    = 4 * 1024 * 1024
	ExecuteAsASPExtensions      = []string{".asp"}
	CLIServerRoot               = "./www"
	serverLocation              = time.UTC
	mouseEnabled                = false
	scriptCache                 *axonvm.ScriptCache

	// Process-wide shared Application and Session instances for the CLI
	sharedCLIApplication = asp.NewApplication()
	sharedCLISession     = asp.NewSession()
)

const tuiHelpText = `
 HOTKEYS:

  F3:   Run current code in Input Area
  F2:   Open and Run an external ASP file
  F7:   Toggle Auto-Run mode (On/Off)
  F6:   Switch focus between Input and Output
  F8:   Clear Output Area
  F5:   Toggle Mouse Support
  F1:   Show this help window
  F4:   Exit the application
  Esc:  Close current dialog/window

 COMMAND LINE USAGE:

  You can execute ASP scripts directly from your terminal
  or batch files using the -r or --run flags:
  
  > axonasp-cli.exe -r "C:\scripts\maintenance.asp"
  
  This is useful for:
  - Local automation and system maintenance tasks.
  - Database cleanup or migration scripts.
  - Batch processing files using G3Zip, G3FC or G3PDF.
  - Running background jobs or scheduled tasks via Task 
    Scheduler/Cron.

 ABOUT:
  G3pix ❖ AxonASP
  is a high-performance, cross-platform Classic ASP engine,
  with support to VBScript and JScript for Web, FastCGI, 
  and CLI, bridging legacy compatibility with modern 
  APIs.
  
  Copyright (C) 2026 G3pix Ltda. All rights reserved.
  Website: https://g3pix.com.br/axonasp
  
  License: MPL 2.0
`

// init loads environment variables and applies TOML-based configuration through Viper.
func init() {
	_ = godotenv.Load()
	axonvm.SetRuntimeVersion(strings.TrimSpace(Version))
	loadCLIConfig()
	applyRuntimeSettings()
}

// loadCLIConfig loads and applies cli/global settings from config/axonasp.toml using Viper.
func loadCLIConfig() {
	v := axonconfig.NewViper()
	if strings.TrimSpace(v.ConfigFileUsed()) == "" {
		fmt.Printf("Warning: %s\n", axonvm.ErrViperReadConfigFailed.String())
	}
	axonconfig.EnableWatchIfConfigured(v, nil)

	EnableCLI = v.GetBool("cli.enable_cli")
	EnableCLIRunFromCommandLine = v.GetBool("cli.enable_cli_run_from_command_line")
	if v.IsSet("cli.force_fresh_compile") {
		CLIForceFreshCompile = v.GetBool("cli.force_fresh_compile")
	}
	DebugASP = v.GetBool("global.enable_asp_debugging")
	axonvm.SetDumpPreprocessedSourceEnabled(v.GetBool("global.dump_preprocessed_source"))
	if executeAsASP := v.GetStringSlice("global.execute_as_asp"); len(executeAsASP) > 0 {
		ExecuteAsASPExtensions = normalizeExtensions(executeAsASP)
	}
	if webRoot := strings.TrimSpace(v.GetString("server.web_root")); webRoot != "" {
		CLIServerRoot = webRoot
	}
	if timezone := strings.TrimSpace(v.GetString("global.default_timezone")); timezone != "" {
		DefaultTimezone = timezone
	}
	if memoryMB := v.GetInt("global.golang_memory_limit_mb"); memoryMB >= 0 {
		MemoryLimitMB = memoryMB
	}
	if vmPoolSize := v.GetInt("global.vm_pool_size"); vmPoolSize >= 0 {
		VMPoolSize = vmPoolSize
	}
	BytecodeCachingMode = strings.TrimSpace(v.GetString("global.bytecode_caching_enabled"))
	if BytecodeCachingMode == "" {
		BytecodeCachingMode = strings.TrimSpace(v.GetString("global.bytecode_caching"))
	}
	if BytecodeCachingMode == "" {
		BytecodeCachingMode = "enabled"
	}
	if cacheSizeMB := v.GetInt("global.cache_max_size_mb"); cacheSizeMB > 0 {
		CacheMaxSizeMB = cacheSizeMB
	}
	axonvm.SetVMPoolSizeLimit(VMPoolSize)
	ScriptTimeout = v.GetInt("global.default_script_timeout")
	if ScriptTimeout <= 0 {
		ScriptTimeout = 60
	}
	if responseBufferLimitMB := v.GetInt("global.response_buffer_limit_mb"); responseBufferLimitMB > 0 {
		ResponseBufferLimitBytes = responseBufferLimitMB * 1024 * 1024
	}

	axonvm.InitGlobalAxonFunctions(v.GetBool("axfunctions.enable_global_ax"))
}

// applyRuntimeSettings applies timezone and Go memory limit based on loaded configuration.
func applyRuntimeSettings() {
	if MemoryLimitMB > 0 {
		debug.SetMemoryLimit(int64(MemoryLimitMB) * 1024 * 1024)
	}

	os.Setenv("TZ", DefaultTimezone)
	location, err := axonvm.ResolveTimezoneLocation(DefaultTimezone)
	if err != nil {
		fmt.Printf("Warning: %s %s, using UTC: %v\n", axonvm.ErrInvalidTimezone.String(), DefaultTimezone, err)
		location = time.UTC
	}
	serverLocation = location
	time.Local = location
}

// normalizeExtensions normalizes file extensions to lowercase ".ext" values.
func normalizeExtensions(values []string) []string {
	cleaned := make([]string, 0, len(values))
	for _, value := range values {
		ext := strings.ToLower(strings.TrimSpace(value))
		if ext == "" {
			continue
		}
		if !strings.HasPrefix(ext, ".") {
			ext = "." + ext
		}
		if !slices.Contains(cleaned, ext) {
			cleaned = append(cleaned, ext)
		}
	}
	return cleaned
}

// isASPExecutionExtension reports whether a file should be executed as ASP based on configured extensions.
func isASPExecutionExtension(filePath string) bool {
	ext := strings.ToLower(filepath.Ext(filePath))
	return slices.Contains(ExecuteAsASPExtensions, ext)
}

// main starts the interactive CLI and handles commands.
func main() {
	workingDir, _ := os.Getwd()
	cacheMode := axonvm.ParseBytecodeCacheMode(BytecodeCachingMode)
	cacheMaxSizeMB := CacheMaxSizeMB
	if CLIForceFreshCompile {
		cacheMode = axonvm.BytecodeCacheDisabled
		cacheMaxSizeMB = 1
	}

	scriptCache = axonvm.NewScriptCache(
		cacheMode,
		filepath.Join("temp", "cache"),
		cacheMaxSizeMB,
	)
	if !CLIForceFreshCompile {
		if err := scriptCache.StartInvalidator([]string{workingDir}); err != nil {
			fmt.Printf("Warning: failed to start bytecode invalidator: %v\n", err)
		}
		defer scriptCache.StopInvalidator()
	}

	if err := axonvm.GetGlobalASA().LoadAndCompile(workingDir, sharedCLIApplication); err != nil {
		fmt.Printf("Warning: Failed to load global.asa: %v\n", err)
	} else if axonvm.GetGlobalASA().IsLoaded() {
		dummyHost := newCLIHost(new(bytes.Buffer), "/", false)
		_ = axonvm.GetGlobalASA().ExecuteApplicationOnStart(dummyHost)
		axonvm.GetGlobalASA().PopulateSessionStaticObjects(sharedCLISession)
		_ = axonvm.GetGlobalASA().ExecuteSessionOnStart(dummyHost)
	}

	defer func() {
		if axonvm.GetGlobalASA().IsLoaded() {
			dummyHost := newCLIHost(new(bytes.Buffer), "/", false)
			_ = axonvm.GetGlobalASA().ExecuteSessionOnEnd(dummyHost)
			_ = axonvm.GetGlobalASA().ExecuteApplicationOnEnd(dummyHost)
		}
	}()

	// Handle command-line arguments for help and direct file execution before starting REPL.
	if len(os.Args) > 1 {
		arg := os.Args[1]
		if arg == "-h" || arg == "--help" {
			printHelp()
			return
		}

		if arg == "-r" || arg == "--run" {
			if !EnableCLIRunFromCommandLine {
				fmt.Printf("Error %d: %s\n", axonvm.ErrCLIRunCommandNotEnabled, axonvm.ErrCLIRunCommandNotEnabled.String())
				os.Exit(int(axonvm.ErrCLIRunCommandNotEnabled))
			}
			if len(os.Args) < 3 {
				fmt.Printf("Error %d: %s\n", axonvm.ErrCLIMissingFilePath, axonvm.ErrCLIMissingFilePath.String())
				printHelp()
				os.Exit(int(axonvm.ErrCLIMissingFilePath))
			}
			runDirectFile(os.Args[2])
			return
		}

	}

	// Check if CLI is enabled in configuration before starting.
	if !EnableCLI {
		fmt.Printf("Error %d: %s\n", axonvm.ErrCLINotEnabled, axonvm.ErrCLINotEnabled.String())
		os.Exit(int(axonvm.ErrCLINotEnabled))
	}

	startTUI()
}

// startTUI initializes and runs the Terminal User Interface.
func startTUI() {
	app := tview.NewApplication()
	app.EnableMouse(mouseEnabled)
	autoRunEnabled := true
	app.SetTitle("G3pix ❖ AxonASP CLI")

	// Theme and Colors
	bgColor := tcell.GetColor("#000080")
	dialogBg := tcell.GetColor("#010043")
	tview.Styles.PrimitiveBackgroundColor = bgColor
	tview.Styles.ContrastBackgroundColor = bgColor
	tview.Styles.BorderColor = tcell.ColorWhite
	tview.Styles.TitleColor = tcell.ColorYellow

	// Header - Full width white background
	header := tview.NewTextView().
		SetDynamicColors(true).
		SetText("[#010043:white]  G3pix ❖ AxonASP ").
		SetTextAlign(tview.AlignLeft)
	header.SetBackgroundColor(tcell.ColorWhite)

	// Main Area
	inputArea := tview.NewTextArea()
	inputArea.SetBorder(true).
		SetTitle(" ASP Code Input ").
		SetTitleAlign(tview.AlignLeft).Blur()
	inputArea.SetBorderColor(tcell.ColorYellow)

	outputArea := tview.NewTextView().
		SetDynamicColors(true).
		SetRegions(true).
		SetToggleHighlights(true).
		SetWordWrap(true).
		SetChangedFunc(func() {
			app.Draw()
		})
	outputArea.SetBorder(true).
		SetTitle(" Output ").
		SetTitleAlign(tview.AlignLeft)

	// System Status Widget (Time + Auto-Run)
	statusWidget := tview.NewTextView().
		SetTextAlign(tview.AlignCenter).
		SetDynamicColors(true)
	statusWidget.SetBorder(true).SetTitle(" AxonASP Status ")

	updateStatus := func() {
		status := "[green]ON[white]"
		if !autoRunEnabled {
			status = "[red]OFF[white]"
		}
		now := time.Now().In(serverLocation)
		statusWidget.SetText(fmt.Sprintf("[yellow]%s[white]  |  Auto-Run: %s", now.Format("2006-01-02 15:04:05"), status))
	}

	go func() {
		for {
			updateStatus()
			app.QueueUpdateDraw(func() {})
			time.Sleep(1 * time.Second)
		}
	}()

	rightPanel := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(outputArea, 0, 1, false).
		AddItem(statusWidget, 3, 1, false)

	clearOutput := func() {
		outputArea.Clear()
	}

	mainArea := tview.NewFlex().
		AddItem(inputArea, 0, 1, true).
		AddItem(rightPanel, 0, 1, false)

	// Navigation Pages
	pages := tview.NewPages()

	footer := tview.NewFlex().SetDirection(tview.FlexRow)

	// Buttons row
	btnRow := tview.NewFlex().
		AddItem(tview.NewButton("Run (F3)").SetSelectedFunc(func() {
			runCode(inputArea.GetText(), outputArea)
		}), 15, 1, false).
		AddItem(tview.NewButton("Open (F2)").SetSelectedFunc(func() {
			showRunFileDialog(app, pages, outputArea)
		}), 15, 1, false).
		AddItem(tview.NewButton("Auto (F7)").SetSelectedFunc(func() {
			autoRunEnabled = !autoRunEnabled
			updateStatus()
		}), 15, 1, false).
		AddItem(tview.NewButton("Focus (F6)").SetSelectedFunc(func() {
			if inputArea.HasFocus() {
				app.SetFocus(outputArea)
			} else {
				app.SetFocus(inputArea)
			}
		}), 16, 1, false).
		AddItem(tview.NewButton("Clear (F8)").SetSelectedFunc(func() {
			clearOutput()
		}), 16, 1, false).
		AddItem(tview.NewButton("Mouse (F5)").SetSelectedFunc(func() {
			mouseEnabled = !mouseEnabled
			app.EnableMouse(mouseEnabled)
		}), 16, 1, false).
		AddItem(tview.NewButton("Help (F1)").SetSelectedFunc(func() {
			showHelpModal(app, pages)
		}), 15, 1, false).
		AddItem(tview.NewButton("Quit (F4)").SetSelectedFunc(func() {
			app.Stop()
		}), 15, 1, false)

	footer.AddItem(btnRow, 1, 1, false)
	footer.SetBackgroundColor(dialogBg)

	// Layout
	layout := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(header, 1, 1, false).
		AddItem(mainArea, 0, 1, true).
		AddItem(footer, 1, 1, false)
	pages.AddPage("main", layout, true, true)

	// Debounce Logic
	var timer *time.Timer
	var mu sync.Mutex
	inputArea.SetChangedFunc(func() {
		if !autoRunEnabled {
			return
		}
		mu.Lock()
		defer mu.Unlock()
		if timer != nil {
			timer.Stop()
		}
		timer = time.AfterFunc(500*time.Millisecond, func() {
			app.QueueUpdateDraw(func() {
				runCode(inputArea.GetText(), outputArea)
			})
		})
	})

	// Input handling
	app.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		switch event.Key() {
		case tcell.KeyF3:
			runCode(inputArea.GetText(), outputArea)
			return nil
		case tcell.KeyF7:
			autoRunEnabled = !autoRunEnabled
			updateStatus()
			return nil
		case tcell.KeyF5:
			mouseEnabled = !mouseEnabled
			app.EnableMouse(mouseEnabled)
			return nil
		case tcell.KeyF2:
			showRunFileDialog(app, pages, outputArea)
			return nil
		case tcell.KeyF6:
			// Cycle focus
			if inputArea.HasFocus() {
				app.SetFocus(outputArea)
			} else {
				app.SetFocus(inputArea)
			}
			return nil
		case tcell.KeyF8:
			clearOutput()
			return nil
		case tcell.KeyF1:
			showHelpModal(app, pages)
			return nil
		case tcell.KeyF4:
			app.Stop()
			return nil
		}
		return event
	})

	if err := app.SetRoot(pages, true).SetFocus(inputArea).Run(); err != nil {
		fmt.Printf("AxonASP TUI Error: %v\n", err)
		os.Exit(1)
	}
}

// showHelpModal displays a modal with help information.
func showHelpModal(app *tview.Application, pages *tview.Pages) {
	dialogBg := tcell.GetColor("#010043")
	textView := tview.NewTextView().
		SetDynamicColors(true).
		SetText(tuiHelpText).
		SetTextAlign(tview.AlignLeft).
		SetWordWrap(true)
	textView.SetBorder(true).
		SetTitle(" G3pix ❖ AxonASP Help ").
		SetTitleAlign(tview.AlignLeft)
	textView.SetBackgroundColor(dialogBg)
	textView.SetBorderColor(tcell.ColorGray)

	// Add 1-point padding via Flex
	paddedView := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(nil, 1, 1, false).
		AddItem(tview.NewFlex().SetDirection(tview.FlexColumn).
			AddItem(nil, 2, 1, false).
			AddItem(textView, 0, 1, true).
			AddItem(nil, 2, 1, false), 0, 1, true).
		AddItem(nil, 1, 1, false)
	paddedView.SetBackgroundColor(dialogBg)

	flex := tview.NewFlex().
		AddItem(nil, 0, 1, false).
		AddItem(tview.NewFlex().SetDirection(tview.FlexRow).
			AddItem(nil, 0, 1, false).
			AddItem(paddedView, 26, 1, true).
			AddItem(nil, 0, 1, false), 65, 1, true).
		AddItem(nil, 0, 1, false)

	textView.SetInputCapture(func(event *tcell.EventKey) *tcell.EventKey {
		if event.Key() == tcell.KeyEscape || event.Rune() == 'q' {
			pages.RemovePage("help")
			return nil
		}
		return event
	})

	pages.AddPage("help", flex, true, true)
	app.SetFocus(textView)
}

// showRunFileDialog displays a dialog to input a file path and execute it.
func showRunFileDialog(app *tview.Application, pages *tview.Pages, outputArea *tview.TextView) {
	dialogBg := tcell.GetColor("#010043")
	form := tview.NewForm()

	form.SetBorder(true).SetTitle(" Run ASP File ").SetTitleAlign(tview.AlignLeft)
	form.SetBackgroundColor(dialogBg)
	form.SetFieldBackgroundColor(tcell.ColorWhite)
	form.SetFieldTextColor(tcell.ColorBlack)
	form.SetButtonBackgroundColor(tcell.ColorSilver)
	form.SetButtonTextColor(tcell.ColorBlack)

	filePath := ""
	form.AddInputField("File Path", "", 48, nil, func(text string) {
		filePath = text
	})

	form.AddButton("Run", func() {
		pages.RemovePage("runfile")
		runTUIFile(filePath, outputArea)
		app.SetFocus(outputArea)
	})

	form.AddButton("Cancel", func() {
		pages.RemovePage("runfile")
	})

	// Instructional text
	helpText := tview.NewTextView().
		SetDynamicColors(true).
		SetTextAlign(tview.AlignCenter).
		SetText("[yellow]You can use a full absolute path or a relative path\n(e.g., www/default.asp or ./myscript.asp)")
	helpText.SetBackgroundColor(dialogBg)

	formBox := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(form, 7, 1, true).
		AddItem(helpText, 2, 1, false)
	formBox.SetBackgroundColor(dialogBg)

	// Add 1-point padding via Flex
	paddedForm := tview.NewFlex().SetDirection(tview.FlexRow).
		AddItem(nil, 1, 1, false).
		AddItem(tview.NewFlex().SetDirection(tview.FlexColumn).
			AddItem(nil, 2, 1, false).
			AddItem(formBox, 0, 1, true).
			AddItem(nil, 2, 1, false), 0, 1, true).
		AddItem(nil, 1, 1, false)
	paddedForm.SetBackgroundColor(dialogBg)

	form.SetFocus(0)

	// Center the form
	flex := tview.NewFlex().
		AddItem(nil, 0, 1, false).
		AddItem(tview.NewFlex().SetDirection(tview.FlexRow).
			AddItem(nil, 0, 1, false).
			AddItem(paddedForm, 12, 1, true).
			AddItem(nil, 0, 10, false), 66, 1, true).
		AddItem(nil, 0, 1, false)

	pages.AddPage("runfile", flex, true, true)
}

// runTUIFile executes an ASP file and prints results to outputArea.
func runTUIFile(path string, outputArea *tview.TextView) {
	outputArea.Clear()
	resolvedPath, err := resolveScriptPath(path)
	if err != nil {
		fmt.Fprintf(outputArea, "[red]Error: %v[white]\n", err)
		return
	}

	if !isASPExecutionExtension(resolvedPath) {
		fmt.Fprintf(outputArea, "[red]Error: %s: %s[white]\n", axonvm.ErrExtensionNotAllowed.String(), resolvedPath)
		return
	}

	virtualPath := scriptPathToVirtualPath(resolvedPath)
	result := executeCLIFile(resolvedPath, virtualPath, true)

	if result.compileErr != nil {
		fmt.Fprintf(outputArea, "[red]%s: %v[white]\n", axonvm.ErrCompileError.String(), result.compileErr)
		return
	}

	if result.output != "" {
		fmt.Fprint(outputArea, result.output)
	}

	if result.runtimeErr != nil {
		fmt.Fprintf(outputArea, "\n[red]%s: %v[white]\n", axonvm.ErrRuntimeError.String(), result.runtimeErr)
	}
}

// runCode executes the ASP code and updates the output area.
func runCode(code string, outputArea *tview.TextView) {
	outputArea.Clear()
	if strings.TrimSpace(code) == "" {
		return
	}

	// Wrap in ASP tags if not present
	source := code
	if !strings.Contains(source, "<%") && !strings.Contains(strings.ToLower(source), "<script") {
		source = "<%\n" + source + "\n%>"
	}

	result := executeCLICode(source, "/cli.asp", true)

	if result.compileErr != nil {
		fmt.Fprintf(outputArea, "[red]%s: %v[white]\n", axonvm.ErrCompileError.String(), result.compileErr)
		return
	}

	if result.output != "" {
		fmt.Fprint(outputArea, result.output)
	}

	if result.runtimeErr != nil {
		fmt.Fprintf(outputArea, "\n[red]%s: %v[white]\n", axonvm.ErrRuntimeError.String(), result.runtimeErr)
	}
}

// resolveScriptPath converts a user-provided path into an absolute existing file path.
func resolveScriptPath(inputPath string) (string, error) {
	absolutePath := inputPath
	if !filepath.IsAbs(absolutePath) {
		workingDir, err := os.Getwd()
		if err != nil {
			return "", axonvm.NewAxonASPError(axonvm.ErrCouldNotResolveCurrentDir, err, axonvm.ErrCouldNotResolveCurrentDir.String(), inputPath, 0)
		}
		absolutePath = filepath.Join(workingDir, inputPath)
	}

	absolutePath = filepath.Clean(absolutePath)
	info, err := os.Stat(absolutePath)
	if err != nil {
		if os.IsNotExist(err) {
			return "", axonvm.NewAxonASPError(axonvm.ErrFileNotFound, err, axonvm.ErrFileNotFound.String(), absolutePath, 0)
		}
		return "", axonvm.NewAxonASPError(axonvm.ErrCouldNotReadFile, err, axonvm.ErrCouldNotReadFile.String(), absolutePath, 0)
	}
	if info.IsDir() {
		return "", axonvm.NewAxonASPError(axonvm.ErrPathIsADirectory, nil, axonvm.ErrPathIsADirectory.String(), absolutePath, 0)
	}

	return absolutePath, nil
}

// scriptPathToVirtualPath maps a filesystem path into a web-style request path for CLI execution.
func scriptPathToVirtualPath(scriptPath string) string {
	workingDir, err := os.Getwd()
	if err != nil || strings.TrimSpace(workingDir) == "" {
		return "/" + filepath.ToSlash(filepath.Base(scriptPath))
	}

	serverRootDir := resolveCLIServerRootDir(workingDir)
	if relToRoot, relErr := filepath.Rel(serverRootDir, scriptPath); relErr == nil {
		cleanRootRelative := filepath.ToSlash(filepath.Clean(relToRoot))
		if cleanRootRelative != "." && cleanRootRelative != "" && !strings.HasPrefix(cleanRootRelative, "../") && cleanRootRelative != ".." {
			return "/" + strings.TrimPrefix(cleanRootRelative, "/")
		}
	}

	relativePath, err := filepath.Rel(workingDir, scriptPath)
	if err != nil {
		return "/" + filepath.ToSlash(filepath.Base(scriptPath))
	}

	cleanRelative := filepath.ToSlash(filepath.Clean(relativePath))
	if strings.HasPrefix(cleanRelative, "../") || cleanRelative == ".." {
		return "/" + filepath.ToSlash(filepath.Base(scriptPath))
	}

	return "/" + strings.TrimPrefix(cleanRelative, "/")
}

// resolveCLIServerRootDir resolves the CLI web root to an absolute path.
func resolveCLIServerRootDir(workingDir string) string {
	root := strings.TrimSpace(CLIServerRoot)
	if root == "" {
		return workingDir
	}
	if filepath.IsAbs(root) {
		return filepath.Clean(root)
	}
	return filepath.Clean(filepath.Join(workingDir, root))
}

// cliExecutionResult stores the output and error results from one CLI code execution.
type cliExecutionResult struct {
	output     string
	compileErr error
	runtimeErr error
}

// executeCLICode compiles and executes ASP source code with the provided virtual path.
func executeCLICode(code string, virtualPath string, tuiMode bool) cliExecutionResult {
	result := cliExecutionResult{}

	compiler := axonvm.NewASPCompiler(code)
	if strings.TrimSpace(virtualPath) != "" {
		compiler.SetSourceName(virtualPath)
	}

	err := compiler.Compile()
	if err != nil {
		result.compileErr = err
		return result
	}

	vm := axonvm.AcquireVMFromCompiler(compiler)

	var outBuf bytes.Buffer
	host := newCLIHost(&outBuf, virtualPath, tuiMode)
	vm.SetHost(host)
	// Set execution mode to bypass caching for interactive REPL input
	if tuiMode {
		vm.SetExecutionMode(axonvm.ExecutionModeTUI)
	} else {
		vm.SetExecutionMode(axonvm.ExecutionModeCLI)
	}
	defer vm.Release()
	wireCLIObjectAliases(vm, compiler)

	runErr := vm.Run()
	host.Response().Flush()

	result.output = outBuf.String()
	result.runtimeErr = runErr
	return result
}

// executeCLIFile executes one ASP file using the shared bytecode cache flow.
func executeCLIFile(filePath string, virtualPath string, tuiMode bool) cliExecutionResult {
	result := cliExecutionResult{}
	if scriptCache == nil {
		scriptCache = axonvm.NewScriptCache(axonvm.BytecodeCacheDisabled, filepath.Join("temp", "cache"), 1)
	}

	// Determine execution mode based on context (TUI or CLI interactive)
	executionMode := axonvm.ExecutionModeCLI
	if tuiMode {
		executionMode = axonvm.ExecutionModeTUI
	}

	program, err := scriptCache.LoadOrCompileWithMode(filePath, executionMode)
	if err != nil {
		result.compileErr = err
		return result
	}

	vm := axonvm.AcquireVMFromCachedProgram(program)

	var outBuf bytes.Buffer
	host := newCLIHost(&outBuf, virtualPath, tuiMode)
	vm.SetHost(host)
	// Set execution mode to bypass caching for interactive file execution
	vm.SetExecutionMode(executionMode)
	defer vm.Release()

	if len(program.GlobalNames) > 0 {
		for idx, name := range program.GlobalNames {
			if strings.EqualFold(strings.TrimSpace(name), "Document") && idx >= 0 && idx < len(vm.Globals) {
				vm.Globals[idx] = axonvm.Value{Type: axonvm.VTNativeObject, Num: 0}
				break
			}
		}
	}

	runErr := vm.Run()
	host.Response().Flush()

	result.output = outBuf.String()
	result.runtimeErr = runErr
	return result
}

// newCLIHost creates a fresh ASP host with CLI-friendly Server and Request context.
func newCLIHost(out *bytes.Buffer, requestPath string, tuiMode bool) *axonvm.MockHost {
	host := axonvm.NewMockHost()
	host.SetOutput(out)
	host.Response().SetMaxBufferBytes(ResponseBufferLimitBytes)

	host.SetApplication(sharedCLIApplication)
	host.SetSession(sharedCLISession)

	workingDir, err := os.Getwd()
	if err != nil || strings.TrimSpace(workingDir) == "" {
		workingDir = "."
	}
	serverRootDir := resolveCLIServerRootDir(workingDir)

	if strings.TrimSpace(requestPath) == "" {
		requestPath = "/cli.asp"
	}

	host.Server().SetRootDir(serverRootDir)
	host.Server().SetRequestPath(requestPath)
	_ = host.Server().SetScriptTimeout(ScriptTimeout)
	host.Request().ServerVars.Add("REQUEST_METHOD", "CLI")
	host.Request().ServerVars.Add("URL", requestPath)
	host.Request().ServerVars.Add("PATH_TRANSLATED", filepath.Join(serverRootDir, filepath.FromSlash(strings.TrimPrefix(requestPath, "/"))))
	if tuiMode {
		host.Request().ServerVars.Add("AXONASP_CLI_TUI", "1")
	}

	return host
}

// wireCLIObjectAliases maps CLI-friendly object aliases to ASP intrinsic objects.
func wireCLIObjectAliases(vm *axonvm.VM, compiler *axonvm.Compiler) {
	if vm == nil || compiler == nil || compiler.Globals == nil {
		return
	}

	if idx, exists := compiler.Globals.Get("Document"); exists && idx >= 0 && idx < len(vm.Globals) {
		vm.Globals[idx] = axonvm.Value{Type: axonvm.VTNativeObject, Num: 0}
	}
}

// printHelp shows the usage instructions for the CLI.
func printHelp() {
	fmt.Println("\033[1mG3pix ❖ AxonASP CLI Usage:\n\033[0m")
	fmt.Println(`  axonasp-cli 
    Starts the interactive REPL.
  axonasp-cli -r,--run <file>
    Runs the specified ASP file directly and returns only its output
  axonasp-cli -h, --help
    Shows this help message.
	
 ABOUT:
  G3pix ❖ AxonASP
  is a high-performance, cross-platform Classic ASP engine,
  with support to VBScript and JScript for Web, FastCGI, and CLI, 
  bridging legacy compatibility with modern APIs.
  
  Copyright (C) 2026 G3pix Ltda. All rights reserved.
  Website: https://g3pix.com.br/axonasp
  
  License: MPL 2.0

  `)
	fmt.Println("\033[0m")
}

// runDirectFile executes an ASP file directly from the command line without REPL prompts.
func runDirectFile(path string) {
	resolvedPath, err := resolveScriptPath(path)
	if err != nil {
		fmt.Printf("Error: %v\n", err)
		os.Exit(1)
	}

	if !isASPExecutionExtension(resolvedPath) {
		fmt.Printf("Error: %s: %s\n", axonvm.ErrExtensionNotAllowed.String(), resolvedPath)
		os.Exit(1)
	}

	virtualPath := scriptPathToVirtualPath(resolvedPath)
	result := executeCLIFile(resolvedPath, virtualPath, false)

	if result.compileErr != nil {
		fmt.Fprintf(os.Stderr, "%s: %v\n", axonvm.ErrCompileError.String(), result.compileErr)
		os.Exit(1)
	}

	if result.output != "" {
		fmt.Print(result.output)

	}

	if result.runtimeErr != nil {
		fmt.Fprintf(os.Stderr, "%s: %v\n", axonvm.ErrRuntimeError.String(), result.runtimeErr)
		os.Exit(1)
	}
}
