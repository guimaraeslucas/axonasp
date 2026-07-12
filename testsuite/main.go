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
//go:generate goversioninfo -icon=icon_testsuite.ico -64=true
package main

import (
	"bytes"
	"fmt"
	"io/fs"
	"os"
	"path/filepath"
	"runtime/debug"
	"slices"
	"sort"
	"strings"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/axonvm"
	"g3pix.com.br/axonasp/axonvm/asp"
	"github.com/joho/godotenv"
	"github.com/spf13/pflag"
)

// Version is injected by the build scripts.
var Version = "0.0.0.0"

// Testsuite configuration values.
var (
	DefaultTimezone               = "UTC"
	MemoryLimitMB                 = 128
	VMPoolSize                    = 50
	BytecodeCachingMode           = "enabled"
	CacheMaxSizeMB                = 128
	ScriptTimeout                 = 60
	ResponseBufferLimitBytes      = 4 * 1024 * 1024
	ExecuteAsASPExtensions        = []string{".asp"}
	ExecuteAsVBScriptExtensions   = []string{".vbs"}
	ExecuteAsJavaScriptExtensions = []string{".js", ".mjs"}
	SuiteEngineMode               = axonvm.EngineModeDefault
	CLIServerRoot                 = "./www"
	TempDir                       = filepath.Join(".", "temp")
	scriptCache                   *axonvm.ScriptCache

	sharedApplication = asp.NewApplication()
	sharedSession     = asp.NewSession()

	testsuiteConfigFilePath string
	testsuiteAboutFlag      bool
	testsuiteTargetPath     string
)

const (
	colorReset  = "\033[0m"
	colorRed    = "\033[31m"
	colorGreen  = "\033[32m"
	colorYellow = "\033[33m"
	colorCyan   = "\033[36m"
)

// suiteResult captures one executed ASP test file and its aggregate outcome.
type suiteResult struct {
	filePath    string
	virtualPath string
	output      string
	compileErr  error
	runtimeErr  error
	summary     axonvm.G3TestSuiteSummary
	reports     []axonvm.G3TestReport
}

// init loads environment variables and applies TOML-based configuration through Viper.
func init() {
	_ = godotenv.Load()
	axonvm.SetRuntimeVersion(strings.TrimSpace(Version))
	configureTestSuiteFlags()
	loadConfig()
	applyRuntimeSettings()
}

// configureTestSuiteFlags defines and parses testsuite command-line flags.
func configureTestSuiteFlags() {
	pflag.Usage = func() {
		fmt.Printf("G3pix ❖ AxonASP TestSuite %s\n", Version)
		fmt.Println("Options available: ")
		pflag.PrintDefaults()
		fmt.Print("\nFor more information, visit: https://g3pix.com.br/axonasp/manual/\n")
	}

	pflag.StringVarP(&testsuiteConfigFilePath, "config.config_file", "c", "", "Path to the AxonASP TOML configuration file to load before running suite discovery and execution.")
	pflag.BoolVarP(&testsuiteAboutFlag, "about", "a", false, "Print AxonASP product and licensing information, then exit.")

	pflag.Parse()

	if testsuiteAboutFlag {
		fmt.Print(axonconfig.AboutG3pixAxonASP())
		os.Exit(0)
	}

	if strings.TrimSpace(testsuiteConfigFilePath) != "" {
		axonconfig.SetCustomConfigPath(testsuiteConfigFilePath)
	}

	if args := pflag.Args(); len(args) > 0 {
		testsuiteTargetPath = args[0]
	}
}

// main scans the requested directory, executes ASP tests, and returns a process exit code based on failures.
func main() {
	workingDir, _ := os.Getwd()
	scriptCache = axonvm.NewScriptCache(
		axonvm.ParseBytecodeCacheMode(BytecodeCachingMode),
		filepath.Join(TempDir, "cache"),
		CacheMaxSizeMB,
	)
	scriptCache.SetWatchedExtensions(ExecuteAsASPExtensions)
	if err := scriptCache.StartInvalidator([]string{workingDir}); err != nil {
		fmt.Printf("%sWarning%s: failed to start bytecode invalidator: %v\n", colorYellow, colorReset, err)
	}
	defer scriptCache.StopInvalidator()

	if err := loadGlobalASA(workingDir); err != nil {
		fmt.Printf("%sWarning%s: failed to load global.asa: %v\n", colorYellow, colorReset, err)
	}
	defer shutdownGlobalASA()

	if strings.TrimSpace(testsuiteTargetPath) == "" {
		pflag.Usage()
		os.Exit(1)
	}

	targetDir, err := resolveDirectoryPath(testsuiteTargetPath)
	if err != nil {
		fmt.Fprintf(os.Stderr, "%sError%s: %v\n", colorRed, colorReset, err)
		os.Exit(1)
	}

	files, err := findTestFiles(targetDir)
	if err != nil {
		fmt.Fprintf(os.Stderr, "%sError%s: %v\n", colorRed, colorReset, err)
		os.Exit(1)
	}
	if len(files) == 0 {
		fmt.Printf("%sNo test files found%s under %s\n", colorYellow, colorReset, targetDir)
		os.Exit(1)
	}

	results := make([]suiteResult, 0, len(files))
	var totalSuites int
	var totalAssertions int64
	var totalPassed int64
	var totalFailed int64
	var executionFailures int

	for _, filePath := range files {
		result := executeSuiteFile(filePath)
		results = append(results, result)
		totalSuites += result.summary.SuiteCount
		totalAssertions += result.summary.Total
		totalPassed += result.summary.Passed
		totalFailed += result.summary.Failed
		if result.compileErr != nil || result.runtimeErr != nil || result.summary.Failed > 0 || (result.summary.Total == 0 && len(result.reports) == 0) {
			executionFailures++
		}
	}

	printResults(results, totalSuites, totalAssertions, totalPassed, totalFailed)
	if executionFailures > 0 || totalFailed > 0 {
		os.Exit(1)
	}
	os.Exit(0)
}

// loadConfig loads runtime configuration shared with the existing CLI execution path.
func loadConfig() {
	v := axonconfig.NewViper()
	if strings.TrimSpace(v.ConfigFileUsed()) == "" {
		fmt.Printf("%sWarning%s: %s\n", colorYellow, colorReset, axonvm.ErrViperReadConfigFailed.String())
	}
	axonconfig.EnableWatchIfConfigured(v, nil)

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
	if tempDir := strings.TrimSpace(v.GetString("global.temp_dir")); tempDir != "" {
		TempDir = filepath.Clean(tempDir)
	}
	if executeAsASP := v.GetStringSlice("global.execute_as_asp"); len(executeAsASP) > 0 {
		ExecuteAsASPExtensions = normalizeExtensions(executeAsASP)
	}
	if executeAsVBS := v.GetStringSlice("global.execute_as_vbscript"); len(executeAsVBS) > 0 {
		ExecuteAsVBScriptExtensions = normalizeExtensions(executeAsVBS)
	}
	if executeAsJS := v.GetStringSlice("global.execute_as_javascript"); len(executeAsJS) > 0 {
		ExecuteAsJavaScriptExtensions = normalizeExtensions(executeAsJS)
	}

	mode := strings.ToLower(strings.TrimSpace(v.GetString("testsuite.engine_mode")))
	switch mode {
	case "vbscript":
		SuiteEngineMode = axonvm.EngineModeVBScript
	case "javascript":
		SuiteEngineMode = axonvm.EngineModeJavaScript
	default:
		SuiteEngineMode = axonvm.EngineModeDefault
	}

	if webRoot := strings.TrimSpace(v.GetString("server.web_root")); webRoot != "" {
		CLIServerRoot = webRoot
	}
	if timeout := v.GetInt("global.default_script_timeout"); timeout > 0 {
		ScriptTimeout = timeout
	}
	if responseBufferLimitMB := v.GetInt("global.response_buffer_limit_mb"); responseBufferLimitMB > 0 {
		ResponseBufferLimitBytes = responseBufferLimitMB * 1024 * 1024
	}

	axonvm.SetVMPoolSizeLimit(VMPoolSize)
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
		fmt.Printf("%sWarning%s: %s %s, using UTC: %v\n", colorYellow, colorReset, axonvm.ErrInvalidTimezone.String(), DefaultTimezone, err)
		location = time.UTC
	}
	time.Local = location
	axonvm.ReloadBuiltinDefaults()
}

// loadGlobalASA primes application and session static objects before the test suite starts.
func loadGlobalASA(workingDir string) error {
	if err := axonvm.GetGlobalASA().LoadAndCompile(workingDir, sharedApplication); err != nil {
		return err
	}
	if !axonvm.GetGlobalASA().IsLoaded() {
		return nil
	}
	dummyHost := newTestsuiteHost(new(bytes.Buffer), "/")
	if err := axonvm.GetGlobalASA().ExecuteApplicationOnStart(dummyHost); err != nil {
		return err
	}
	axonvm.GetGlobalASA().PopulateSessionStaticObjects(sharedSession)
	return axonvm.GetGlobalASA().ExecuteSessionOnStart(dummyHost)
}

// shutdownGlobalASA executes final application and session shutdown hooks when global.asa is loaded.
func shutdownGlobalASA() {
	if !axonvm.GetGlobalASA().IsLoaded() {
		return
	}
	dummyHost := newTestsuiteHost(new(bytes.Buffer), "/")
	_ = axonvm.GetGlobalASA().ExecuteSessionOnEnd(dummyHost)
	_ = axonvm.GetGlobalASA().ExecuteApplicationOnEnd(dummyHost)
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
		duplicate := slices.Contains(cleaned, ext)
		if !duplicate {
			cleaned = append(cleaned, ext)
		}
	}
	return cleaned
}

// resolveDirectoryPath converts a user-provided directory path into an absolute existing directory path.
func resolveDirectoryPath(inputPath string) (string, error) {
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
	if !info.IsDir() {
		return "", fmt.Errorf("path is not a directory: %s", absolutePath)
	}

	return absolutePath, nil
}

// findTestFiles recursively returns ASP files matching *test.asp or test_*.asp in stable order.
func findTestFiles(root string) ([]string, error) {
	files := make([]string, 0, 16)
	err := filepath.WalkDir(root, func(path string, entry fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() {
			return nil
		}
		if !isASPExecutionExtension(path) {
			return nil
		}
		name := strings.ToLower(entry.Name())
		if strings.HasSuffix(name, "test.asp") || strings.HasPrefix(name, "test_") {
			files = append(files, filepath.Clean(path))
		}
		return nil
	})
	if err != nil {
		return nil, err
	}
	sort.Strings(files)
	return files, nil
}

// isASPExecutionExtension reports whether a file should be executed based on the current engine mode.
func isASPExecutionExtension(filePath string) bool {
	ext := strings.ToLower(filepath.Ext(filePath))
	switch SuiteEngineMode {
	case axonvm.EngineModeVBScript:
		return containsFold(ExecuteAsVBScriptExtensions, ext)
	case axonvm.EngineModeJavaScript:
		return containsFold(ExecuteAsJavaScriptExtensions, ext)
	default:
		return containsFold(ExecuteAsASPExtensions, ext)
	}
}

func containsFold(slice []string, val string) bool {
	for _, s := range slice {
		if strings.EqualFold(s, val) {
			return true
		}
	}
	return false
}

// executeSuiteFile loads, runs, and extracts the G3Test summary for one ASP test file.
func executeSuiteFile(filePath string) suiteResult {
	result := suiteResult{filePath: filePath, virtualPath: scriptPathToVirtualPath(filePath)}
	program, err := scriptCache.LoadOrCompile(filePath)
	if err != nil {
		result.compileErr = err
		return result
	}

	vm := axonvm.AcquireVMFromCachedProgram(program)
	defer vm.Release()

	var outBuf bytes.Buffer
	host := newTestsuiteHost(&outBuf, result.virtualPath)
	vm.SetHost(host)
	wireObjectAliases(vm, program)

	runErr := vm.Run()
	host.Response().Flush()

	result.output = outBuf.String()
	result.runtimeErr = runErr
	result.reports = vm.GetG3TestReports()
	result.summary = vm.GetG3TestSuiteSummary()
	return result
}

// newTestsuiteHost creates a fresh ASP host with CLI-style Server and Request context for test execution.
func newTestsuiteHost(out *bytes.Buffer, requestPath string) *axonvm.MockHost {
	host := axonvm.NewMockHost()
	host.SetOutput(out)
	host.Response().SetMaxBufferBytes(ResponseBufferLimitBytes)
	host.SetApplication(sharedApplication)
	host.SetSession(sharedSession)

	workingDir, err := os.Getwd()
	if err != nil || strings.TrimSpace(workingDir) == "" {
		workingDir = "."
	}
	serverRootDir := resolveServerRootDir(workingDir)
	if strings.TrimSpace(requestPath) == "" {
		requestPath = "/testsuite.asp"
	}

	host.Server().SetRootDir(serverRootDir)
	host.Server().SetRequestPath(requestPath)
	_ = host.Server().SetScriptTimeout(ScriptTimeout)
	host.Request().ServerVars.Add("REQUEST_METHOD", "CLI")
	host.Request().ServerVars.Add("URL", requestPath)
	host.Request().ServerVars.Add("PATH_TRANSLATED", filepath.Join(serverRootDir, filepath.FromSlash(strings.TrimPrefix(requestPath, "/"))))
	return host
}

// wireObjectAliases preserves CLI-compatible global aliases during cached execution.
func wireObjectAliases(vm *axonvm.VM, program axonvm.CachedProgram) {
	if len(program.GlobalNames) == 0 {
		return
	}
	for idx := range program.GlobalNames {
		if strings.EqualFold(strings.TrimSpace(program.GlobalNames[idx]), "Document") && idx >= 0 && idx < len(vm.Globals) {
			vm.Globals[idx] = axonvm.Value{Type: axonvm.VTNativeObject, Num: 0}
			break
		}
	}
}

// scriptPathToVirtualPath maps a filesystem path into a web-style request path for testsuite execution.
func scriptPathToVirtualPath(scriptPath string) string {
	workingDir, err := os.Getwd()
	if err != nil || strings.TrimSpace(workingDir) == "" {
		return "/" + filepath.ToSlash(filepath.Base(scriptPath))
	}

	serverRootDir := resolveServerRootDir(workingDir)
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

// resolveServerRootDir resolves the CLI web root to an absolute path.
func resolveServerRootDir(workingDir string) string {
	root := strings.TrimSpace(CLIServerRoot)
	if root == "" {
		return workingDir
	}
	if filepath.IsAbs(root) {
		return filepath.Clean(root)
	}
	return filepath.Clean(filepath.Join(workingDir, root))
}

// printResults prints one PHPUnit-style execution report with ANSI coloring.
func printResults(results []suiteResult, totalSuites int, totalAssertions int64, totalPassed int64, totalFailed int64) {
	for _, result := range results {
		statusColor := colorGreen
		statusLabel := "PASS"
		if result.compileErr != nil || result.runtimeErr != nil || result.summary.Failed > 0 || (result.summary.Total == 0 && len(result.reports) == 0) {
			statusColor = colorRed
			statusLabel = "FAIL"
		}

		fmt.Printf("%s[%s]%s %s\n", statusColor, statusLabel, colorReset, result.filePath)
		if result.compileErr != nil {
			fmt.Printf("  %sCompile%s: %v\n", colorRed, colorReset, result.compileErr)
			continue
		}
		if result.runtimeErr != nil {
			fmt.Printf("  %sRuntime%s: %v\n", colorRed, colorReset, result.runtimeErr)
		}
		if result.summary.Total == 0 && len(result.reports) == 0 {
			fmt.Printf("  %sNo G3Test assertions executed%s\n", colorYellow, colorReset)
		}
		for _, failure := range result.summary.Failures {
			fmt.Printf("  %s- %s%s\n", colorRed, failure.Message, colorReset)
		}
		if strings.TrimSpace(result.output) != "" {
			fmt.Printf("  %sOutput%s:\n%s\n", colorCyan, colorReset, indentBlock(result.output, "    "))
		}
		fmt.Printf("  Assertions: %d, Passed: %d, Failed: %d\n", result.summary.Total, result.summary.Passed, result.summary.Failed)
	}

	statusColor := colorGreen
	statusLabel := "PASS"
	if totalFailed > 0 {
		statusColor = colorRed
		statusLabel = "FAIL"
	}
	fmt.Printf("\n%s%s%s Suites: %d  Assertions: %d  Passed: %d  Failed: %d\n", statusColor, statusLabel, colorReset, totalSuites, totalAssertions, totalPassed, totalFailed)
}

// indentBlock indents multi-line output for readable terminal summaries.
func indentBlock(text string, prefix string) string {
	trimmed := strings.TrimRight(text, "\r\n")
	if trimmed == "" {
		return prefix
	}
	lines := strings.Split(trimmed, "\n")
	for i := range lines {
		lines[i] = prefix + strings.TrimRight(lines[i], "\r")
	}
	return strings.Join(lines, "\n")
}
