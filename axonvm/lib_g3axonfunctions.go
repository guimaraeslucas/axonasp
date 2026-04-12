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
	"bytes"
	"crypto/md5"
	"crypto/sha1"
	"crypto/sha256"
	"encoding/base64"
	"encoding/csv"
	"fmt"
	"html"
	"io"
	"math"
	"math/rand"
	"mime"
	"net"
	"net/http"
	"net/mail"
	"net/url"
	"os"
	"os/exec"
	"os/user"
	"path/filepath"
	"regexp"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"time"

	"g3pix.com.br/axonasp/axonconfig"
	"github.com/spf13/viper"
)

type AxonLibrary struct {
	vm *VM
}

type axFunctionConfig struct {
	enableServerShutdown bool
	defaultCSSPath       string
	defaultLogoPath      string
}

var axFunctionConfigOnce sync.Once
var cachedAxFunctionConfig axFunctionConfig

// filterEnvironmentEntries removes pseudo/internal environment keys (e.g., Windows '=C:=...') from os.Environ output.
func filterEnvironmentEntries(envList []string) []string {
	if len(envList) == 0 {
		return envList
	}
	filtered := make([]string, 0, len(envList))
	for i := 0; i < len(envList); i++ {
		entry := envList[i]
		if entry == "" {
			continue
		}
		if entry[0] == '=' {
			continue
		}
		filtered = append(filtered, entry)
	}
	return filtered
}

// AxonGlobalFunctionNames contains Ax* functions that can be optionally exposed as global VBScript built-ins.
var AxonGlobalFunctionNames = []string{
	"axenginename",
	"axversion",
	"axruntimeinfo",
	"axgetenv",
	"axgetconfig",
	"axgetconfigkeys",
	"axuserhomedirpath",
	"axuserconfigdirpath",
	"axcachedirpath",
	"axispathseparator",
	"axchangetimes",
	"axchangemode",
	"axcreatelink",
	"axchangeowner",
	"axshutdownaxonaspserver",
	"axchangedir",
	"axcurrentdir",
	"axhostnamevalue",
	"axclearenvironment",
	"axenvironmentlist",
	"axenvironmentvalue",
	"axprocessid",
	"axeffectiveuserid",
	"axdirseparator",
	"axpathlistseparator",
	"axintegersizebytes",
	"axplatformbits",
	"axexecutablepath",
	"axexecute",
	"axsysteminfo",
	"axcurrentuser",
	"axw",
	"axmax",
	"axmin",
	"axintegermax",
	"axintegermin",
	"axceil",
	"axfloor",
	"axrand",
	"axnumberformat",
	"axpi",
	"axsmallestfloatvalue",
	"axfloatprecisiondigits",
	"axcount",
	"axexplode",
	"aximplode",
	"axarrayreverse",
	"axrange",
	"axstringreplace",
	"axpad",
	"axrepeat",
	"axucfirst",
	"axwordcount",
	"axnl2br",
	"axtrim",
	"axstringgetcsv",
	"axmd5",
	"axsha1",
	"axhash",
	"axbase64encode",
	"axbase64decode",
	"axurldecode",
	"axrawurldecode",
	"axrgbtohex",
	"axhextorgb",
	"axgetlogo",
	"axhtmlspecialchars",
	"axstriptags",
	"axfiltervalidateip",
	"axfiltervalidateemail",
	"axisint",
	"axisfloat",
	"axctypealpha",
	"axctypealnum",
	"axempty",
	"axisset",
	"axtime",
	"axdate",
	"axlastmodified",
	"axgetremotefile",
	"axgenerateguid",
	"axgetdefaultcss",
}

// AxonGlobalFunctionPointers maps 1:1 to AxonGlobalFunctionNames and routes directly to AxonLibrary dispatch logic.
var AxonGlobalFunctionPointers = buildAxonGlobalFunctionPointers(AxonGlobalFunctionNames)

// buildAxonGlobalFunctionPointers compiles a stable list of VM-aware built-ins from Axon method names.
func buildAxonGlobalFunctionPointers(names []string) []BuiltinFunc {
	pointers := make([]BuiltinFunc, 0, len(names))
	for _, name := range names {
		pointers = append(pointers, axonGlobalBuiltin(name))
	}
	return pointers
}

// axonGlobalBuiltin adapts one AxonLibrary method as a global built-in while reusing the same implementation.
func axonGlobalBuiltin(methodName string) BuiltinFunc {
	return func(vm *VM, args []Value) (Value, error) {
		if vm == nil {
			return NewEmpty(), nil
		}
		lib := &AxonLibrary{vm: vm}
		return lib.DispatchMethod(methodName, args), nil
	}
}

// loadAxFunctionConfig reads Ax function toggles from config/axonasp.toml.
func loadAxFunctionConfig() axFunctionConfig {
	axFunctionConfigOnce.Do(func() {
		cfg := axFunctionConfig{}
		v := axonconfig.NewViper()
		cfg.enableServerShutdown = v.GetBool("axfunctions.enable_axservershutdown_function")
		cfg.defaultCSSPath = v.GetString("axfunctions.ax_default_css_path")
		cfg.defaultLogoPath = v.GetString("axfunctions.ax_default_logo_path")
		cachedAxFunctionConfig = cfg
	})
	return cachedAxFunctionConfig
}

// newAxonConfigViper creates a Viper instance loaded with axonasp.toml.
// If global.viper_automatic_env is true in the config, it also activates
// environment variable overrides using dot-to-underscore key replacement.
// Returns the Viper instance and true when the config file was found and read.
func newAxonConfigViper() (*viper.Viper, bool) {
	v := axonconfig.NewViper()
	return v, strings.TrimSpace(v.ConfigFileUsed()) != ""
}

// loadAxConfigValue reads one configuration key from axonasp.toml using Viper.
// When global.viper_automatic_env is enabled, environment variables take precedence over file values.
func loadAxConfigValue(configKey string) (interface{}, bool) {
	key := strings.TrimSpace(configKey)
	if key == "" {
		return nil, false
	}

	v, loaded := newAxonConfigViper()
	if !loaded {
		return nil, false
	}

	if !v.IsSet(key) {
		return nil, false
	}

	return v.Get(key), true
}

// loadAllAxConfigKeys returns every configuration key present in axonasp.toml.
// The returned slice is sorted alphabetically.
func loadAllAxConfigKeys() []string {
	v, loaded := newAxonConfigViper()
	if !loaded {
		return nil
	}
	return v.AllKeys()
}

// axonConfigCandidates returns known locations for config/axonasp.toml.
func axonConfigCandidates() []string {
	candidates := []string{
		filepath.Join("config", "axonasp.toml"),
		filepath.Join("..", "config", "axonasp.toml"),
	}
	if ex, exErr := os.Executable(); exErr == nil {
		candidates = append(candidates, filepath.Join(filepath.Dir(ex), "config", "axonasp.toml"))
	}
	return candidates
}

// resolveAxonConfigPath returns the first existing config path, or a stable absolute fallback.
func resolveAxonConfigPath() string {
	candidates := axonConfigCandidates()
	for i := 0; i < len(candidates); i++ {
		candidate := candidates[i]
		if _, err := os.Stat(candidate); err == nil {
			if abs, absErr := filepath.Abs(candidate); absErr == nil {
				return abs
			}
			return candidate
		}
	}
	if abs, err := filepath.Abs(filepath.Join("config", "axonasp.toml")); err == nil {
		return abs
	}
	return filepath.Join("config", "axonasp.toml")
}

// axRuntimeBanner returns the AxonASP legal and attribution text block for diagnostics output.
func axRuntimeBanner() string {
	return "AxonASP Server\n" +
		"Copyright (C) 2026 G3pix Ltda. All rights reserved.\n" +
		"Developed by Lucas Guimarães - G3pix Ltda\n\n" +
		"Project URL: https://g3pix.com.br/axonasp\n\n" +
		"This Source Code Form is subject to the terms of the Mozilla Public\n" +
		"License, v. 2.0. If a copy of the MPL was not distributed with this\n" +
		"file, You can obtain one at https://mozilla.org/MPL/2.0/.\n" +
		"Attribution Notice:\n" +
		"If this software is used in other projects, the name \"AxonASP Server\"\n" +
		"must be cited in the documentation or \"About\" section.\n\n" +
		"Contribution Policy:\n" +
		"Modifications to the core source code of AxonASP Server must be\n" +
		"made available under this same license terms."
}

// buildAxRuntimeInfoReport renders a phpinfo-like text report with runtime and config details.
func buildAxRuntimeInfoReport(vm *VM) string {
	var b strings.Builder
	b.Grow(4096)

	now := time.Now()
	hostName, _ := os.Hostname()
	currentDir, _ := os.Getwd()
	execPath, _ := os.Executable()
	homeDir, _ := os.UserHomeDir()
	configPath := resolveAxonConfigPath()
	cacheDir := filepath.Join(".temp", "cache")
	cacheAbs, cacheErr := filepath.Abs(cacheDir)
	if cacheErr == nil {
		cacheDir = cacheAbs
	}
	b.WriteString(axRuntimeBanner())
	b.WriteString("\n\nAXONASP RUNTIME INFORMATION\n")
	b.WriteString("===========================\n")
	b.WriteString("Timestamp: ")
	b.WriteString(now.Format(time.RFC3339))
	b.WriteString("\n")
	b.WriteString("Engine: AxonASP\n")
	b.WriteString("Version: ")
	b.WriteString(GetRuntimeVersion())
	b.WriteString("\n")
	b.WriteString("Go Runtime: ")
	b.WriteString(runtime.Version())
	b.WriteString("\n")
	b.WriteString("Platform: ")
	b.WriteString(runtime.GOOS)
	b.WriteString("/")
	b.WriteString(runtime.GOARCH)
	b.WriteString("\n")
	b.WriteString("CPU Cores: ")
	b.WriteString(strconv.Itoa(runtime.NumCPU()))
	b.WriteString("\n")
	b.WriteString("GOMAXPROCS: ")
	b.WriteString(strconv.Itoa(runtime.GOMAXPROCS(0)))
	b.WriteString("\n\n")

	b.WriteString("SERVER CONTEXT\n")
	b.WriteString("--------------\n")
	b.WriteString("Hostname: ")
	b.WriteString(hostName)
	b.WriteString("\n")
	b.WriteString("PID: ")
	b.WriteString(strconv.Itoa(os.Getpid()))
	b.WriteString("\n")
	b.WriteString("Current Directory: ")
	b.WriteString(currentDir)
	b.WriteString("\n")
	b.WriteString("Executable Path: ")
	b.WriteString(execPath)
	b.WriteString("\n")
	b.WriteString("User Home Directory: ")
	b.WriteString(homeDir)
	b.WriteString("\n")
	b.WriteString("User Config File: ")
	b.WriteString(configPath)
	b.WriteString("\n")
	b.WriteString("Cache Directory: ")
	b.WriteString(cacheDir)
	b.WriteString("\n")
	b.WriteString("Path Separator: ")
	b.WriteString(string(os.PathSeparator))
	b.WriteString("\n")
	b.WriteString("Path List Separator: ")
	b.WriteString(string(os.PathListSeparator))
	b.WriteString("\n")
	b.WriteString("Integer Size (bits): ")
	b.WriteString(strconv.Itoa(strconv.IntSize))
	b.WriteString("\n\n")

	b.WriteString("MEMORY SNAPSHOT\n")
	b.WriteString("---------------\n")
	var ms runtime.MemStats
	runtime.ReadMemStats(&ms)
	b.WriteString("Alloc Bytes: ")
	b.WriteString(strconv.FormatUint(ms.Alloc, 10))
	b.WriteString("\n")
	b.WriteString("Total Alloc Bytes: ")
	b.WriteString(strconv.FormatUint(ms.TotalAlloc, 10))
	b.WriteString("\n")
	b.WriteString("Sys Bytes: ")
	b.WriteString(strconv.FormatUint(ms.Sys, 10))
	b.WriteString("\n")
	b.WriteString("Heap Objects: ")
	b.WriteString(strconv.FormatUint(ms.HeapObjects, 10))
	b.WriteString("\n")
	b.WriteString("Num GC: ")
	b.WriteString(strconv.FormatUint(uint64(ms.NumGC), 10))
	b.WriteString("\n\n")

	b.WriteString("CONFIGURATION (config/axonasp.toml)\n")
	b.WriteString("------------------------------------\n")
	v, loaded := newAxonConfigViper()
	if !loaded {
		b.WriteString("Config status: not loaded\n\n")
	} else {
		b.WriteString("Config status: loaded\n")
		keys := v.AllKeys()
		for i := 0; i < len(keys); i++ {
			key := keys[i]
			b.WriteString("- ")
			b.WriteString(key)
			b.WriteString(" = ")
			b.WriteString(fmt.Sprintf("%v", v.Get(key)))
			b.WriteString("\n")
		}
		b.WriteString("\n")
	}

	if vm != nil && vm.host != nil && vm.host.Server() != nil {
		b.WriteString("AXONASP SERVER SETTINGS\n")
		b.WriteString("-----------------------\n")
		b.WriteString("Script Timeout (seconds): ")
		b.WriteString(strconv.Itoa(vm.host.Server().GetScriptTimeout()))
		b.WriteString("\n")
	}

	b.WriteString("\n")
	return b.String()
}

// axConfigValueToVMValue converts Viper config values into VM-compatible Value types.
func axConfigValueToVMValue(raw interface{}) Value {
	switch v := raw.(type) {
	case nil:
		return NewEmpty()
	case string:
		return NewString(v)
	case bool:
		return NewBool(v)
	case int:
		return NewInteger(int64(v))
	case int8:
		return NewInteger(int64(v))
	case int16:
		return NewInteger(int64(v))
	case int32:
		return NewInteger(int64(v))
	case int64:
		return NewInteger(v)
	case uint:
		return NewInteger(int64(v))
	case uint8:
		return NewInteger(int64(v))
	case uint16:
		return NewInteger(int64(v))
	case uint32:
		return NewInteger(int64(v))
	case uint64:
		return NewInteger(int64(v))
	case float32:
		return NewDouble(float64(v))
	case float64:
		return NewDouble(v)
	case []string:
		values := make([]Value, len(v))
		for i := 0; i < len(v); i++ {
			values[i] = NewString(v[i])
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
	case []interface{}:
		values := make([]Value, len(v))
		for i := 0; i < len(v); i++ {
			values[i] = axConfigValueToVMValue(v[i])
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}
	default:
		return NewString(fmt.Sprintf("%v", raw))
	}
}

// newAxonLibrary instantiates the Axon intrinsic custom functions library.
func (vm *VM) newAxonLibrary() Value {
	obj := &AxonLibrary{vm: vm}
	id := vm.nextDynamicNativeID
	vm.nextDynamicNativeID++
	vm.axonItems[id] = obj
	return Value{Type: VTNativeObject, Num: id}
}

// DispatchPropertyGet acts as a getter. In VBScript, properties are frequently accessed like methods without args.
func (al *AxonLibrary) DispatchPropertyGet(propertyName string) Value {
	return al.DispatchMethod(propertyName, nil)
}

// DispatchMethod provides O(1) string matching resolution for all custom Ax* functions.
// We strictly use switch blocks and avoid reflection to maintain execution performance and zero allocations where possible.
func (al *AxonLibrary) DispatchMethod(methodName string, args []Value) Value {
	funcLower := strings.ToLower(methodName)

	switch funcLower {

	// --- System / Environment Functions ---

	case "axenginename":
		return NewString("AxonASP")

	case "axversion":
		return NewString(GetRuntimeVersion())

	case "axruntimeinfo":
		return NewString(buildAxRuntimeInfoReport(al.vm))

	case "axuserhomedirpath":
		homeDir, err := os.UserHomeDir()
		if err == nil && strings.TrimSpace(homeDir) != "" {
			return NewString(homeDir)
		}
		if currentUser, userErr := user.Current(); userErr == nil && currentUser != nil && strings.TrimSpace(currentUser.HomeDir) != "" {
			return NewString(currentUser.HomeDir)
		}
		if val := os.Getenv("USERPROFILE"); val != "" {
			return NewString(val)
		}
		if val := os.Getenv("HOME"); val != "" {
			return NewString(val)
		}
		return NewString("")

	case "axuserconfigdirpath":
		return NewString(resolveAxonConfigPath())

	case "axcachedirpath":
		cachePath := filepath.Join(".temp", "cache")
		if abs, err := filepath.Abs(cachePath); err == nil {
			cachePath = abs
		}
		if !strings.HasSuffix(cachePath, string(os.PathSeparator)) {
			cachePath += string(os.PathSeparator)
		}
		return NewString(cachePath)

	case "axispathseparator":
		if len(args) == 0 {
			return NewBool(false)
		}
		candidate := args[0].String()
		if candidate == "" {
			return NewBool(false)
		}
		runes := []rune(candidate)
		if len(runes) != 1 {
			return NewBool(false)
		}
		return NewBool(os.IsPathSeparator(uint8(runes[0])))

	case "axchangetimes":
		if len(args) < 3 {
			return NewBool(false)
		}
		path := args[0].String()
		if strings.TrimSpace(path) == "" {
			return NewBool(false)
		}
		atime := time.Unix(int64(al.vm.asInt(args[1])), 0)
		mtime := time.Unix(int64(al.vm.asInt(args[2])), 0)
		if err := os.Chtimes(path, atime, mtime); err != nil {
			return NewBool(false)
		}
		return NewBool(true)

	case "axchangemode":
		if len(args) < 2 {
			return NewBool(false)
		}
		path := args[0].String()
		modeText := strings.TrimSpace(args[1].String())
		if strings.TrimSpace(path) == "" || modeText == "" {
			return NewBool(false)
		}
		parsedMode, err := strconv.ParseUint(modeText, 8, 32)
		if err != nil {
			return NewBool(false)
		}
		if err := os.Chmod(path, os.FileMode(parsedMode)); err != nil {
			return NewBool(false)
		}
		return NewBool(true)

	case "axcreatelink":
		if len(args) < 2 {
			return NewBool(false)
		}
		src := args[0].String()
		dst := args[1].String()
		if strings.TrimSpace(src) == "" || strings.TrimSpace(dst) == "" {
			return NewBool(false)
		}
		if err := os.Link(src, dst); err != nil {
			return NewBool(false)
		}
		return NewBool(true)

	case "axchangeowner":
		if len(args) < 3 {
			return NewBool(false)
		}
		path := args[0].String()
		if strings.TrimSpace(path) == "" {
			return NewBool(false)
		}
		uid := al.vm.asInt(args[1])
		gid := al.vm.asInt(args[2])
		if err := os.Chown(path, uid, gid); err != nil {
			return NewBool(false)
		}
		return NewBool(true)

	case "axgetenv":
		if len(args) == 0 {
			return NewString("")
		}
		return NewString(os.Getenv(args[0].String()))

	case "axgetconfig":
		if len(args) == 0 {
			return NewEmpty()
		}
		raw, ok := loadAxConfigValue(args[0].String())
		if !ok {
			return NewEmpty()
		}
		return axConfigValueToVMValue(raw)

	case "axgetconfigkeys":
		keys := loadAllAxConfigKeys()
		if len(keys) == 0 {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		values := make([]Value, len(keys))
		for i := 0; i < len(keys); i++ {
			values[i] = NewString(keys[i])
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}

	case "axshutdownaxonaspserver":
		if !loadAxFunctionConfig().enableServerShutdown {
			return NewBool(false)
		}
		os.Exit(int(ShutdownFunctionFromASP))
		return NewBool(true)

	case "axchangedir":
		if len(args) == 0 {
			return NewBool(false)
		}
		if err := os.Chdir(args[0].String()); err != nil {
			return NewBool(false)
		}
		return NewBool(true)

	case "axcurrentdir":
		wd, err := os.Getwd()
		if err != nil {
			return NewString("")
		}
		return NewString(wd)

	case "axhostnamevalue":
		host, err := os.Hostname()
		if err != nil {
			return NewString("")
		}
		return NewString(host)

	case "axclearenvironment":
		os.Clearenv()
		return NewBool(true)

	case "axenvironmentlist":
		envList := filterEnvironmentEntries(os.Environ())
		values := make([]Value, len(envList))
		for i, v := range envList {
			values[i] = NewString(v)
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, values)}

	case "axenvironmentvalue":
		if len(args) == 0 {
			return NewString("")
		}
		val, found := os.LookupEnv(args[0].String())
		if found {
			return NewString(val)
		}
		if len(args) > 1 {
			return args[1]
		}
		return NewString("")

	case "axprocessid":
		return NewInteger(int64(os.Getpid()))

	case "axeffectiveuserid":
		if runtime.GOOS == "windows" {
			return NewInteger(-1)
		}
		return NewInteger(int64(os.Geteuid()))

	case "axdirseparator":
		return NewString(string(os.PathSeparator))

	case "axpathlistseparator":
		return NewString(string(os.PathListSeparator))

	case "axintegersizebytes":
		return NewInteger(int64(strconv.IntSize / 8))

	case "axplatformbits":
		return NewInteger(int64(strconv.IntSize))

	case "axexecutablepath":
		execPath, err := os.Executable()
		if err != nil {
			return NewString("")
		}
		return NewString(execPath)

	case "axexecute":
		if len(args) == 0 {
			return NewBool(false)
		}
		cmdStr := args[0].String()
		if cmdStr == "" {
			return NewBool(false)
		}
		var cmd *exec.Cmd
		if runtime.GOOS == "windows" {
			cmd = exec.Command("cmd.exe", "/c", cmdStr)
		} else {
			cmd = exec.Command("sh", "-c", cmdStr)
		}
		var stdout bytes.Buffer
		cmd.Stdout = &stdout
		cmd.Stderr = &stdout
		_ = cmd.Run()
		output := strings.TrimRight(stdout.String(), "\r\n")
		return NewString(output)

	case "axsysteminfo":
		mode := "a"
		if len(args) > 0 {
			candidate := strings.ToLower(strings.TrimSpace(args[0].String()))
			if candidate != "" {
				mode = candidate
			}
		}
		hostName, _ := os.Hostname()
		switch mode {
		case "s":
			return NewString(runtime.GOOS)
		case "n":
			return NewString(hostName)
		case "v":
			return NewString(runtime.Version())
		case "m":
			return NewString(runtime.GOARCH)
		default:
			return NewString(fmt.Sprintf("%s %s %s %s", runtime.GOOS, hostName, runtime.Version(), runtime.GOARCH))
		}

	case "axcurrentuser":
		current, err := user.Current()
		if err == nil && current != nil && current.Username != "" {
			return NewString(current.Username)
		}
		if runtime.GOOS == "windows" {
			if val := os.Getenv("USERNAME"); val != "" {
				return NewString(val)
			}
		}
		if val := os.Getenv("USER"); val != "" {
			return NewString(val)
		}
		return NewString("")

	// --- Document / Response Functions ---

	case "axw", "document.write", "documentwrite":
		if len(args) == 0 {
			return NewEmpty()
		}
		if al.vm.host != nil && al.vm.host.Response() != nil {
			if funcLower == "axw" {
				al.vm.host.Response().Write(html.EscapeString(args[0].String()))
			} else {
				al.vm.host.Response().Write(args[0].String())
			}
		}
		return NewEmpty()

	// --- Math / Numeric Functions ---

	case "axmax":
		if len(args) == 0 {
			return NewInteger(0)
		}
		maxVal := al.vm.asFloat(args[0])
		for _, arg := range args[1:] {
			v := al.vm.asFloat(arg)
			if v > maxVal {
				maxVal = v
			}
		}
		return NewDouble(maxVal)

	case "axmin":
		if len(args) == 0 {
			return NewInteger(0)
		}
		minVal := al.vm.asFloat(args[0])
		for _, arg := range args[1:] {
			v := al.vm.asFloat(arg)
			if v < minVal {
				minVal = v
			}
		}
		return NewDouble(minVal)

	case "axintegermax":
		maxInt := int(^uint(0) >> 1)
		return NewInteger(int64(maxInt))

	case "axintegermin":
		maxInt := int(^uint(0) >> 1)
		return NewInteger(int64(-maxInt - 1))

	case "axceil":
		if len(args) == 0 {
			return NewInteger(0)
		}
		return NewDouble(math.Ceil(al.vm.asFloat(args[0])))

	case "axfloor":
		if len(args) == 0 {
			return NewInteger(0)
		}
		return NewDouble(math.Floor(al.vm.asFloat(args[0])))

	case "axrand":
		if len(args) == 0 {
			return NewInteger(int64(rand.Int()))
		}
		if len(args) == 1 {
			maxBound := al.vm.asInt(args[0])
			if maxBound <= 0 {
				return NewInteger(0)
			}
			return NewInteger(int64(rand.Intn(int(maxBound) + 1)))
		}
		minBound := al.vm.asInt(args[0])
		maxBound := al.vm.asInt(args[1])
		if minBound > maxBound {
			minBound, maxBound = maxBound, minBound
		}
		diff := maxBound - minBound
		if diff <= 0 {
			return NewInteger(int64(minBound))
		}
		return NewInteger(int64(minBound) + int64(rand.Intn(int(diff)+1)))

	case "axnumberformat":
		if len(args) == 0 {
			return NewString("")
		}
		num := al.vm.asFloat(args[0])
		decimals := 2
		decPoint := "."
		thousandsSep := ","

		if len(args) > 1 {
			decimals = int(al.vm.asInt(args[1]))
		}
		if len(args) > 2 {
			decPoint = args[2].String()
		}
		if len(args) > 3 {
			thousandsSep = args[3].String()
		}

		formatted := fmt.Sprintf("%.*f", decimals, num)
		parts := strings.Split(formatted, ".")

		intPart := parts[0]
		if thousandsSep != "" {
			var result strings.Builder
			for i, ch := range intPart {
				if i > 0 && (len(intPart)-i)%3 == 0 && ch != '-' {
					result.WriteString(thousandsSep)
				}
				result.WriteRune(ch)
			}
			intPart = result.String()
		}

		if decimals > 0 && len(parts) > 1 {
			return NewString(intPart + decPoint + parts[1])
		}
		return NewString(intPart)

	case "axpi":
		return NewDouble(math.Pi)

	case "axsmallestfloatvalue":
		return NewDouble(math.SmallestNonzeroFloat64)

	case "axfloatprecisiondigits":
		return NewInteger(15)

	// --- Array Functions ---

	case "axcount":
		if len(args) == 0 {
			return NewInteger(0)
		}
		v := args[0]
		if v.Type == VTArray && v.Arr != nil {
			return NewInteger(int64(len(v.Arr.Values)))
		}
		return NewInteger(0)

	case "axexplode":
		if len(args) < 2 {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		delimiter := args[0].String()
		str := args[1].String()
		limit := -1
		if len(args) > 2 {
			limit = int(al.vm.asInt(args[2]))
		}

		var parts []string
		if delimiter == "" {
			parts = strings.Split(str, "")
		} else {
			parts = strings.Split(str, delimiter)
		}

		if limit > 0 && len(parts) > limit {
			parts = parts[:limit]
		}

		result := make([]Value, len(parts))
		for i, p := range parts {
			result[i] = NewString(p)
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, result)}

	case "aximplode":
		if len(args) < 2 {
			return NewString("")
		}
		glue := args[0].String()
		v := args[1]
		if v.Type == VTArray && v.Arr != nil {
			var strs []string
			for _, val := range v.Arr.Values {
				strs = append(strs, al.vm.valueToString(val))
			}
			return NewString(strings.Join(strs, glue))
		}
		return NewString("")

	case "axarrayreverse":
		if len(args) == 0 {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		v := args[0]
		if v.Type == VTArray && v.Arr != nil {
			length := len(v.Arr.Values)
			result := make([]Value, length)
			for i, val := range v.Arr.Values {
				result[length-1-i] = val
			}
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, result)}
		}
		return v

	case "axrange":
		if len(args) < 2 {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		start := al.vm.asInt(args[0])
		end := al.vm.asInt(args[1])
		step := int(1)
		if len(args) > 2 {
			step = al.vm.asInt(args[2])
		}
		if step == 0 {
			step = 1
		}
		var result []Value
		if step > 0 {
			for i := start; i <= end; i += step {
				result = append(result, NewInteger(int64(i)))
			}
		} else {
			for i := start; i >= end; i += step {
				result = append(result, NewInteger(int64(i)))
			}
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, result)}

	// --- String Functions ---

	case "axstringreplace":
		if len(args) < 3 {
			return NewString("")
		}
		search := args[0].String()
		replace := args[1].String()
		subject := args[2].String()
		return NewString(strings.ReplaceAll(subject, search, replace))

	case "axpad":
		if len(args) < 2 {
			return NewString("")
		}
		str := args[0].String()
		length := int(al.vm.asInt(args[1]))
		padString := " "
		padType := 1

		if len(args) > 2 {
			padString = args[2].String()
		}
		if len(args) > 3 {
			padType = int(al.vm.asInt(args[3]))
		}

		if len(str) >= length {
			return NewString(str)
		}

		padLen := length - len(str)
		padding := ""
		for len(padding) < padLen {
			padding += padString
		}
		padding = padding[:padLen]

		switch padType {
		case 0: // LEFT
			return NewString(padding + str)
		case 2: // BOTH
			leftPad := padLen / 2
			rightPad := padLen - leftPad
			leftPadding := ""
			rightPadding := ""
			for len(leftPadding) < leftPad {
				leftPadding += padString
			}
			leftPadding = leftPadding[:leftPad]
			for len(rightPadding) < rightPad {
				rightPadding += padString
			}
			rightPadding = rightPadding[:rightPad]
			return NewString(leftPadding + str + rightPadding)
		default: // RIGHT
			return NewString(str + padding)
		}

	case "axrepeat":
		if len(args) < 2 {
			return NewString("")
		}
		str := args[0].String()
		times := int(al.vm.asInt(args[1]))
		if times < 0 {
			times = 0
		}
		return NewString(strings.Repeat(str, times))

	case "axucfirst":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		if len(str) == 0 {
			return NewString(str)
		}
		runes := []rune(str)
		runes[0] = []rune(strings.ToUpper(string(runes[0])))[0]
		return NewString(string(runes))

	case "axwordcount":
		if len(args) == 0 {
			return NewInteger(0)
		}
		str := args[0].String()
		format := 0
		if len(args) > 1 {
			format = int(al.vm.asInt(args[1]))
		}

		words := strings.Fields(str)

		if format == 1 {
			result := make([]Value, len(words))
			for i, w := range words {
				result[i] = NewString(w)
			}
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, result)}
		}
		return NewInteger(int64(len(words)))

	case "axnl2br":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		str = strings.ReplaceAll(str, "\r\n", "<br>")
		str = strings.ReplaceAll(str, "\n", "<br>")
		str = strings.ReplaceAll(str, "\r", "<br>")
		return NewString(str)

	case "axtrim":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		chars := " \t\n\r\v\f"
		if len(args) > 1 {
			chars = args[1].String()
		}
		return NewString(strings.Trim(str, chars))

	case "axstringgetcsv":
		if len(args) == 0 {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}
		str := args[0].String()
		delimiter := ","

		if len(args) > 1 {
			delimiter = args[1].String()
		}

		reader := csv.NewReader(strings.NewReader(str))
		if len(delimiter) > 0 {
			reader.Comma = rune(delimiter[0])
		}

		record, err := reader.Read()
		if err != nil {
			return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, []Value{})}
		}

		result := make([]Value, len(record))
		for i, v := range record {
			result[i] = NewString(v)
		}
		return Value{Type: VTArray, Arr: NewVBArrayFromValues(0, result)}

	// --- Hashing / Encoding Functions ---

	case "axmd5":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		return NewString(fmt.Sprintf("%x", md5.Sum([]byte(str))))

	case "axsha1":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		return NewString(fmt.Sprintf("%x", sha1.Sum([]byte(str))))

	case "axhash":
		if len(args) < 2 {
			return NewString("")
		}
		algo := strings.ToLower(args[0].String())
		str := args[1].String()
		switch algo {
		case "sha256":
			return NewString(fmt.Sprintf("%x", sha256.Sum256([]byte(str))))
		case "sha1":
			return NewString(fmt.Sprintf("%x", sha1.Sum([]byte(str))))
		case "md5":
			return NewString(fmt.Sprintf("%x", md5.Sum([]byte(str))))
		}
		return NewString("")

	case "axbase64encode":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		return NewString(base64.StdEncoding.EncodeToString([]byte(str)))

	case "axbase64decode":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		decoded, err := base64.StdEncoding.DecodeString(str)
		if err != nil {
			return NewString("")
		}
		return NewString(string(decoded))

	case "axurldecode":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		decoded, err := url.QueryUnescape(str)
		if err != nil {
			return NewString(str)
		}
		return NewString(decoded)

	case "axrawurldecode":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		str = strings.ReplaceAll(str, "+", " ")
		decoded, err := url.QueryUnescape(str)
		if err != nil {
			return NewString(str)
		}
		return NewString(decoded)

	case "axrgbtohex":
		if len(args) < 3 {
			return NewString("#000000")
		}
		r := al.vm.asInt(args[0]) & 0xFF
		g := al.vm.asInt(args[1]) & 0xFF
		b := al.vm.asInt(args[2]) & 0xFF
		return NewString(fmt.Sprintf("#%02X%02X%02X", r, g, b))

	case "axhextorgb":
		if len(args) == 0 {
			return NewString("rgb(0,0,0)")
		}
		r, g, b, ok := parseHTMLHexColor(args[0].String())
		if !ok {
			return NewString("rgb(0,0,0)")
		}
		return NewString(fmt.Sprintf("rgb(%d,%d,%d)", r, g, b))

	case "axhtmlspecialchars":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		return NewString(html.EscapeString(str))

	case "axstriptags":
		if len(args) == 0 {
			return NewString("")
		}
		str := args[0].String()
		re := regexp.MustCompile(`<[^>]*>`)
		result := re.ReplaceAllString(str, "")
		return NewString(result)

	// --- Validation Functions ---

	case "axfiltervalidateip":
		if len(args) == 0 {
			return NewBool(false)
		}
		ipStr := args[0].String()
		ipObj := net.ParseIP(ipStr)
		return NewBool(ipObj != nil)

	case "axfiltervalidateemail":
		if len(args) == 0 {
			return NewBool(false)
		}
		emailStr := args[0].String()
		_, err := mail.ParseAddress(emailStr)
		return NewBool(err == nil)

	// --- Type Checking Functions ---

	case "axisint":
		if len(args) == 0 {
			return NewBool(false)
		}
		t := args[0].Type
		return NewBool(t == VTInteger)

	case "axisfloat":
		if len(args) == 0 {
			return NewBool(false)
		}
		t := args[0].Type
		return NewBool(t == VTDouble)

	case "axctypealpha":
		if len(args) == 0 {
			return NewBool(false)
		}
		str := args[0].String()
		if str == "" {
			return NewBool(false)
		}
		for _, ch := range str {
			if !((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z')) {
				return NewBool(false)
			}
		}
		return NewBool(true)

	case "axctypealnum":
		if len(args) == 0 {
			return NewBool(false)
		}
		str := args[0].String()
		if str == "" {
			return NewBool(false)
		}
		for _, ch := range str {
			if !((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9')) {
				return NewBool(false)
			}
		}
		return NewBool(true)

	case "axempty":
		if len(args) == 0 {
			return NewBool(true)
		}
		v := args[0]
		if v.Type == VTEmpty || v.Type == VTNull {
			return NewBool(true)
		}
		if v.Type == VTString && v.String() == "" {
			return NewBool(true)
		}
		if v.Type == VTInteger && v.Num == 0 {
			return NewBool(true)
		}
		if v.Type == VTDouble && v.Flt == 0 {
			return NewBool(true)
		}
		if v.Type == VTBool && v.Num == 0 {
			return NewBool(true)
		}
		return NewBool(false)

	case "axisset":
		if len(args) == 0 {
			return NewBool(false)
		}
		v := args[0]
		return NewBool(v.Type != VTEmpty && v.Type != VTNull)

	// --- Date/Time Functions ---

	case "axtime":
		return NewInteger(time.Now().In(builtinCurrentLocation(al.vm)).Unix())

	case "axdate":
		if len(args) == 0 {
			return NewString("")
		}
		format := args[0].String()
		location := builtinCurrentLocation(al.vm)
		localeTag := strings.ToLower(builtinLocaleTag(al.vm))
		timestamp := time.Now().In(location).Unix()
		if len(args) > 1 {
			timestamp = int64(al.vm.asInt(args[1]))
		}
		t := time.Unix(timestamp, 0).In(location)
		return NewString(al.formatDateEx(format, t, localeTag))

	case "axlastmodified":
		if al.vm.host != nil && al.vm.host.Server() != nil {
			location := builtinCurrentLocation(al.vm)
			path := al.vm.host.Server().MapPath("")
			info, err := os.Stat(path)
			if err == nil {
				return NewInteger(info.ModTime().In(location).Unix())
			}
		}
		return NewInteger(0)

	// --- Network / Remote ---

	case "axgetremotefile":
		if len(args) == 0 {
			return NewBool(false)
		}
		reqUrl := args[0].String()
		if reqUrl == "" || (!strings.HasPrefix(reqUrl, "http://") && !strings.HasPrefix(reqUrl, "https://")) {
			return NewBool(false)
		}
		client := &http.Client{Timeout: 10 * time.Second}
		resp, err := client.Get(reqUrl)
		if err != nil || resp.StatusCode != 200 {
			return NewBool(false)
		}
		defer resp.Body.Close()
		body, err := io.ReadAll(resp.Body)
		if err != nil {
			return NewBool(false)
		}
		return NewString(string(body))

	case "axgenerateguid":
		b := make([]byte, 16)
		rand.Read(b)
		b[6] = (b[6] & 0x0f) | 0x40
		b[8] = (b[8] & 0x3f) | 0x80
		return NewString(fmt.Sprintf("%x-%x-%x-%x-%x", b[0:4], b[4:6], b[6:8], b[8:10], b[10:]))

	// --- Configuration Helpers ---

	// axgetdefaultcss returns the value of axfunctions.ax_default_css_path from axonasp.toml.
	// It is intended to be used by ASP pages that need to inline or emit the default
	// stylesheet content without hard-coding file locations in every page.
	case "axgetdefaultcss":
		cssPath := loadAxFunctionConfig().defaultCSSPath
		if strings.TrimSpace(cssPath) == "" {
			return NewString("")
		}
		cssData, err := os.ReadFile(cssPath)
		if err != nil {
			return NewString("")
		}
		return NewString(string(cssData))

	// axgetlogo reads the configured logo file and returns an inline base64 data URI.
	case "axgetlogo":
		logoPath := loadAxFunctionConfig().defaultLogoPath
		if strings.TrimSpace(logoPath) == "" {
			return NewString("")
		}
		logoData, err := os.ReadFile(logoPath)
		if err != nil || len(logoData) == 0 {
			return NewString("")
		}
		ext := strings.ToLower(filepath.Ext(logoPath))
		mimeType := mime.TypeByExtension(ext)
		if mimeType == "" {
			mimeType = http.DetectContentType(logoData)
		}
		if mimeType == "" {
			mimeType = "application/octet-stream"
		}
		return NewString("data:" + mimeType + ";base64," + base64.StdEncoding.EncodeToString(logoData))

	}

	return NewEmpty()
}

// parseHTMLHexColor converts #RRGGBB/#RGB HTML color strings into RGB channels.
func parseHTMLHexColor(input string) (int, int, int, bool) {
	hex := strings.TrimSpace(input)
	hex = strings.TrimPrefix(hex, "#")

	if len(hex) == 3 {
		hex = strings.Repeat(string(hex[0]), 2) + strings.Repeat(string(hex[1]), 2) + strings.Repeat(string(hex[2]), 2)
	}

	if len(hex) != 6 {
		return 0, 0, 0, false
	}

	r64, err := strconv.ParseUint(hex[0:2], 16, 8)
	if err != nil {
		return 0, 0, 0, false
	}
	g64, err := strconv.ParseUint(hex[2:4], 16, 8)
	if err != nil {
		return 0, 0, 0, false
	}
	b64, err := strconv.ParseUint(hex[4:6], 16, 8)
	if err != nil {
		return 0, 0, 0, false
	}

	return int(r64), int(g64), int(b64), true
}

// formatDateEx replicates PHP-like date formatting patterns used heavily in some Ax* routines.
func (al *AxonLibrary) formatDateEx(format string, t time.Time, localeTag string) string {
	monthsLong := localizedMonthNames(localeTag, false)
	monthsShort := localizedMonthNames(localeTag, true)
	weekdaysLong := localizedWeekdayNames(localeTag, false)
	weekdaysShort := localizedWeekdayNames(localeTag, true)

	var result strings.Builder
	result.Grow(len(format) + 16)
	escaped := false

	for i := 0; i < len(format); i++ {
		ch := format[i]
		if escaped {
			result.WriteByte(ch)
			escaped = false
			continue
		}
		if ch == '\\' {
			escaped = true
			continue
		}

		switch ch {
		case 'Y':
			result.WriteString(fmt.Sprintf("%d", t.Year()))
		case 'y':
			result.WriteString(fmt.Sprintf("%02d", t.Year()%100))
		case 'm':
			result.WriteString(fmt.Sprintf("%02d", t.Month()))
		case 'n':
			result.WriteString(fmt.Sprintf("%d", t.Month()))
		case 'd':
			result.WriteString(fmt.Sprintf("%02d", t.Day()))
		case 'j':
			result.WriteString(fmt.Sprintf("%d", t.Day()))
		case 'H':
			result.WriteString(fmt.Sprintf("%02d", t.Hour()))
		case 'G':
			result.WriteString(fmt.Sprintf("%d", t.Hour()))
		case 'h':
			hour := t.Hour() % 12
			if hour == 0 {
				hour = 12
			}
			result.WriteString(fmt.Sprintf("%02d", hour))
		case 'g':
			hour := t.Hour() % 12
			if hour == 0 {
				hour = 12
			}
			result.WriteString(fmt.Sprintf("%d", hour))
		case 'i':
			result.WriteString(fmt.Sprintf("%02d", t.Minute()))
		case 's':
			result.WriteString(fmt.Sprintf("%02d", t.Second()))
		case 'a':
			if t.Hour() < 12 {
				result.WriteString("am")
			} else {
				result.WriteString("pm")
			}
		case 'A':
			if t.Hour() < 12 {
				result.WriteString("AM")
			} else {
				result.WriteString("PM")
			}
		case 'w':
			result.WriteString(fmt.Sprintf("%d", t.Weekday()))
		case 'z':
			result.WriteString(fmt.Sprintf("%d", t.YearDay()-1))
		case 'F':
			monthIdx := int(t.Month()) - 1
			if monthIdx >= 0 && monthIdx < len(monthsLong) {
				result.WriteString(monthsLong[monthIdx])
			} else {
				result.WriteString(t.Month().String())
			}
		case 'M':
			monthIdx := int(t.Month()) - 1
			if monthIdx >= 0 && monthIdx < len(monthsShort) {
				result.WriteString(monthsShort[monthIdx])
			} else {
				result.WriteString(t.Month().String()[:3])
			}
		case 'l':
			weekdayIdx := int(t.Weekday())
			if weekdayIdx >= 0 && weekdayIdx < len(weekdaysLong) {
				result.WriteString(weekdaysLong[weekdayIdx])
			} else {
				result.WriteString(t.Weekday().String())
			}
		case 'D':
			weekdayIdx := int(t.Weekday())
			if weekdayIdx >= 0 && weekdayIdx < len(weekdaysShort) {
				result.WriteString(weekdaysShort[weekdayIdx])
			} else {
				result.WriteString(t.Weekday().String()[:3])
			}
		default:
			result.WriteByte(ch)
		}
	}

	if escaped {
		result.WriteByte('\\')
	}

	return result.String()
}
