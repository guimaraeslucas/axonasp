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
	"errors"
	"log"
	"os"
	"path/filepath"
	"runtime"
	"strconv"
	"strings"
	"sync"
	"sync/atomic"

	"g3pix.com.br/axonasp/axonconfig"
	"g3pix.com.br/axonasp/axonvm/asp"
)

var (
	internalErrorConsoleLogger = log.New(os.Stderr, "", log.LstdFlags)
	internalErrorLogEnabled    atomic.Bool
	internalErrorConfigOnce    sync.Once
	internalErrorConfigured    bool
	internalErrorLogRootPath   atomic.Value
	internalErrorLogFileMu     sync.Mutex
)

// init initializes internal error logging flags from configuration and environment.
func init() {
	internalErrorLogEnabled.Store(resolveErrorLogFileEnabled())
}

// AxonASPError represents a platform-level AxonASP error with optional context.
type AxonASPError struct {
	Code        AxonASPErrorCode
	Description string
	FileName    string
	Line        int
	Cause       error
}

// Error formats the AxonASP error using the catalog message and optional context.
func (e *AxonASPError) Error() string {
	if e == nil {
		return ""
	}

	return FormatAxonASPError(e.Code, e.Cause, e.Description, e.FileName, e.Line)
}

// Unwrap returns the wrapped underlying error when one is available.
func (e *AxonASPError) Unwrap() error {
	if e == nil {
		return nil
	}

	return e.Cause
}

// NewAxonASPError creates a structured AxonASP error without logging it.
func NewAxonASPError(code AxonASPErrorCode, err error, description string, fileName string, line int) *AxonASPError {
	return &AxonASPError{
		Code:        code,
		Description: strings.TrimSpace(description),
		FileName:    strings.TrimSpace(fileName),
		Line:        line,
		Cause:       err,
	}
}

// AsAxonASPError extracts a structured AxonASP error from the provided error value.
func AsAxonASPError(err error) (*AxonASPError, bool) {
	if err == nil {
		return nil, false
	}

	var axErr *AxonASPError
	if errors.As(err, &axErr) {
		return axErr, true
	}

	return nil, false
}

// SetInternalErrorLogEnabled enables or disables synchronous error.log writes.
func SetInternalErrorLogEnabled(enabled bool) {
	internalErrorLogEnabled.Store(enabled)
}

// SetInternalErrorLogRootPath sets the root directory used for temp/error.log output.
func SetInternalErrorLogRootPath(rootPath string) {
	trimmedPath := strings.TrimSpace(rootPath)
	if trimmedPath == "" {
		return
	}

	absPath, err := filepath.Abs(trimmedPath)
	if err != nil {
		return
	}

	internalErrorLogRootPath.Store(absPath)
}

// resolveErrorLogFileEnabled reads global.enable_error_log_file from config/axonasp.toml and environment.
func resolveErrorLogFileEnabled() bool {
	internalErrorConfigOnce.Do(func() {
		v := axonconfig.NewViper()

		// Preserve legacy behavior for this internal logger: environment overrides are always accepted.
		v.SetEnvKeyReplacer(strings.NewReplacer(".", "_"))
		v.AutomaticEnv()

		internalErrorConfigured = v.GetBool("global.enable_error_log_file")
		if envValue := strings.TrimSpace(os.Getenv("GLOBAL_ENABLE_ERROR_LOG_FILE")); envValue != "" {
			internalErrorConfigured = strings.EqualFold(envValue, "true") || envValue == "1"
		}
		if envValue := strings.TrimSpace(os.Getenv("ENABLE_ERROR_LOG_FILE")); envValue != "" {
			internalErrorConfigured = strings.EqualFold(envValue, "true") || envValue == "1"
		}
	})

	return internalErrorConfigured
}

// FormatAxonASPError builds a compact AxonASP error message with code and context.
func FormatAxonASPError(code AxonASPErrorCode, err error, description string, fileName string, line int) string {
	catalogMessage := code.String()
	trimmedDescription := strings.TrimSpace(description)
	trimmedFileName := strings.TrimSpace(fileName)

	if trimmedDescription == "" {
		trimmedDescription = catalogMessage
	}

	if trimmedFileName == "" && line <= 0 {
		callerFile, callerLine := resolveCallerLocation(2)
		trimmedFileName = callerFile
		line = callerLine
	}

	var builder strings.Builder
	builder.Grow(128)
	builder.WriteString("AxonASP Error [")
	builder.WriteString(strconv.Itoa(int(code)))
	builder.WriteString("] ")
	builder.WriteString(catalogMessage)

	if !strings.EqualFold(trimmedDescription, catalogMessage) {
		builder.WriteString(" | ")
		builder.WriteString(trimmedDescription)
	}

	if trimmedFileName != "" {
		builder.WriteString(" | File: ")
		builder.WriteString(trimmedFileName)
	}

	if line > 0 {
		builder.WriteString(" | Line: ")
		builder.WriteString(strconv.Itoa(line))
	}

	if err != nil {
		cause := strings.TrimSpace(err.Error())
		if cause != "" {
			builder.WriteString(" | Cause: ")
			builder.WriteString(cause)
		}
	}

	return builder.String()
}

// ReportInternalError logs an AxonASP internal error and returns the structured error instance.
func ReportInternalError(code AxonASPErrorCode, err error, description string, fileName string, line int) *AxonASPError {
	if strings.TrimSpace(fileName) == "" && line <= 0 {
		callerFile, callerLine := resolveCallerLocation(1)
		fileName = callerFile
		line = callerLine
	}

	axErr := NewAxonASPError(code, err, description, fileName, line)
	return LogInternalError(axErr)
}

// LogInternalError logs a structured AxonASP error to the console and error.log.
func LogInternalError(err error) *AxonASPError {
	axErr, ok := AsAxonASPError(err)
	if !ok {
		axErr = NewAxonASPError(ErrInternalError, err, "", "", 0)
	}

	message := axErr.Error()
	internalErrorConsoleLogger.Print(message)
	writeInternalErrorLog(message)

	return axErr
}

// LogASPProcessedError logs an ASP/VBScript processed error to console and optional error.log.
func LogASPProcessedError(err *asp.ASPError, context string) {
	if err == nil {
		return
	}

	normalized := err.Clone()
	trimmedContext := strings.TrimSpace(context)

	var builder strings.Builder
	builder.Grow(256)
	builder.WriteString("ASPError")
	if trimmedContext != "" {
		builder.WriteString(" | Context: ")
		builder.WriteString(trimmedContext)
	}
	builder.WriteString(" | Source: ")
	builder.WriteString(strings.TrimSpace(normalized.Source))
	builder.WriteString(" | Number: ")
	builder.WriteString(strconv.Itoa(normalized.Number))
	builder.WriteString(" | Description: ")
	builder.WriteString(strings.TrimSpace(normalized.Description))

	if fileName := strings.TrimSpace(normalized.File); fileName != "" {
		builder.WriteString(" | File: ")
		builder.WriteString(fileName)
	}
	if normalized.Line > 0 {
		builder.WriteString(" | Line: ")
		builder.WriteString(strconv.Itoa(normalized.Line))
	}
	if normalized.Column > 0 {
		builder.WriteString(" | Column: ")
		builder.WriteString(strconv.Itoa(normalized.Column))
	}

	message := builder.String()
	internalErrorConsoleLogger.Print(message)
	writeInternalErrorLog(message)
}

// writeInternalErrorLog appends a formatted message directly to temp/error.log.
// It uses a mutex to prevent concurrent write corruption and creates the file and
// parent directory on first use. Failures are silently ignored to avoid recursion.
func writeInternalErrorLog(message string) {
	if !internalErrorLogEnabled.Load() {
		return
	}

	logFilePath := filepath.Join(currentInternalErrorLogRootPath(), "temp", "error.log")

	internalErrorLogFileMu.Lock()
	defer internalErrorLogFileMu.Unlock()

	if err := os.MkdirAll(filepath.Dir(logFilePath), 0o755); err != nil {
		return
	}

	file, err := os.OpenFile(logFilePath, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0o644)
	if err != nil {
		return
	}
	defer file.Close()

	logger := log.New(file, "", log.LstdFlags)
	logger.Print(message)
}

// resolveRuntimeRootPath resolves the runtime root used for config and temp/error.log placement.
func resolveRuntimeRootPath() string {
	workingDir, err := os.Getwd()
	if err == nil {
		if _, statErr := os.Stat(filepath.Join(workingDir, "config", "axonasp.toml")); statErr == nil {
			return workingDir
		}
	}

	executablePath, err := os.Executable()
	if err == nil {
		executableDir := filepath.Dir(executablePath)
		if _, statErr := os.Stat(filepath.Join(executableDir, "config", "axonasp.toml")); statErr == nil {
			return executableDir
		}
		return executableDir
	}

	if workingDir != "" {
		return workingDir
	}

	return "."
}

// currentInternalErrorLogRootPath returns the configured error-log root path or resolves it from runtime context.
func currentInternalErrorLogRootPath() string {
	if value := internalErrorLogRootPath.Load(); value != nil {
		if path, ok := value.(string); ok {
			trimmedPath := strings.TrimSpace(path)
			if trimmedPath != "" {
				return trimmedPath
			}
		}
	}

	return resolveRuntimeRootPath()
}

// resolveCallerLocation returns a best-effort file name and line number for error reporting.
func resolveCallerLocation(skip int) (string, int) {
	_, fileName, line, ok := runtime.Caller(skip + 1)
	if !ok {
		return "", 0
	}

	workingDir, err := os.Getwd()
	if err == nil {
		if relativePath, relErr := filepath.Rel(workingDir, fileName); relErr == nil {
			fileName = relativePath
		}
	}

	return filepath.ToSlash(filepath.Clean(fileName)), line
}
