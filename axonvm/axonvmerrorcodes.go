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

// The errors in this AxonASP file must be used uniquely and exclusively for
// errors in our GoLang platform implementation, during the execution of
// services, servers, and CLI.
// VBScript/ASP interpretation errors must obligatorily use and return the
// errors indicated in /vbscript/vberrorcodes.go and maintain the Microsoft
// standard for similarity.
type AxonASPErrorCode int

const (
	HTTPBadRequest          AxonASPErrorCode = 400
	HTTPUnauthorized        AxonASPErrorCode = 401
	HTTPForbidden           AxonASPErrorCode = 403
	HTTPNotFound            AxonASPErrorCode = 404
	HTTPMethodNotAllowed    AxonASPErrorCode = 405
	HTTPPayloadTooLarge     AxonASPErrorCode = 413
	HTTPURITooLong          AxonASPErrorCode = 414
	HTTPInternalServerError AxonASPErrorCode = 500
	HTTPNotImplemented      AxonASPErrorCode = 501
	HTTPBadGateway          AxonASPErrorCode = 502
	HTTPServiceUnavailable  AxonASPErrorCode = 503
	HTTPGatewayTimeout      AxonASPErrorCode = 504

	ErrInvalidConfig             AxonASPErrorCode = 1000
	ErrInvalidEnv                AxonASPErrorCode = 1001
	ErrRootDirNotSet             AxonASPErrorCode = 1002
	ErrRootDirInvalid            AxonASPErrorCode = 1003
	ErrRootDirectoryDoesNotExist AxonASPErrorCode = 1004
	ErrPortInvalid               AxonASPErrorCode = 1005
	ErrServerLocationInvalid     AxonASPErrorCode = 1006
	ErrCouldNotListenOn          AxonASPErrorCode = 1007
	ErrCOMProviderModeInvalid    AxonASPErrorCode = 1008
	ErrDefaultPagesInvalid       AxonASPErrorCode = 1009
	ErrDebugInvalid              AxonASPErrorCode = 1010
	ErrInvalidLocale             AxonASPErrorCode = 1011
	ErrInvalidTimezone           AxonASPErrorCode = 1012
	ErrViperReadConfigFailed     AxonASPErrorCode = 1013

	ErrMissingFilePath             AxonASPErrorCode = 2000
	ErrFileNotFound                AxonASPErrorCode = 2001
	ErrCouldNotReadFile            AxonASPErrorCode = 2002
	ErrCouldNotResolveCurrentDir   AxonASPErrorCode = 2003
	ErrPathIsADirectory            AxonASPErrorCode = 2004
	ErrBadFileName                 AxonASPErrorCode = 2005
	ErrWrongFileSize               AxonASPErrorCode = 2006
	ErrWrongFileType               AxonASPErrorCode = 2007
	ErrFileTypeNotAllowed          AxonASPErrorCode = 2008
	ErrExtensionNotAllowed         AxonASPErrorCode = 2009
	ErrExtensionNotEnabledInGlobal AxonASPErrorCode = 2010
	ErrFailedToReadASPFile         AxonASPErrorCode = 2011

	ErrRuntimeError           AxonASPErrorCode = 3000
	ErrPanic                  AxonASPErrorCode = 3001
	ErrInternalGolangPanic    AxonASPErrorCode = 3002
	ErrInternalError          AxonASPErrorCode = 3003
	ErrOutOfMemory            AxonASPErrorCode = 3004
	ErrMemoryLimitExceeded    AxonASPErrorCode = 3005
	ErrOverflow               AxonASPErrorCode = 3006
	ErrUnderflow              AxonASPErrorCode = 3007
	ErrTimeExpired            AxonASPErrorCode = 3008
	ErrTimeExecutionError     AxonASPErrorCode = 3009
	ErrExpired                AxonASPErrorCode = 3010
	ErrServerForcedToShutdown AxonASPErrorCode = 3011

	ErrCompileError                         AxonASPErrorCode = 4000
	ErrScriptTimeout                        AxonASPErrorCode = 4001
	ErrFunctionNotImplemented               AxonASPErrorCode = 4002
	ErrNotImplemented                       AxonASPErrorCode = 4003
	ErrErrorOnLibrary                       AxonASPErrorCode = 4004
	ErrErrorOnCustomFunction                AxonASPErrorCode = 4005
	ErrAxonVMError                          AxonASPErrorCode = 4006
	ErrInvalidProcedureCallOrArg            AxonASPErrorCode = 4007
	ErrInvalidArrayBondType                 AxonASPErrorCode = 4008
	ErrInteractiveFunctionNotSupportedInASP AxonASPErrorCode = 4009
	ErrResponseBufferLimitExceeded          AxonASPErrorCode = 4010
	ErrScriptTimeoutDetachedGoroutine       AxonASPErrorCode = 4011
	ErrLibraryDisabled                      AxonASPErrorCode = 4012

	ErrInvalidCacheVersion          AxonASPErrorCode = 5000
	ErrInvalidCacheFile             AxonASPErrorCode = 5001
	ErrCacheCleanupInvalid          AxonASPErrorCode = 5002
	ErrIncludeCacheMaxMemoryInvalid AxonASPErrorCode = 5003

	ErrFastCGIPipeClosed       AxonASPErrorCode = 6000
	ErrFastCGIProtocolError    AxonASPErrorCode = 6001
	ErrCLIArgumentMissing      AxonASPErrorCode = 6002
	ErrInvalidName             AxonASPErrorCode = 6003
	ErrThisIsATest             AxonASPErrorCode = 6004
	ErrCLIRunCommandNotEnabled AxonASPErrorCode = 6005
	ErrCLINotEnabled           AxonASPErrorCode = 6006
	ErrCLIMissingFilePath      AxonASPErrorCode = 6007
	ErrPageCounterDisabled     AxonASPErrorCode = 6008

	ErrServiceCreateFailed                AxonASPErrorCode = 6100
	ErrServiceLoggerCreateFailed          AxonASPErrorCode = 6101
	ErrServiceRunFailed                   AxonASPErrorCode = 6102
	ErrServiceControlCommandFailed        AxonASPErrorCode = 6103
	ErrServiceResolveExecutablePathFailed AxonASPErrorCode = 6104
	ErrServiceExecutableNotFound          AxonASPErrorCode = 6105
	ErrServiceStartProcessFailed          AxonASPErrorCode = 6106
	ErrServiceStopProcessFailed           AxonASPErrorCode = 6107
	ErrServiceChildExitedUnexpectedly     AxonASPErrorCode = 6108
	ErrServiceInvalidEnvironmentVariable  AxonASPErrorCode = 6109

	ErrG3FCInvalidHeader       AxonASPErrorCode = 7000
	ErrG3FCDecryptionFailed    AxonASPErrorCode = 7001
	ErrG3FCDecompressionFailed AxonASPErrorCode = 7002
	ErrG3FCMismatchedChecksum  AxonASPErrorCode = 7003
	ErrG3FCPasswordRequired    AxonASPErrorCode = 7004
	ErrG3FCFileNotFound        AxonASPErrorCode = 7005
	ShutdownFunctionFromASP    AxonASPErrorCode = 7006

	// Request BinaryRead / Form mutual-exclusion errors (IIS ASP 0206 / 0207).
	ErrRequestFormAfterBinaryRead AxonASPErrorCode = 8000
	ErrRequestBinaryReadAfterForm AxonASPErrorCode = 8001
	// ADODB.Stream state/property constraint errors.
	ErrADODBStreamObjectClosed      AxonASPErrorCode = 8010
	ErrADODBStreamTypeConstraint    AxonASPErrorCode = 8011
	ErrADODBStreamCharsetConstraint AxonASPErrorCode = 8012
	ErrADODBStreamInvalidArgument   AxonASPErrorCode = 8013
	ErrADODBStreamIOError           AxonASPErrorCode = 8014

	// G3DB native database library errors (9000–9099).
	ErrG3DBConnectionAlreadyOpen AxonASPErrorCode = 9000
	ErrG3DBConnectionNotOpen     AxonASPErrorCode = 9001
	ErrG3DBRequiresDriverAndDSN  AxonASPErrorCode = 9002
	ErrG3DBQueryRequiresSQL      AxonASPErrorCode = 9003
	ErrG3DBExecRequiresSQL       AxonASPErrorCode = 9004
	ErrG3DBPrepareRequiresSQL    AxonASPErrorCode = 9005
	ErrG3DBUnsupportedDriver     AxonASPErrorCode = 9006
	ErrG3DBPingFailed            AxonASPErrorCode = 9007
	ErrG3DBQueryFailed           AxonASPErrorCode = 9008
	ErrG3DBExecFailed            AxonASPErrorCode = 9009
	ErrG3DBPrepareFailed         AxonASPErrorCode = 9010
	ErrG3DBTransactionFailed     AxonASPErrorCode = 9011
	ErrG3DBScanFailed            AxonASPErrorCode = 9012
	ErrG3DBResultSetClosed       AxonASPErrorCode = 9013
	ErrG3DBStatementClosed       AxonASPErrorCode = 9014
	ErrG3DBTransactionClosed     AxonASPErrorCode = 9015
	ErrG3DBMissingConfigKeys     AxonASPErrorCode = 9016

	// G3SEARCH native search library errors (9020-9024).
	ErrG3SearchDocsPathMissing  AxonASPErrorCode = 9020
	ErrG3SearchIndexPathMissing AxonASPErrorCode = 9021
	ErrG3SearchIndexOpenFailed  AxonASPErrorCode = 9022
	ErrG3SearchIndexWriteFailed AxonASPErrorCode = 9023
	ErrG3SearchSearchFailed     AxonASPErrorCode = 9024

	// G3FILEUPLOADER native file upload library errors (9100-9199).
	ErrG3FUNoHTTPRequest        AxonASPErrorCode = 9100
	ErrG3FUFormParseFailed      AxonASPErrorCode = 9101
	ErrG3FUNoMultipartData      AxonASPErrorCode = 9102
	ErrG3FUFileFieldNotFound    AxonASPErrorCode = 9103
	ErrG3FUExtensionNotAllowed  AxonASPErrorCode = 9104
	ErrG3FUFileSizeExceedsMax   AxonASPErrorCode = 9105
	ErrG3FUInvalidTargetDir     AxonASPErrorCode = 9106
	ErrG3FUTempFileCreateFailed AxonASPErrorCode = 9107
	ErrG3FUTempFileWriteFailed  AxonASPErrorCode = 9108
	ErrG3FUTempFileSyncFailed   AxonASPErrorCode = 9109
	ErrG3FUFinalMoveFailed      AxonASPErrorCode = 9110
	ErrG3FUOpenFileFailed       AxonASPErrorCode = 9111

	// G3AXONLIVE reactive component framework errors (10000-10099).
	ErrG3ALNotInitialized         AxonASPErrorCode = 10000
	ErrG3ALInvalidSessionID       AxonASPErrorCode = 10001
	ErrG3ALInvalidComponentID     AxonASPErrorCode = 10002
	ErrG3ALInvalidEventName       AxonASPErrorCode = 10003
	ErrG3ALMissingXHeader         AxonASPErrorCode = 10004
	ErrG3ALBodyReadFailed         AxonASPErrorCode = 10005
	ErrG3ALInvalidJSONPayload     AxonASPErrorCode = 10006
	ErrG3ALSessionNotRegistered   AxonASPErrorCode = 10007
	ErrG3ALPagePathOutsideRoot    AxonASPErrorCode = 10008
	ErrG3ALComponentLimitExceeded AxonASPErrorCode = 10009
	ErrG3ALTimerDelayInvalid      AxonASPErrorCode = 10010
	ErrG3ALWebSocketNotSupported  AxonASPErrorCode = 10011
	ErrG3ALResponseAlreadyEnded   AxonASPErrorCode = 10012
	ErrG3ALDisabled               AxonASPErrorCode = 10013
)

var AxonASPErrorMessages = map[AxonASPErrorCode]string{
	// HTTP Standard
	HTTPBadRequest:          "Bad Request",
	HTTPUnauthorized:        "Unauthorized",
	HTTPForbidden:           "Forbidden",
	HTTPNotFound:            "Not Found",
	HTTPMethodNotAllowed:    "Method Not Allowed",
	HTTPPayloadTooLarge:     "Payload Too Large",
	HTTPURITooLong:          "URI Too Long",
	HTTPInternalServerError: "Internal Server Error",
	HTTPNotImplemented:      "Not Implemented",
	HTTPBadGateway:          "Bad Gateway",
	HTTPServiceUnavailable:  "Service Unavailable",
	HTTPGatewayTimeout:      "Gateway Timeout",

	// Config / Setup
	ErrInvalidConfig:             "Invalid configuration",
	ErrInvalidEnv:                "Invalid .env file or configuration",
	ErrRootDirNotSet:             "Root directory not set",
	ErrRootDirInvalid:            "Root directory invalid",
	ErrRootDirectoryDoesNotExist: "Warning: Root directory does not exist",
	ErrPortInvalid:               "Port invalid",
	ErrServerLocationInvalid:     "Server location invalid",
	ErrCouldNotListenOn:          "Could not listen on specified port/address",
	ErrCOMProviderModeInvalid:    "COM provider mode invalid",
	ErrDefaultPagesInvalid:       "Default pages invalid",
	ErrDebugInvalid:              "Debug invalid",
	ErrInvalidLocale:             "Invalid locale",
	ErrInvalidTimezone:           "Invalid timezone",
	ErrViperReadConfigFailed:     "Viper: Failed to read configuration file, using defaults",

	// File System
	ErrMissingFilePath:             "Missing file path",
	ErrFileNotFound:                "File not found",
	ErrCouldNotReadFile:            "Could not read file",
	ErrCouldNotResolveCurrentDir:   "Could not resolve current directory",
	ErrPathIsADirectory:            "Path is a directory",
	ErrBadFileName:                 "Bad file name",
	ErrWrongFileSize:               "Wrong file size",
	ErrWrongFileType:               "Wrong file type",
	ErrFileTypeNotAllowed:          "File type not allowed",
	ErrExtensionNotAllowed:         "Extension not allowed",
	ErrExtensionNotEnabledInGlobal: "The selected file extension is not enabled in global.execute_as_asp.",
	ErrFailedToReadASPFile:         "Failed to read the requested ASP file.",

	// Runtime / Execution
	ErrRuntimeError:           "Runtime error",
	ErrPanic:                  "Panic",
	ErrInternalGolangPanic:    "Internal golang panic",
	ErrInternalError:          "Internal error",
	ErrOutOfMemory:            "Out of memory",
	ErrMemoryLimitExceeded:    "Memory limit exceeded",
	ErrOverflow:               "Overflow",
	ErrUnderflow:              "Underflow",
	ErrTimeExpired:            "Time expired",
	ErrTimeExecutionError:     "Time execution error",
	ErrExpired:                "Expired",
	ErrServerForcedToShutdown: "Server forced to shutdown",

	// Script / AxonVM
	ErrCompileError:                         "Compile Error",
	ErrScriptTimeout:                        "Script timeout",
	ErrFunctionNotImplemented:               "Function not implemented",
	ErrNotImplemented:                       "Not implemented",
	ErrErrorOnLibrary:                       "Error on library",
	ErrErrorOnCustomFunction:                "Error on custom function",
	ErrAxonVMError:                          "AxonVM error",
	ErrInvalidProcedureCallOrArg:            "Invalid procedure call or argument",
	ErrInvalidArrayBondType:                 "Invalis Array Bond/Type",
	ErrInteractiveFunctionNotSupportedInASP: "Interactive desktop functions are not supported in ASP server-side execution",
	ErrResponseBufferLimitExceeded:          "Response buffer limit exceeded",
	ErrScriptTimeoutDetachedGoroutine:       "Script timeout reached and execution goroutine was detached",
	ErrLibraryDisabled:                      "The requested library was not compiled into this AxonASP executable. You must compile the server without the `lib_%s_disabled` build tag to enable it.",

	// Cache
	ErrInvalidCacheVersion:          "Invalid cache version",
	ErrInvalidCacheFile:             "Invalid cache file",
	ErrCacheCleanupInvalid:          "Cache cleanup invalid",
	ErrIncludeCacheMaxMemoryInvalid: "Include cache max memory invalid",

	// FastCGI / CLI / Service / Misc
	ErrFastCGIPipeClosed:                  "FastCGI pipe closed unexpectedly",
	ErrFastCGIProtocolError:               "FastCGI protocol error",
	ErrCLIArgumentMissing:                 "Required CLI argument missing",
	ErrInvalidName:                        "Invalid name",
	ErrThisIsATest:                        "This is a test",
	ErrCLIRunCommandNotEnabled:            "CLI run command not enabled in configuration",
	ErrCLINotEnabled:                      "CLI not enabled in configuration",
	ErrCLIMissingFilePath:                 "CLI: Missing file path for -r option",
	ErrPageCounterDisabled:                "MSWC.PageCounter is disabled. Please enable it in the server configuration at config/axonasp.toml",
	ErrServiceCreateFailed:                "Service wrapper failed to create service instance",
	ErrServiceLoggerCreateFailed:          "Service wrapper failed to create service logger",
	ErrServiceRunFailed:                   "Service wrapper failed while running service loop",
	ErrServiceControlCommandFailed:        "Service wrapper failed to execute control command",
	ErrServiceResolveExecutablePathFailed: "Service wrapper failed to resolve configured executable path",
	ErrServiceExecutableNotFound:          "Service wrapper executable target was not found",
	ErrServiceStartProcessFailed:          "Service wrapper failed to start configured executable",
	ErrServiceStopProcessFailed:           "Service wrapper failed to stop child process",
	ErrServiceChildExitedUnexpectedly:     "Service wrapper detected unexpected child process termination",
	ErrServiceInvalidEnvironmentVariable:  "Service wrapper found an invalid environment variable entry",

	ErrG3FCInvalidHeader:       "Invalid G3FC archive header or magic number",
	ErrG3FCDecryptionFailed:    "G3FC decryption failed: incorrect password or corrupted data",
	ErrG3FCDecompressionFailed: "G3FC decompression failed",
	ErrG3FCMismatchedChecksum:  "G3FC checksum mismatch: file may be corrupted",
	ErrG3FCPasswordRequired:    "G3FC password required for this encrypted archive",
	ErrG3FCFileNotFound:        "Requested file not found in G3FC archive",
	ShutdownFunctionFromASP:    "Shutdown function called from ASP script",

	// Request BinaryRead / Form mutual-exclusion.
	ErrRequestFormAfterBinaryRead: "Cannot use Request.Form after calling Request.BinaryRead",
	ErrRequestBinaryReadAfterForm: "Cannot call Request.BinaryRead after using Request.Form",
	// ADODB.Stream state/property constraint errors.
	ErrADODBStreamObjectClosed:      "Operation is not allowed when the object is closed",
	ErrADODBStreamTypeConstraint:    "The stream Type property cannot be changed when Position is not zero",
	ErrADODBStreamCharsetConstraint: "The stream Charset property cannot be set when Position is not zero",
	ErrADODBStreamInvalidArgument:   "Arguments are of the wrong type, are out of range, or are in conflict with one another",
	ErrADODBStreamIOError:           "ADODB.Stream I/O error: file operation failed",

	// G3DB native database library
	ErrG3DBConnectionAlreadyOpen: "G3DB: connection is already open; call Close first",
	ErrG3DBConnectionNotOpen:     "G3DB: connection is not open; call Open or OpenFromEnv first",
	ErrG3DBRequiresDriverAndDSN:  "G3DB.Open: requires two arguments: driver and DSN",
	ErrG3DBQueryRequiresSQL:      "G3DB.Query: requires an SQL string argument",
	ErrG3DBExecRequiresSQL:       "G3DB.Exec: requires an SQL string argument",
	ErrG3DBPrepareRequiresSQL:    "G3DB.Prepare: requires an SQL string argument",
	ErrG3DBUnsupportedDriver:     "G3DB: unsupported database driver",
	ErrG3DBPingFailed:            "G3DB: connection test (Ping) failed",
	ErrG3DBQueryFailed:           "G3DB: query execution failed",
	ErrG3DBExecFailed:            "G3DB: exec statement failed",
	ErrG3DBPrepareFailed:         "G3DB: statement preparation failed",
	ErrG3DBTransactionFailed:     "G3DB: transaction operation failed",
	ErrG3DBScanFailed:            "G3DB: row scan failed",
	ErrG3DBResultSetClosed:       "G3DB: result set is already closed",
	ErrG3DBStatementClosed:       "G3DB: prepared statement is already closed",
	ErrG3DBTransactionClosed:     "G3DB: transaction is already closed",
	ErrG3DBMissingConfigKeys:     "G3DB.OpenFromEnv: missing or incomplete configuration keys in axonasp.toml",

	// G3SEARCH native search library
	ErrG3SearchDocsPathMissing:  "G3SEARCH.BuildIndex: DocsPath is required",
	ErrG3SearchIndexPathMissing: "G3SEARCH: IndexPath is required",
	ErrG3SearchIndexOpenFailed:  "G3SEARCH: failed to open index",
	ErrG3SearchIndexWriteFailed: "G3SEARCH: failed to write index",
	ErrG3SearchSearchFailed:     "G3SEARCH: search execution failed",

	// G3FILEUPLOADER native file upload library
	ErrG3FUNoHTTPRequest:        "G3FILEUPLOADER: No HTTP request context found",
	ErrG3FUFormParseFailed:      "G3FILEUPLOADER: Failed to parse multipart form data",
	ErrG3FUNoMultipartData:      "G3FILEUPLOADER: No multipart form data received",
	ErrG3FUFileFieldNotFound:    "G3FILEUPLOADER: File field '%s' not found in request",
	ErrG3FUExtensionNotAllowed:  "G3FILEUPLOADER: File extension '%s' is not allowed",
	ErrG3FUFileSizeExceedsMax:   "G3FILEUPLOADER: File size exceeds the maximum allowed limit",
	ErrG3FUInvalidTargetDir:     "G3FILEUPLOADER: The specified target directory is invalid or inaccessible",
	ErrG3FUTempFileCreateFailed: "G3FILEUPLOADER: Failed to create temporary file for upload",
	ErrG3FUTempFileWriteFailed:  "G3FILEUPLOADER: Failed to write data to temporary file",
	ErrG3FUTempFileSyncFailed:   "G3FILEUPLOADER: Failed to sync temporary file to disk",
	ErrG3FUFinalMoveFailed:      "G3FILEUPLOADER: Failed to move uploaded file to its final destination",
	ErrG3FUOpenFileFailed:       "G3FILEUPLOADER: Failed to open uploaded file for processing",

	// G3AXONLIVE reactive component framework
	ErrG3ALNotInitialized:         "G3AXONLIVE: InitPage must be called before using AxonLive methods",
	ErrG3ALInvalidSessionID:       "G3AXONLIVE: sessionId is required and cannot be empty",
	ErrG3ALInvalidComponentID:     "G3AXONLIVE: componentId is required and cannot be empty",
	ErrG3ALInvalidEventName:       "G3AXONLIVE: eventName is required and cannot be empty",
	ErrG3ALMissingXHeader:         "G3AXONLIVE: missing or invalid X-G3AxonLive request header",
	ErrG3ALBodyReadFailed:         "G3AXONLIVE: failed to read request body",
	ErrG3ALInvalidJSONPayload:     "G3AXONLIVE: invalid or malformed JSON payload",
	ErrG3ALSessionNotRegistered:   "G3AXONLIVE: session is not registered or has expired; reload the page",
	ErrG3ALPagePathOutsideRoot:    "G3AXONLIVE: registered page path is outside the configured web root directory",
	ErrG3ALComponentLimitExceeded: "G3AXONLIVE: component patch limit per response exceeded",
	ErrG3ALTimerDelayInvalid:      "G3AXONLIVE: SetTimer delay must be a positive integer (milliseconds)",
	ErrG3ALWebSocketNotSupported:  "G3AXONLIVE: WebSocket connections are not supported in FastCGI mode; use the HTTP server",
	ErrG3ALResponseAlreadyEnded:   "G3AXONLIVE: EndAsyncResponse has already been called for this request",
	ErrG3ALDisabled:               "G3AXONLIVE: the G3AxonLive library is disabled; enable it in axonasp.toml under [g3axonlive]",
}

func (e AxonASPErrorCode) String() string {
	if msg, ok := AxonASPErrorMessages[e]; ok {
		return msg
	}
	return "Unknown AxonASP Error"
}
