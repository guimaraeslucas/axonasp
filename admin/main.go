//go:build !wasm

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
//go:generate goversioninfo -icon=icon_admin.ico -64=true
package main

import (
	"context"
	"embed"
	"encoding/json"
	"fmt"
	"html/template"
	"io"
	"io/fs"
	"log"
	"mime"
	"net"
	"net/http"
	"os"
	"os/exec"
	"path"
	"path/filepath"
	"reflect"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"time"

	"g3pix.com.br/axonasp/axonconfig"

	"github.com/shirou/gopsutil/v3/cpu"
	"github.com/shirou/gopsutil/v3/disk"
	"github.com/shirou/gopsutil/v3/host"
	"github.com/shirou/gopsutil/v3/mem"
	"github.com/shirou/gopsutil/v3/process"

	"github.com/pelletier/go-toml/v2"
	"github.com/spf13/pflag"
)

//go:embed www-interface/*
var wwwInterfaceFS embed.FS

var Version = "0.0.0.0"

// GlobalConfig maps the [global] configuration section.
type GlobalConfig struct {
	DefaultCharset            string   `toml:"default_charset" comment:"default_charset is the character encoding used by the server/fastcgi when serving content. It ensures that text is displayed correctly in browsers. UTF-8 is a common choice as it supports a wide range of characters from different languages. Our implementation does not support other charsets, but this setting is here for compatibility with existing ASP applications that may convert characters to the specified charset before sending them to the client, so we allow it to be set in the headers of the response, but it will not affect the actual encoding used by the server, which is always UTF-8. CLI will always use UTF-8 encoding regardless of this setting."`
	DefaultMslcid             int      `toml:"default_mslcid" comment:"default_mslcid is the default locale identifier (LCID) used by the server for ASP applications. LCIDs are used to specify the locale settings for an application, such as date and time formats, number formats, and language. The value 1046 corresponds to Portuguese (Brazil). You can change this value to match the desired locale for your applications. For example, 1033 is English (United States), 1031 is German (Germany), etc. A full list of LCIDs can be found on Microsoft's documentation or in mslcid.go file in our source code. The server will also use this to set the default locale for the ASP scripting engine, which can affect how certain functions behave based on the locale settings."`
	DefaultScriptTimeout      int      `toml:"default_script_timeout" comment:"The amount of time in seconds that the server will wait for an ASP script to execute before timing out. If a script takes longer than this time to execute,the server will terminate the script and return an error to the client. You can adjust this value based on the expected execution time of your scripts. A common default is 60 seconds, but you may want to increase it for long-running scripts or decrease it for better performance and resource management."`
	ResponseBufferLimitMB     int      `toml:"response_buffer_limit_mb" comment:"The maximum buffered Response output size in megabytes before AxonASP aborts execution with a runtime error. This protects the server from unbounded in-memory buffering while Response.Buffer is enabled. The limit is applied consistently in HTTP, FastCGI, and CLI execution."`
	DefaultTimezone           string   `toml:"default_timezone" comment:"The timezone setting specifies the default timezone for the server. This can affect how dates and times are displayed and processed in your ASP applications. Setting it to \"UTC\" means that the server will use Coordinated Universal Time as the default timezone. You can change this to a specific timezone if needed, such as \"America/New_York\" or \"Europe/London\". Make sure to use a valid timezone identifier from the IANA Time Zone Database. You can also use UTC+offsets like \"UTC+2\" or \"UTC-5\" to specify a timezone relative to UTC, but it's generally recommended to use named timezones for better clarity and to account for daylight saving time changes. The server will use this timezone setting for date and time functions in your ASP scripts, as well as for logging and other time-related operations. Note that this setting does not affect the system timezone of the server itself, which may be different from the timezone used by the AxonASP server."`
	EnableASPDebugging        bool     `toml:"enable_asp_debugging" comment:"When enabled, the http/fastcgi server will provide additional debugging information for ASP scripts, which can be helpful during development. However, it may also expose sensitive information about the server and should be disabled in production environments for security reasons. The CLI will always provide detailed error messages regardless of this setting, as it is intended for development and debugging purposes. ATTENTION: This will also enable go pprof endpoints on the proxy server version, which can be accessed at /debug/pprof and can provide detailed information about the server's performance and resource usage, but it can also pose a security risk if exposed to unauthorized users. If for any reason you need to enable ASP debugging in production, make sure to secure the pprof endpoints properly."`
	EnableLogFiles            bool     `toml:"enable_log_files" comment:"When enabled, the http/fastcgi server will create an error.log/console.log file in ./temp. The CLI will always provide detailed error messages regardless of this setting, as it is intended for development and debugging purposes. This option also enable the loggin of console.log, console.info, console.error and console.warn outputs in the error.log/console.log file, which can be useful for debugging purposes. However, it may also consume disk space over time if there are a lot of errors or console outputs being logged, so it's generally recommended to keep this setting disabled in production environments and only enable it during development or when you need to troubleshoot specific issues with your ASP scripts. Make sure to monitor the size of the error.log file and implement log rotation or cleanup strategies as needed to prevent it from consuming too much disk space."`
	EnableErrorLogFile        bool     `toml:"enable_error_log_file" comment:"This configuration was deprecated in version 2.1, don't use it anymore, use enable_log_files instead. We keep it here for compatibility with existing configuration files."`
	DumpPreprocessedSource    bool     `toml:"dump_preprocessed_source" comment:"When enabled, the server will output in ./temp the full file before the engine compiles it, good for error handling and debugging, but it can consume a lot of disk space if you have a lot of traffic or large scripts, so it's generally recommended to keep this setting disabled in production environments and only enable it during development or when you need to troubleshoot specific issues with your ASP scripts."`
	CleanSessionsOnStartup    bool     `toml:"clean_sessions_on_startup" comment:"When enabled, the server will clear all existing sessions when it starts up. This can help prevent issues with stale or corrupted session data, but it will also log out any users who are currently logged in when the server restarts. You can disable this if you want to preserve sessions across server restarts, but be aware that it may lead to issues if there are problems with the session data and persistense that should not ocurr."`
	BytecodeCachingEnabled    string   `toml:"bytecode_caching_enabled" comment:"When enabled, the server will cache the compiled bytecode of ASP scripts in memory and disk for faster execution on subsequent requests. This can significantly improve performance for frequently accessed scripts, as it avoids the overhead of recompiling the script on each request. Values can be \"enabled\" (default), \"memory-only\", \"disk-only\" or \"disabled\". \"enabled\" will cache compiled scripts in both memory and disk, \"memory-only\" will cache compiled scripts only in memory (tier 1), \"disk-only\" will cache compiled scripts only on disk (tier 2), and \"disabled\" will not cache compiled scripts at all, which can significantly degrade performance but may be useful for development or troubleshooting purposes."`
	CacheMaxSizeMB            int      `toml:"cache_max_size_mb" comment:"The size of the compiled ASP scripts cache in megabytes. When an ASP script is requested for the first time, the server compiles it and stores the compiled version in memory for faster execution on subsequent requests. If the number of cached scripts exceeds this limit, the server will remove the least recently used scripts from the cache to free up memory. Setting this value to a reasonable number can help improve performance by keeping frequently accessed scripts in memory, but setting it too high may lead to increased memory usage, while setting it too low may result in more frequent recompilation of scripts, which can degrade performance. Adjust this value based on the size and traffic of your website, as well as the available memory resources on your server."`
	CleanCacheOnStartup       bool     `toml:"clean_cache_on_startup" comment:"When enabled, the server will clear its cache of compiled ASP scripts when it starts up. This can help ensure that any changes to your ASP files are picked up immediately when the server restarts, but it may also increase the startup time of the server as it needs to recompile all scripts. You can disable this if you want to keep the compiled scripts in cache across restarts for faster startup, but be aware that changes to ASP files may not take effect until the server is restarted again."`
	VMPoolSize                int      `toml:"vm_pool_size" comment:"The size of the pool of virtual machines (VMs) used to execute ASP scripts. Each VM can execute one script at a time, so having a pool of VMs allows the server to handle multiple requests concurrently faster. Setting this value to a bigger number can help improve performance by allowing more scripts to be executed simultaneously, but setting it too high may lead to increased memory usage and resource contention, while setting it too low may result in slower response times during periods of high traffic. Adjust this value based on the expected traffic to your website and the available resources on your server. You should use a number bigger than 1. A pool size of 10 VMs in a server with 512mb of memory can respond to approximately 2000 simultaneous requests for simple pages."`
	GolangMemoryLimitMB       int      `toml:"golang_memory_limit_mb" comment:"The maximum amount of memory in megabytes that the Go runtime is allowed to use. This can help prevent the server from consuming too much memory and potentially crashing. Setting it to 0 means no limit, but it's generally recommended to set a reasonable limit to ensure stability. If your server has limited memory resources, you may want to set this to a lower value to prevent out-of-memory errors. Setting this value too low may lead to performance issues or out-of-memory errors, while setting it too high may allow the server to consume more memory than is available, leading to crashes. Also note that this setting may not be strictly enforced by the Go runtime, and actual memory usage may vary based on the workload and garbage collection behavior. If your server is missing requests, low the vm_pool_size and up the memory limit, as this usually means the requests are getting blocked by the Garbage Collector. This directive is more important than  vm_pool_size, and directly influence how some libraries like zstd work."`
	SessionFlushIntervalSecs  int      `toml:"session_flush_interval_seconds" comment:"Interval in seconds used to asynchronously flush dirty in-memory sessions to disk. A value greater than 0 keeps session writes off the request hot path while still guaranteeing a safe flush on process shutdown."`
	AdodbPlatformArchitecture string   `toml:"adodb_platform_architecture" comment:"The architecture of the platform for which the ADODB library is called from (just for Access Database and in Windows). This is important for compatibility with the database drivers used by your ASP applications. If you are running a 64-bit operating system, you should set this to \"amd64\". If you are running a 32-bit operating system, you should set this to \"386\". You can use the 386 ADODB on your 64-bit Windows server, but you need to install the 32-bit version. You can also set it to \"auto\" to let the server automatically detect the architecture of the platform it is running on."`
	ExecuteAsASP              []string `toml:"execute_as_asp" comment:"List of file extensions that will be treated as ASP scripts and executed by the server. You can add or remove extensions from this list based on your needs. For example, if you want to execute .aspx files as ASP scripts, you can add \".aspx\" to the list. Make sure to include the dot before the extension. The server will check the requested file's extension against this list to determine whether to execute it as an ASP script or serve it as a static file."`
	ExecuteAsVBScript         []string `toml:"execute_as_vbscript" comment:"List of file extensions that will be treated as VBScript and executed by the server. You can add or remove extensions from this list based on your needs. Make sure to include the dot before the extension. The server will check the requested file's extension against this list to determine whether to execute it or serve it as a static file. This will only be used if engine_mode is set to vbscript."`
	ExecuteAsJavaScript       []string `toml:"execute_as_javascript" comment:"List of file extensions that will be treated as JavaScript and executed by the server. You can add or remove extensions from this list based on your needs. Make sure to include the dot before the extension. The server will check the requested file's extension against this list to determine whether to execute it or serve it as a static file. This will only be used if engine_mode is set to javascript."`
	ViperWatchConfig          bool     `toml:"viper_watch_config" comment:"When enabled, the server will watch for changes in the configuration file and automatically reload the configuration without needing to restart the server. This can be useful for making changes to the server settings on the fly, but it may also introduce some overhead as the server needs to monitor the file for changes. It's generally recommended to keep this setting disabled in production environments for better performance and stability, and only enable it during development or when you need to make frequent changes to the configuration. This setting isn't full implented yet."`
	ViperAutomaticEnv         bool     `toml:"viper_automatic_env" comment:"When enabled, the server will automatically read configuration values from environment variables that match the settings in this configuration file. This allows you to easily override settings without modifying the configuration file directly, which can be especially useful in containerized environments or when using a secrets management solution. The environment variables should be in uppercase and use underscores instead of dots. For example, to override the default_charset setting, you would set an environment variable named DEFAULT_CHARSET with the desired value. This provides flexibility in managing configurations across different environments (development, staging, production) without changing the code or configuration files."`
	TempDir                   string   `toml:"temp_dir" comment:"Directory for temporary files used by the engine. This directory is used for storing temporary files created during the execution of ASP scripts, such as session data, cached compiled scripts, and other temporary resources. Make sure this directory is writable by the server process and has sufficient space to accommodate the temporary files generated by your applications. You can change this path to a different directory if needed, but ensure that it is properly secured and not accessible to unauthorized users."`
}

// CliConfig maps the [cli] configuration section.
type CliConfig struct {
	EnableCli                   bool   `toml:"enable_cli" comment:"When enabled, the server will allow the TUI interface that can be used to test ASP scripts and VBScript. However, it can also pose a security risk if not used carefully, as it allows any user with access to the CLI to execute scripts in TUI. It's generally recommended to keep this setting disabled unless you have a specific use case that requires it and you trust the scripts that will be using this functionality. This setting need to be enabled for enable_cli_run_from_command_line to work."`
	EnableCliRunFromCommandLine bool   `toml:"enable_cli_run_from_command_line" comment:"When enabled, you can run ASP scripts directly from the command line using \"axonasp-cli.exe -r/--run script.asp\", which can be useful for running maintenance tasks or scheduled jobs without needing to access them through the web server, but it can also pose a security risk if not used carefully, as it allows any script with access to the CLI to be executed. It's generally recommended to keep this setting disabled unless you have a specific use case that requires it and you trust the scripts that will be using this functionality."`
	ForceFreshCompile           bool   `toml:"force_fresh_compile" comment:"When true, CLI execution always recompiles scripts as fresh runs and bypasses bytecode caching. Set to false to allow CLI to follow global.bytecode_caching_enabled behavior. When using TUI it's generally recommended to set this to true to ensure that you are always running the latest version of your scripts and to avoid any potential issues with stale bytecode. If you're using the CLI to run scripts (-r) directly from the command line, you can set this to false to take advantage of bytecode caching for improved performance, but be aware that changes to your scripts may not take effect until the bytecode cache is refreshed or cleared."`
	EngineMode                  string `toml:"engine_mode" comment:"Set the engine mode. Can be set to: default to execute ASP pages/code, vbscript to execute vbscript only or javascript, to execute javascript only. This will also change the way the testsuite works."`
}

// ServerConfig maps the [server] configuration section.
type ServerConfig struct {
	DefaultErrorPagesDirectory string   `toml:"default_error_pages_directory" comment:"The directory where the server will look for default error pages. When an error occurs (e.g., 404 Not Found, 500 Internal Server Error), the server will check this directory for corresponding error page files (e.g., 404.html, 500.asp) and serve them to the client. If no custom error page is found, the server will return a default error message. You can customize this directory and the error pages to provide a better user experience when errors occur on your website. This configuration may be overridden by settings in the web.config file of your ASP application, allowing you to specify different error pages for different applications or directories."`
	WebRoot                    string   `toml:"web_root" comment:"The root directory for the web server. This is the base directory from which the server will serve files. When a client makes a request, the server will look for the requested file within this directory. For example, if the web_root is set to \"./www\" and a client requests \"/index.html\", the server will look for \"./www/index.html\". Make sure to set this to the correct path where your ASP applications and static files are located. This configuration can't be be overridden by settings in the web.config file of your ASP application."`
	DefaultPages               []string `toml:"default_pages" comment:"List of default pages to try when a directory is accessed. The server will look for these files in order and serve the first one it finds."`
	ServerPort                 int      `toml:"server_port" comment:"The port on which the web server will listen and you should redirect your reverse proxy."`
	BlockedExtensions          []string `toml:"blocked_extensions" comment:"List of file extensions that the server will block from being served. This is a security measure to prevent access to sensitive files that should not be exposed to clients. The server will return a 404 Not found error if a client tries to access a file with one of these extensions. You can customize this list based on the types of files you want to protect. It's important to include any file types that may contain sensitive information or executable code that should not be accessible through the web server."`
	BlockedFiles               []string `toml:"blocked_files" comment:"List of files that the server will block from being served directly. This is a security measure to prevent access to sensitive files that should not be exposed to clients. The server will return a 404 Not found error if a client tries to access the file. You can customize this list based on the types of files you want to protect. It's important to include any file types that may contain sensitive information or executable code that should not be accessible through the web server."`
	BlockedDirs                []string `toml:"blocked_dirs" comment:"List of files that the server will block from being served directly. This is a security measure to prevent access to sensitive files that should not be exposed to clients. The server will return a 404 Not found error if a client tries to access the file. You can customize this list based on the types of files you want to protect. It's important to include any file types that may contain sensitive information or executable code that should not be accessible through the web server."`
	EnableWebConfig            bool     `toml:"enable_webconfig" comment:"Allows web.config files to override certain settings for the web server. When set to true, the server will read web.config files in the directories of the web root and apply some settings specified in those files. This allows for more granular control over the behavior of the server for specific applications. It won't work for directories outside the root. For example, you can use web.config files to specify custom error pages, and implement redirections and virtual directories. If set to false, the server will ignore any web.config file and use only the settings specified in this main configuration file."`
	EnableDirectoryListing     bool     `toml:"enable_directory_listing" comment:"When enabled, the server will allow directory listing for directories that do not contain a default page. This means that if a client requests a directory and there is no default page (e.g., index.html) in that directory, the server will return a list of the files and subdirectories within that directory. This can be useful for development and debugging purposes, but it can also pose a security risk if sensitive files are exposed. It's generally recommended to keep this setting disabled in production environments to prevent unauthorized access to directory contents."`
	DirectoryListingTemplate   string   `toml:"directory_listing_template" comment:"The path to the HTML template used for directory listing when enable_directory_listing is set to true. This template should include placeholders (see the default directory listing template) where the server will inject the list of files and directories. You can customize this template to match the design of your website and provide a better user experience when directory listing is enabled. Make sure to set this to the correct path where your custom directory listing template is located."`
	EngineMode                 string   `toml:"engine_mode" comment:"Set the engine mode. Can be set to: default to execute ASP pages/code, vbscript to execute vbscript only or javascript, to execute javascript only."`
}

// FastcgiConfig maps the [fastcgi] configuration section.
type FastcgiConfig struct {
	DefaultPages []string `toml:"default_pages" comment:"List of default pages to try when a directory is accessed. The server will look for these files in order and serve the first one it finds when requested for a directory. This is similar to the default_pages setting in the [server] section, but it applies specifically to the FastCGI server. You can customize this list based on the default pages you want to serve for directories when using FastCGI."`
	ServerPort   int      `toml:"server_port" comment:"Set the port number to the fastcgi server. Can also be a path to socket, e.g. \"unix:/tmp/axonasp.sock\" on *nix systems"`
	EngineMode   string   `toml:"engine_mode" comment:"Set the engine mode. Can be set to: default to execute ASP pages/code, vbscript to execute vbscript only or javascript, to execute javascript only."`
}

// G3dbConfig maps the [g3db] configuration section.
type G3dbConfig struct {
	MysqlDatabase     string `toml:"mysql_database" comment:"MySQL Database Configuration (G3DB)"`
	MysqlHost         string `toml:"mysql_host"`
	MysqlPass         string `toml:"mysql_pass"`
	MysqlPort         int    `toml:"mysql_port"`
	MysqlUser         string `toml:"mysql_user"`
	PostgresHost      string `toml:"postgres_host" comment:"PostgreSQL Database Configuration (G3DB)"`
	PostgresUser      string `toml:"postgres_user"`
	PostgressDatabase string `toml:"postgress_database"`
	PostgressPass     string `toml:"postgress_pass"`
	PostgressPort     int    `toml:"postgress_port"`
	PostgressSslMode  string `toml:"postgress_ssl_mode"`
	MssqlDatabase     string `toml:"mssql_database" comment:"MS SQL Server Database Configuration (G3DB)"`
	MssqlHost         string `toml:"mssql_host"`
	MssqlPass         string `toml:"mssql_pass"`
	MssqlPort         int    `toml:"mssql_port"`
	MssqlUser         string `toml:"mssql_user"`
	SqliteBusyTimeout int    `toml:"sqlite_busy_timeout" comment:"SQLite Database Configuration (G3DB)"`
	SqlitePath        string `toml:"sqlite_path"`
	OracleDsn         string `toml:"oracle_dsn" comment:"Oracle Database Configuration (G3DB)\nProvide either oracle_dsn (a full go-ora/v2 URL) or individual host/port/user/pass/service keys.\nDSN format: oracle://user:password@host:port/service_name"`
	OracleHost        string `toml:"oracle_host"`
	OraclePass        string `toml:"oracle_pass"`
	OraclePort        int    `toml:"oracle_port"`
	OracleService     string `toml:"oracle_service"`
	OracleUser        string `toml:"oracle_user"`
}

// G3mailConfig maps the [g3mail] configuration section.
type G3mailConfig struct {
	SmptHost string `toml:"smpt_host" comment:"Mail configuration for AxonASP Server. This section contains settings related to sending emails from your ASP applications, such as SMTP server details and authentication credentials. Properly configuring these settings is essential for ensuring that your applications can send emails successfully, whether it's for user notifications, password resets, or other email functionalities. Adjust these settings according to your email service provider's requirements and the needs of your applications. This configuration is used by the built-in mail functionality in the ASP scripting engine, which allows you to send emails using any of the g3mail object in your ASP scripts. For better security, it's recommended to use environment variables or a secure secrets management solution to store sensitive information like email credentials instead of hardcoding them in the configuration file, especially in production environments. You can set a .env file in the root of the server executable with the same variables defined here, and the server will load them and override the values in this configuration file, allowing you to keep sensitive information out of your version control system and easily manage different configurations for development and production environments."`
	SmtpFrom string `toml:"smtp_from"`
	SmtpPass string `toml:"smtp_pass"`
	SmtpPort int    `toml:"smtp_port"`
	SmtpUser string `toml:"smtp_user"`
}

// G3axonliveConfig maps the [g3axonlive] configuration section.
type G3axonliveConfig struct {
	G3axonliveActive                 bool `toml:"g3axonlive_active" comment:"G3AXONLIVE Configuration - Reactive Component Framework\nThe G3AXONLIVE library provides a native reactive component framework for building stateful ASP pages\nwhere components update asynchronously without full page reloads. All business logic stays on the server.\nWhen g3axonlive_active is set to false, the library is completely disabled and no resources are allocated.\n\nEnable/Disable the G3AXONLIVE reactive component library. When disabled, the /g3al/ endpoint is not registered\nand the G3AXONLIVE object cannot be created via Server.CreateObject(\"G3AXONLIVE\")."`
	G3axonliveCleanupIntervalMinutes int  `toml:"g3axonlive_cleanup_interval_minutes" comment:"Cleanup interval in minutes for orphaned component states. When a component state has not been accessed\nfor this duration, it will be automatically removed from memory by the background cleanup goroutine.\nThis prevents memory leaks from users who close their browser without properly closing the connection.\nMinimum recommended value is 5 minutes."`
	G3axonliveAutoCleanup            bool `toml:"g3axonlive_auto_cleanup" comment:"Enable automatic background cleanup goroutine for orphaned component states.\nWhen enabled, the G3AXONLIVE library will spawn a background goroutine on first instantiation\nthat periodically removes stale component states. Set to false to disable automatic cleanup and\nmanage cleanup manually via StopCleanup()/StartCleanup() methods."`
	MaxComponentsPerResponse         int  `toml:"max_components_per_response" comment:"Maximum number of component patches allowed per EndAsyncResponse() call.\nIf the server tries to register more patches than this limit, an error is raised.\nIncrease this value if your page renders many independently-updated components at once."`
}

// AxfunctionsConfig maps the [axfunctions] configuration section.
type AxfunctionsConfig struct {
	EnableGlobalAx                 bool   `toml:"enable_global_ax" comment:"When enabled, the Ax functions will be avaliable on the global context, without the need to call Server.Object(). This will break compatibility with Classic ASP default"`
	EnableAxServerShutdownFunction bool   `toml:"enable_axservershutdown_function" comment:"When enabled, the server will allow the use of the axshutdownaxonaspserver() function in ASP scripts, which can be used to programmatically shut down the server. This can be useful for maintenance or when you want to allow certain scripts to trigger a server shutdown, but it can also pose a severe security risk if not used carefully, as it allows any script with access to this function to shut down the server. It's generally recommended to keep this setting disabled unless you have a specific use case that requires it and you trust the scripts that will be using this function."`
	AxDefaultCssPath               string `toml:"ax_default_css_path" comment:"The path to the CSS file used by the built-in AxonASP pages (e.g., error pages). This allows you to customize the appearance of these pages by providing your own CSS file."`
	AxDefaultLogoPath              string `toml:"ax_default_logo_path" comment:"The path to the logo used by the built-in AxonASP pages (e.g., error pages). This allows you to customize the appearance of these pages by providing your own logo file. It will be returned as an inline base64 image, so you can use any image format supported by browsers (e.g., PNG, JPEG, SVG) and it will be displayed correctly on the pages."`
}

// McpConfig maps the [mcp] configuration section.
type McpConfig struct {
	McpMode    string `toml:"mcp_mode" comment:"You can set it to \"stdio\" or \"sse\", if sse, set mcp_sse_port to the desired port. The SSE will be served at http://localhost:{mcp_sse_port}/sse, and you can send commands to http://localhost:{mcp_sse_port}/command. This setting determines the mode of communication for the MCP (Management Control Panel) tool. When set to \"stdio\", the MCP will communicate through standard input and output, which is suitable for local management and debugging. When set to \"sse\", the MCP will use Server-Sent Events (SSE) to provide a real-time interface for managing the server, which can be accessed remotely through a web browser or other SSE-compatible client. The SSE mode allows for more interactive and dynamic management of the server, but it may require additional configuration and security considerations if exposed to remote clients. If you opt for the SSE mode, make sure to secure the endpoint properly using a reverse proxy like nginx."`
	McpSsePort int    `toml:"mcp_sse_port" comment:"The port on which the MCP SSE server will listen if mcp_mode is set to \"sse\". This allows you to use the MCP functionality over SSE, which can be useful for remote management and integration with other tools. If mcp_mode is set to \"stdio\", this setting will be ignored and the MCP will communicate through standard input and output instead."`
	McpDocs    string `toml:"mcp_docs" comment:"The path to the markdown file that contains the documentation for the MCP tool. This file should be formatted in a very specific way that allows the MCP to parse it and extract relevant information based on user queries. The documentation should include details about the available functions, objects, libraries, and other resources in the AxonASP server, along with examples and usage instructions. Properly maintaining this documentation is crucial for ensuring that users can effectively utilize the MCP tool to get accurate information about the server's capabilities and how to use them in their ASP applications. You can customize this path based on where you store your documentation file, but make sure it is accessible by the server and properly formatted for parsing."`
}

// MswcConfig maps the [mswc] configuration section.
type MswcConfig struct {
	PagecounterEnabled             bool   `toml:"pagecounter_enabled" comment:"When enabled, the MSWC.PageCounter component will be available for use in ASP scripts. This component allows you to easily track and display the number of hits or visits to a page. When enabled, the server will read and update the hit count from the file specified in pagecounter_file whenever the MSWC.PageCounter component is used in an ASP script. This can be useful for tracking page popularity and visitor engagement on your website. However, it may also introduce some overhead due to file I/O operations, especially on high-traffic websites. It will start a goroutine to save the hit count to the file at regular intervals defined by pagecounter_save_interval_seconds, which can help mitigate performance issues by reducing the frequency of file writes and only if memory values changed. It's generally recommended to enable this setting only if you need the page hit counting functionality provided by the MSWC.PageCounter component, and to monitor the performance impact on your server if you have a high-traffic website."`
	PagecounterFile                string `toml:"pagecounter_file" comment:"The path to the file used by the MSWC.PageCounter component to store the hit count. This file will be created if it does not exist, and the server will read and update the hit count in this file whenever the MSWC.PageCounter component is used in an ASP script. Make sure that the server has write permissions to this file and its directory, as it needs to update the hit count each time the component is used. You can customize this path based on your preferences, but it's generally recommended to keep it within a directory that is not publicly accessible through the web server for security reasons. This is a Go-specific binary (gob) file that is used to efficiently store the hit count data, and it is not meant to be edited manually."`
	PagecounterSaveIntervalSeconds int    `toml:"pagecounter_save_interval_seconds" comment:"The interval in seconds at which the MSWC.PageCounter component will save the hit count to the file specified in pagecounter_file. This setting helps to reduce the frequency of file writes, which can improve performance, especially on high-traffic websites. The server will keep the hit count in memory and only write it to the file at the specified intervals. You can adjust this interval based on your needs and the expected traffic to your website. A shorter interval will provide more up-to-date hit counts but may increase disk I/O, while a longer interval will reduce disk I/O but may result in less accurate hit counts if the server is restarted or crashes before the next save."`
}

// ServiceConfig maps the [service] configuration section.
type ServiceConfig struct {
	ServiceName                 string   `toml:"service_name" comment:"These settings are only relevant when running the server in service mode using the service wrapper, and it will be ignored when running in normal mode. \n\nThe name of the service when running in service mode. This is used to identify the service in the operating system's service manager (e.g., Windows Services). You can set this to a descriptive name that reflects the purpose of the service, such as \"AxonASP Server\". Make sure to choose a unique name if you have multiple services running on the same machine to avoid conflicts."`
	ServiceDisplayName          string   `toml:"service_display_name" comment:"The display name of the service when running in service mode. This is used to provide a more user-friendly name for the service in the operating system's service manager. You can set this to a descriptive name that reflects the purpose of the service, such as \"AxonASP Server\". This is the name that will be shown in the list of services, so it can be more descriptive than the actual service name used for identification."`
	ServiceDescription          string   `toml:"service_description" comment:"The description of the service when running in service mode. This is used to provide additional information about the service in the operating system's service manager. You can set this to a brief description that explains what the service does, such as \"AxonASP Service running AxonASP Server. This is a wrapper to our server.\""`
	ServiceExecutablePath       string   `toml:"service_executable_path" comment:"Path to the executable that will be run when the service starts, it can be the same as the main server executable or the fast-cgi version. Make sure to set this to the correct path of the executable you want to run when the service starts. Don't append the file extension, the service wrapper will automatically add .exe on Windows."`
	ServiceEnvironmentVariables []string `toml:"service_environment_variables" comment:"List of environment variables to set for the service, in the format \"KEY=VALUE\". These environment variables will be available to the service process when it runs, allowing you to configure the behavior of the server or your ASP applications without modifying the configuration file directly. You can add as many environment variables as needed, and they will be set in the service's environment when it starts. This is especially useful for setting sensitive information like database credentials or API keys in a secure way, such as using a secrets management solution or environment variable management in your deployment platform."`
}

// JavascriptConfig maps the [javascript] configuration section.
type JavascriptConfig struct {
	EnableNodeCompatibility bool `toml:"enable_node_compatibility" comment:"Enable the compatibility mode with Node.js, this is a experimental feature."`
}

// Config is the canonical schema for axonasp.toml.
type Config struct {
	Global      GlobalConfig      `toml:"global"`
	Cli         CliConfig         `toml:"cli"`
	Server      ServerConfig      `toml:"server"`
	Fastcgi     FastcgiConfig     `toml:"fastcgi"`
	G3db        G3dbConfig        `toml:"g3db"`
	G3mail      G3mailConfig      `toml:"g3mail"`
	G3axonlive  G3axonliveConfig  `toml:"g3axonlive"`
	Axfunctions AxfunctionsConfig `toml:"axfunctions"`
	Mcp         McpConfig         `toml:"mcp"`
	Mswc        MswcConfig        `toml:"mswc"`
	Service     ServiceConfig     `toml:"service"`
	Javascript  JavascriptConfig  `toml:"javascript"`
}

// FPMPoolConfig is the canonical schema for a pool file in ./fpm/fpm.d/*.conf.
type FPMPoolConfig struct {
	SiteName      string `toml:"site_name" comment:"Site name used in logs and process titles for this pool."`
	UID           uint32 `toml:"uid" comment:"User ID used to run this worker process. It must have read access to app_path and config_file and write access to socket/tmp_dir when applicable."`
	GID           uint32 `toml:"gid" comment:"Group ID used to run this worker process. It should match the permissions required by socket/tmp_dir and application files."`
	Socket        string `toml:"socket" comment:"FastCGI listener endpoint for this pool. You can use a unix socket path, unix:/path form, host:port, or a plain TCP port."`
	ConfigFile    string `toml:"config_file" comment:"Absolute path to the AxonASP configuration file used by this worker. The FPM supervisor controls the FastCGI endpoint, so fastcgi.server_port inside this file is ignored."`
	GlobalAsaPath string `toml:"global_asa_path" comment:"Optional directory used to resolve global.asa for this pool. Prefer absolute paths in production to avoid path ambiguity."`
	AppPath       string `toml:"app_path" comment:"Working directory (CWD) used when launching this worker process. It should point to the AxonASP application root."`
	MemoryLimitMB int    `toml:"memory_limit_mb" comment:"Per-worker memory limit in MB. The supervisor can restart workers that exceed this value to protect host stability."`
	MaxRestarts   int    `toml:"max_restarts" comment:"Maximum restart attempts for this worker pool. Set to 0 to disable the cap."`
	TmpDir        string `toml:"tmp_dir" comment:"Temporary directory used by the worker process. Ensure write permissions for the configured UID/GID."`
}

// ProcessTelemetry tracks runtime counts and memory usage for monitored executables.
type ProcessTelemetry struct {
	Name           string `json:"name"`
	Count          int    `json:"count"`
	MemoryBytes    uint64 `json:"memory_bytes"`
	ExecutablePath string `json:"executable_path"`
}

// HomeTelemetry contains host and process information rendered on the Home view.
type HomeTelemetry struct {
	HostName           string             `json:"host_name"`
	System             string             `json:"system"`
	Processor          string             `json:"processor"`
	TotalMemoryBytes   uint64             `json:"total_memory_bytes"`
	AvailableMemory    uint64             `json:"available_memory_bytes"`
	UsedMemoryBytes    uint64             `json:"used_memory_bytes"`
	UsedMemoryPercent  float64            `json:"used_memory_percent"`
	DiskTotalBytes     uint64             `json:"disk_total_bytes"`
	DiskFreeBytes      uint64             `json:"disk_free_bytes"`
	TempDirPath        string             `json:"temp_dir_path"`
	TempDirSizeBytes   uint64             `json:"temp_dir_size_bytes"`
	TempDirSizeKnown   bool               `json:"temp_dir_size_known"`
	WebRootPath        string             `json:"web_root_path"`
	WebRootSizeBytes   uint64             `json:"web_root_size_bytes"`
	WebRootSizeKnown   bool               `json:"web_root_size_known"`
	ProcessStats       []ProcessTelemetry `json:"process_stats"`
	ProcessTotalCount  int                `json:"process_total_count"`
	ProcessTotalMemory uint64             `json:"process_total_memory"`
	CollectionWarnings []string           `json:"collection_warnings"`
}

// SectionField describes a single configuration field within a section.
type SectionField struct {
	Key          string
	Type         string
	Description  string
	DefaultValue any
	CurrentValue any
}

// Section describes a configuration section.
type Section struct {
	Name   string
	Fields []SectionField
}

// PageData represents the structure passed to the dashboard template.
type PageData struct {
	Sections      []Section
	ActiveSection Section
	ResolvedPath  string
	ActiveView    string
	IsWindows     bool
	IsGUIEnv      bool
	FPMDir        string
	FPMPools      []string
	ActivePool    string
	FPMFields     []SectionField
	HomeTelemetry HomeTelemetry
}

// NewDefaultConfig instantiates a Config populated with baseline default values.
func NewDefaultConfig() Config {
	return Config{
		Global: GlobalConfig{
			DefaultCharset:            "UTF-8",
			DefaultMslcid:             1033,
			DefaultScriptTimeout:      60,
			ResponseBufferLimitMB:     4,
			DefaultTimezone:           "UTC",
			EnableASPDebugging:        true,
			EnableLogFiles:            true,
			EnableErrorLogFile:        true,
			DumpPreprocessedSource:    false,
			CleanSessionsOnStartup:    true,
			BytecodeCachingEnabled:    "enabled",
			CacheMaxSizeMB:            100,
			CleanCacheOnStartup:       true,
			VMPoolSize:                10,
			GolangMemoryLimitMB:       256,
			SessionFlushIntervalSecs:  120,
			AdodbPlatformArchitecture: "auto",
			ExecuteAsASP:              []string{".asp"},
			ExecuteAsVBScript:         []string{".vbs"},
			ExecuteAsJavaScript:       []string{".js", ".mjs"},
			ViperWatchConfig:          false,
			ViperAutomaticEnv:         true,
			TempDir:                   "./temp",
		},
		Cli: CliConfig{
			EnableCli:                   true,
			EnableCliRunFromCommandLine: true,
			ForceFreshCompile:           true,
			EngineMode:                  "default",
		},
		Server: ServerConfig{
			DefaultErrorPagesDirectory: "./www/error-pages",
			WebRoot:                    "./www/",
			DefaultPages: []string{
				"index.asp",
				"default.asp",
				"index.html",
				"default.html",
				"default.htm",
				"index.htm",
				"home.asp",
				"home.html",
				"home.htm",
				"main.asp",
				"main.html",
				"main.htm",
				"index.txt",
			},
			ServerPort: 8801,
			BlockedExtensions: []string{
				".asax", ".ascx", ".master", ".skin", ".browser", ".sitemap", ".config", ".cs", ".csproj", ".vb", ".vbproj", ".webinfo", ".licx", ".resx", ".resources", ".mdb", ".vjsproj", ".java", ".jsl", ".ldb", ".dsdgm", ".ssdgm", ".lsad", ".ssmap", ".cd", ".dsprototype", ".lsaprototype", ".sdm", ".sdmDocument", ".mdf", ".ldf", ".ad", ".dd", ".ldd", ".sd", ".adprototype", ".lddprototype", ".exclude", ".refresh", ".compiled", ".msgx", ".vsdisco", ".rules", ".asa", ".inc", ".exe", ".dll", ".env", ".htaccess", ".git", ".gitignore", ".seg", ".snp", ".log",
			},
			BlockedFiles: []string{
				"MyInfo.xml",
			},
			BlockedDirs: []string{
				"./www/error-pages",
				"./www/axonasp-pages",
			},
			EnableWebConfig:          true,
			EnableDirectoryListing:   true,
			DirectoryListingTemplate: "./www/axonasp-pages/directory-listing.html",
			EngineMode:               "default",
		},
		Fastcgi: FastcgiConfig{
			DefaultPages: []string{
				"index.asp",
				"default.asp",
				"index.html",
				"default.html",
				"default.htm",
				"index.htm",
				"home.asp",
				"home.html",
				"home.htm",
				"main.asp",
				"main.html",
				"main.htm",
				"index.txt",
			},
			ServerPort: 9000,
			EngineMode: "default",
		},
		G3db: G3dbConfig{
			MysqlDatabase:     "test",
			MysqlHost:         "localhost",
			MysqlPass:         "password",
			MysqlPort:         3306,
			MysqlUser:         "root",
			PostgresHost:      "localhost",
			PostgresUser:      "postgres",
			PostgressDatabase: "test",
			PostgressPass:     "password",
			PostgressPort:     5432,
			PostgressSslMode:  "disable",
			MssqlDatabase:     "test",
			MssqlHost:         "localhost",
			MssqlPass:         "password",
			MssqlPort:         1433,
			MssqlUser:         "sa",
			SqliteBusyTimeout: 5000,
			SqlitePath:        "./database.db",
			OracleDsn:         "",
			OracleHost:        "localhost",
			OraclePass:        "password",
			OraclePort:        1521,
			OracleService:     "ORCLCDB",
			OracleUser:        "system",
		},
		G3mail: G3mailConfig{
			SmptHost: "smtp.example.com",
			SmtpFrom: "sender@example.com",
			SmtpPass: "your_password",
			SmtpPort: 587,
			SmtpUser: "your_email@example.com",
		},
		G3axonlive: G3axonliveConfig{
			G3axonliveActive:                 true,
			G3axonliveCleanupIntervalMinutes: 5,
			G3axonliveAutoCleanup:            true,
			MaxComponentsPerResponse:         200,
		},
		Axfunctions: AxfunctionsConfig{
			EnableGlobalAx:                 true,
			EnableAxServerShutdownFunction: false,
			AxDefaultCssPath:               "./www/axonasp-pages/css/axonasp.css",
			AxDefaultLogoPath:              "./www/axonasp-pages/images/logo.svg",
		},
		Mcp: McpConfig{
			McpMode:    "stdio",
			McpSsePort: 8000,
			McpDocs:    "./www/manual/md/",
		},
		Mswc: MswcConfig{
			PagecounterEnabled:             false,
			PagecounterFile:                "./temp/hitcnt.gob",
			PagecounterSaveIntervalSeconds: 120,
		},
		Service: ServiceConfig{
			ServiceName:                 "AxonASPServer",
			ServiceDisplayName:          "G3pix AxonASP Server",
			ServiceDescription:          "AxonASP Service running AxonASP Server. This is a wrapper used by axonasp-http or axonasp-fastcgi.",
			ServiceExecutablePath:       "./axonasp-http",
			ServiceEnvironmentVariables: []string{},
		},
		Javascript: JavascriptConfig{
			EnableNodeCompatibility: true,
		},
	}
}

// NewDefaultFPMPoolConfig instantiates default values for a new FPM pool file.
func NewDefaultFPMPoolConfig() FPMPoolConfig {
	return FPMPoolConfig{
		SiteName:      "example.com",
		UID:           1001,
		GID:           1001,
		Socket:        "/var/run/axonasp/example.com.sock",
		ConfigFile:    "/opt/axonasp/config/axonasp.toml",
		GlobalAsaPath: "/opt/axonasp/www/",
		AppPath:       "/opt/axonasp/",
		MemoryLimitMB: 128,
		MaxRestarts:   3,
		TmpDir:        "/opt/axonasp/temp",
	}
}

// resolveConfigPath mimics axonconfig/loader.go path candidate resolution logic.
func resolveConfigPath() string {
	configCandidates := []string{
		filepath.Join("config", "axonasp.toml"),
		filepath.Join("..", "config", "axonasp.toml"),
		filepath.Join("..", "..", "config", "axonasp.toml"),
	}
	if executablePath, err := os.Executable(); err == nil {
		configCandidates = append(configCandidates, filepath.Join(filepath.Dir(executablePath), "config", "axonasp.toml"))
	}

	for _, candidate := range configCandidates {
		if _, err := os.Stat(candidate); err == nil {
			if abs, err := filepath.Abs(candidate); err == nil {
				return abs
			}
			return candidate
		}
	}

	abs, _ := filepath.Abs(filepath.Join("config", "axonasp.toml"))
	return abs
}

// resolveFPMConfigDir resolves ./fpm/fpm.d using candidate search similar to config resolution.
func resolveFPMConfigDir() string {
	dirCandidates := []string{
		filepath.Join("fpm", "fpm.d"),
		filepath.Join("..", "fpm", "fpm.d"),
		filepath.Join("..", "..", "fpm", "fpm.d"),
	}
	if executablePath, err := os.Executable(); err == nil {
		dirCandidates = append(dirCandidates, filepath.Join(filepath.Dir(executablePath), "fpm", "fpm.d"))
	}

	for _, candidate := range dirCandidates {
		if info, err := os.Stat(candidate); err == nil && info.IsDir() {
			if abs, err := filepath.Abs(candidate); err == nil {
				return abs
			}
			return candidate
		}
	}

	abs, _ := filepath.Abs(filepath.Join("fpm", "fpm.d"))
	return abs
}

// createNewConfig writes default values with native comments to target path.
func createNewConfig(target string) error {
	dir := filepath.Dir(target)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}
	defaultCfg := NewDefaultConfig()
	data, err := toml.Marshal(defaultCfg)
	if err != nil {
		return err
	}
	return os.WriteFile(target, data, 0644)
}

// createNewFPMConfig writes a default FPM pool file with native comments to target path.
func createNewFPMConfig(target string) error {
	dir := filepath.Dir(target)
	if err := os.MkdirAll(dir, 0755); err != nil {
		return err
	}
	defaultPool := NewDefaultFPMPoolConfig()
	data, err := toml.Marshal(defaultPool)
	if err != nil {
		return err
	}
	return os.WriteFile(target, data, 0644)
}

// listFPMConfigFiles returns all pool *.conf filenames sorted alphabetically.
func listFPMConfigFiles(fpmDir string) ([]string, error) {
	entries, err := os.ReadDir(fpmDir)
	if err != nil {
		return nil, err
	}
	files := make([]string, 0, len(entries))
	for _, entry := range entries {
		if entry.IsDir() {
			continue
		}
		name := entry.Name()
		if strings.EqualFold(filepath.Ext(name), ".conf") {
			files = append(files, name)
		}
	}
	sort.Strings(files)
	return files, nil
}

// normalizeFPMFileName sanitizes a user-provided pool filename and enforces a .conf suffix.
func normalizeFPMFileName(input string) (string, error) {
	name := strings.TrimSpace(input)
	if name == "" {
		return "", fmt.Errorf("filename is required")
	}
	if strings.ContainsAny(name, `/\\`) {
		return "", fmt.Errorf("filename must not contain path separators")
	}
	cleanName := filepath.Base(name)
	if cleanName == "." || cleanName == ".." {
		return "", fmt.Errorf("invalid filename")
	}
	if !strings.EqualFold(filepath.Ext(cleanName), ".conf") {
		cleanName += ".conf"
	}
	return cleanName, nil
}

// getFPMFields reflects over current and default pool configuration to produce schema metadata.
func getFPMFields(current FPMPoolConfig, def FPMPoolConfig) []SectionField {
	valCurrent := reflect.ValueOf(current)
	valDefault := reflect.ValueOf(def)
	typ := reflect.TypeFor[FPMPoolConfig]()

	fields := make([]SectionField, 0, typ.NumField())
	for i := 0; i < typ.NumField(); i++ {
		field := typ.Field(i)
		key := field.Tag.Get("toml")
		if key == "" {
			key = strings.ToLower(field.Name)
		}
		desc := field.Tag.Get("comment")

		currentValue := valCurrent.Field(i).Interface()
		defaultValue := valDefault.Field(i).Interface()

		fieldType := "string"
		switch field.Type.Kind() {
		case reflect.Bool:
			fieldType = "bool"
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
			reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			fieldType = "int"
		}

		fields = append(fields, SectionField{
			Key:          key,
			Type:         fieldType,
			Description:  desc,
			DefaultValue: defaultValue,
			CurrentValue: currentValue,
		})
	}

	return fields
}

// updateFPMField mutates FPM pool configuration settings based on string form payloads.
func updateFPMField(cfg *FPMPoolConfig, fieldTomlName, strVal string) {
	val := reflect.ValueOf(cfg).Elem()
	typ := val.Type()
	for i := 0; i < val.NumField(); i++ {
		field := typ.Field(i)
		key := field.Tag.Get("toml")
		if key == "" {
			key = strings.ToLower(field.Name)
		}
		if !strings.EqualFold(key, fieldTomlName) {
			continue
		}

		subVal := val.Field(i)
		switch subVal.Kind() {
		case reflect.String:
			subVal.SetString(strVal)
		case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
			if intVal, err := strconv.ParseInt(strVal, 10, 64); err == nil {
				subVal.SetInt(intVal)
			}
		case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
			if uintVal, err := strconv.ParseUint(strVal, 10, 64); err == nil {
				subVal.SetUint(uintVal)
			}
		case reflect.Bool:
			subVal.SetBool(strVal == "true" || strVal == "on")
		}
		return
	}
}

// collectHomeTelemetry gathers host memory/system details and monitored process stats.
func collectHomeTelemetry(resolvedConfigPath string) HomeTelemetry {
	telemetry := HomeTelemetry{}
	activeCfg := NewDefaultConfig()
	if loadedCfg, err := loadActiveConfigForTelemetry(resolvedConfigPath); err == nil {
		activeCfg = loadedCfg
	} else {
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "config read fallback to defaults: "+err.Error())
	}

	telemetry.TempDirPath = resolveConfigRelativePath(resolvedConfigPath, activeCfg.Global.TempDir)
	telemetry.WebRootPath = resolveConfigRelativePath(resolvedConfigPath, activeCfg.Server.WebRoot)

	hostName, err := os.Hostname()
	if err == nil {
		telemetry.HostName = hostName
	} else {
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "hostname unavailable: "+err.Error())
	}

	if hostInfo, err := host.Info(); err == nil {
		sysName := strings.TrimSpace(strings.TrimSpace(hostInfo.Platform + " " + hostInfo.PlatformVersion))
		if sysName == "" {
			sysName = runtime.GOOS
		}
		telemetry.System = sysName
	} else {
		telemetry.System = runtime.GOOS
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "system info unavailable: "+err.Error())
	}

	telemetry.Processor = runtime.GOARCH
	if cpuInfo, err := cpu.Info(); err == nil && len(cpuInfo) > 0 {
		if strings.TrimSpace(cpuInfo[0].ModelName) != "" {
			telemetry.Processor = cpuInfo[0].ModelName
		}
	} else if err != nil {
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "cpu info unavailable: "+err.Error())
	}

	if vmStats, err := mem.VirtualMemory(); err == nil {
		telemetry.TotalMemoryBytes = vmStats.Total
		telemetry.AvailableMemory = vmStats.Available
		telemetry.UsedMemoryBytes = vmStats.Used
		telemetry.UsedMemoryPercent = vmStats.UsedPercent
	} else {
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "memory info unavailable: "+err.Error())
	}

	if usage, err := disk.Usage(diskUsageTarget(resolvedConfigPath)); err == nil {
		telemetry.DiskTotalBytes = usage.Total
		telemetry.DiskFreeBytes = usage.Free
	} else {
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "disk usage unavailable: "+err.Error())
	}

	if tempSize, err := directorySize(telemetry.TempDirPath); err == nil {
		telemetry.TempDirSizeBytes = tempSize
		telemetry.TempDirSizeKnown = true
	} else {
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "temp_dir size unavailable: "+err.Error())
	}

	if webRootSize, err := directorySize(telemetry.WebRootPath); err == nil {
		telemetry.WebRootSizeBytes = webRootSize
		telemetry.WebRootSizeKnown = true
	} else {
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "web_root size unavailable: "+err.Error())
	}

	targetMap := map[string]*ProcessTelemetry{
		"axonasp-http":    {Name: "axonasp-http", ExecutablePath: resolveManagedExecutablePath("axonasp-http")},
		"axonasp-fastcgi": {Name: "axonasp-fastcgi", ExecutablePath: resolveManagedExecutablePath("axonasp-fastcgi")},
		"axonasp-fpm":     {Name: "axonasp-fpm", ExecutablePath: resolveManagedExecutablePath("axonasp-fpm")},
		"axonasp-service": {Name: "axonasp-service", ExecutablePath: resolveManagedExecutablePath("axonasp-service")},
		"axonasp-cli":     {Name: "axonasp-cli", ExecutablePath: resolveManagedExecutablePath("axonasp-cli")},
	}

	processes, err := process.Processes()
	if err != nil {
		telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "process list unavailable: "+err.Error())
	} else {
		for _, proc := range processes {
			name, nameErr := proc.Name()
			if nameErr != nil {
				continue
			}
			normalized := strings.ToLower(strings.TrimSuffix(name, filepath.Ext(name)))
			entry, ok := targetMap[normalized]
			if !ok {
				continue
			}
			entry.Count++
			if memInfo, memErr := proc.MemoryInfo(); memErr == nil && memInfo != nil {
				entry.MemoryBytes += memInfo.RSS
			}
		}
	}

	telemetry.ProcessStats = []ProcessTelemetry{
		*targetMap["axonasp-http"],
		*targetMap["axonasp-fastcgi"],
		*targetMap["axonasp-fpm"],
		*targetMap["axonasp-service"],
		*targetMap["axonasp-cli"],
	}

	for _, processStat := range telemetry.ProcessStats {
		telemetry.ProcessTotalCount += processStat.Count
		telemetry.ProcessTotalMemory += processStat.MemoryBytes
	}

	return telemetry
}

// loadActiveConfigForTelemetry loads the active TOML file used by the admin session.
func loadActiveConfigForTelemetry(configPath string) (Config, error) {
	cfg := NewDefaultConfig()
	data, err := os.ReadFile(configPath)
	if err != nil {
		return cfg, err
	}
	if err := toml.Unmarshal(data, &cfg); err != nil {
		return cfg, err
	}
	return cfg, nil
}

// resolveConfigRelativePath resolves a TOML path relative to the active config directory.
func resolveConfigRelativePath(configPath, rawPath string) string {
	trimmed := strings.TrimSpace(rawPath)
	if trimmed == "" {
		return ""
	}
	if filepath.IsAbs(trimmed) {
		return filepath.Clean(trimmed)
	}

	// Match runtime behavior first: relative paths in config are effectively resolved from process CWD.
	if cwd, err := os.Getwd(); err == nil && strings.TrimSpace(cwd) != "" {
		candidate := filepath.Clean(filepath.Join(cwd, trimmed))
		if _, statErr := os.Stat(candidate); statErr == nil {
			return candidate
		}
	}

	// Secondary fallback: resolve from admin executable directory.
	execCandidate := filepath.Clean(filepath.Join(getAdminExecutableDir(), trimmed))
	if _, err := os.Stat(execCandidate); err == nil {
		return execCandidate
	}

	// Last resort: resolve from active config file directory.
	baseDir := filepath.Dir(configPath)
	return filepath.Clean(filepath.Join(baseDir, trimmed))
}

// diskUsageTarget returns an OS-appropriate disk usage root for host telemetry.
func diskUsageTarget(configPath string) string {
	if runtime.GOOS == "windows" {
		vol := filepath.VolumeName(configPath)
		if vol == "" {
			return "C:\\"
		}
		return vol + "\\"
	}
	return "/"
}

// directorySize computes a best-effort recursive directory byte size.
func directorySize(dirPath string) (uint64, error) {
	if strings.TrimSpace(dirPath) == "" {
		return 0, fmt.Errorf("path is empty")
	}
	info, err := os.Stat(dirPath)
	if err != nil {
		return 0, err
	}
	if !info.IsDir() {
		return uint64(info.Size()), nil
	}

	var total uint64
	var firstErr error
	walkErr := filepath.WalkDir(dirPath, func(_ string, d fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			if firstErr == nil {
				firstErr = walkErr
			}
			return nil
		}
		if d.IsDir() {
			return nil
		}
		entryInfo, entryErr := d.Info()
		if entryErr != nil {
			if firstErr == nil {
				firstErr = entryErr
			}
			return nil
		}
		if entryInfo.Size() > 0 {
			total += uint64(entryInfo.Size())
		}
		return nil
	})
	if walkErr != nil {
		return total, walkErr
	}
	if firstErr != nil {
		return total, firstErr
	}
	return total, nil
}

// formatBytes returns a user-facing memory size string from bytes.
func formatBytes(size uint64) string {
	const unit = 1024
	if size < unit {
		return fmt.Sprintf("%d B", size)
	}
	div, exp := uint64(unit), 0
	for n := size / unit; n >= unit; n /= unit {
		div *= unit
		exp++
	}
	suffix := []string{"KB", "MB", "GB", "TB"}
	if exp >= len(suffix) {
		exp = len(suffix) - 1
	}
	return fmt.Sprintf("%.2f %s", float64(size)/float64(div), suffix[exp])
}

// getAdminExecutableDir returns the directory where axonadmin is running from.
func getAdminExecutableDir() string {
	execPath, err := os.Executable()
	if err != nil {
		cwd, cwdErr := os.Getwd()
		if cwdErr != nil {
			return "."
		}
		return cwd
	}
	return filepath.Dir(execPath)
}

// resolveManagedExecutablePath resolves the managed process executable in the admin executable directory.
func resolveManagedExecutablePath(processName string) string {
	baseDir := getAdminExecutableDir()
	fileName := processName
	if runtime.GOOS == "windows" {
		fileName += ".exe"
	}
	return filepath.Join(baseDir, fileName)
}

// findManagedProcesses returns all running process handles that match the managed name.
func findManagedProcesses(processName string) ([]*process.Process, error) {
	allProcs, err := process.Processes()
	if err != nil {
		return nil, err
	}
	matches := make([]*process.Process, 0)
	for _, proc := range allProcs {
		name, nameErr := proc.Name()
		if nameErr != nil {
			continue
		}
		normalized := strings.ToLower(strings.TrimSuffix(name, filepath.Ext(name)))
		if normalized == strings.ToLower(processName) {
			matches = append(matches, proc)
		}
	}
	return matches, nil
}

// startManagedProcess starts one managed executable instance with optional CLI arguments.
func startManagedProcess(processName, rawArgs string) error {
	parsedArgs, err := parseCommandLineArgs(rawArgs)
	if err != nil {
		return fmt.Errorf("failed to parse parameters: %w", err)
	}

	executablePath := resolveManagedExecutablePath(processName)
	if _, statErr := os.Stat(executablePath); statErr != nil {
		if os.IsNotExist(statErr) {
			return fmt.Errorf("executable not found: %s", executablePath)
		}
		return fmt.Errorf("unable to access executable: %w", statErr)
	}

	if processName == "axonasp-cli" {
		if isGUIEnvironment() {
			return startCLIProcessVisibleTerminal(executablePath, parsedArgs)
		}
		if len(parsedArgs) == 0 {
			return fmt.Errorf("parameters are required for axonasp-cli in headless environments")
		}
	}

	cmd := exec.Command(executablePath, parsedArgs...)
	cmd.Dir = filepath.Dir(executablePath)
	configureDetachedProcess(cmd)
	if err := cmd.Start(); err != nil {
		return fmt.Errorf("failed to start process %s: %w", processName, err)
	}
	return nil
}

// isGUIEnvironment determines whether the host can spawn a visible terminal window.
func isGUIEnvironment() bool {
	return runtime.GOOS == "windows" || runtime.GOOS == "darwin"
}

// startCLIProcessVisibleTerminal launches axonasp-cli in a visible terminal for GUI environments.
func startCLIProcessVisibleTerminal(executablePath string, args []string) error {
	switch runtime.GOOS {
	case "windows":
		cmd := exec.Command(executablePath, args...)
		cmd.Dir = filepath.Dir(executablePath)
		configureVisibleConsoleProcess(cmd)
		if err := cmd.Start(); err != nil {
			return fmt.Errorf("failed to start visible terminal: %w", err)
		}
		return nil
	case "darwin":
		var commandLine strings.Builder
		commandLine.WriteString(shellQuoteForSh(executablePath))
		for _, arg := range args {
			commandLine.WriteString(" ")
			commandLine.WriteString(shellQuoteForSh(arg))
		}
		script := fmt.Sprintf("tell application \"Terminal\" to do script %q", commandLine.String())
		if err := exec.Command("osascript", "-e", script).Start(); err != nil {
			return fmt.Errorf("failed to open Terminal.app: %w", err)
		}
		return nil
	default:
		return fmt.Errorf("visible terminal mode unsupported on %s", runtime.GOOS)
	}
}

// shellQuoteForSh quotes a single argument for POSIX shell command composition.
func shellQuoteForSh(raw string) string {
	if raw == "" {
		return "''"
	}
	return "'" + strings.ReplaceAll(raw, "'", "'\"'\"'") + "'"
}

// parseCommandLineArgs splits a shell-like argument string, preserving quoted values.
func parseCommandLineArgs(input string) ([]string, error) {
	trimmed := strings.TrimSpace(input)
	if trimmed == "" {
		return []string{}, nil
	}

	var args []string
	var current strings.Builder
	inSingleQuote := false
	inDoubleQuote := false
	escapeNext := false

	for _, ch := range trimmed {
		switch {
		case escapeNext:
			current.WriteRune(ch)
			escapeNext = false
		case ch == '\\':
			escapeNext = true
		case ch == '\'' && !inDoubleQuote:
			inSingleQuote = !inSingleQuote
		case ch == '"' && !inSingleQuote:
			inDoubleQuote = !inDoubleQuote
		case (ch == ' ' || ch == '\t') && !inSingleQuote && !inDoubleQuote:
			if current.Len() > 0 {
				args = append(args, current.String())
				current.Reset()
			}
		default:
			current.WriteRune(ch)
		}
	}

	if escapeNext || inSingleQuote || inDoubleQuote {
		return nil, fmt.Errorf("unterminated quoted argument")
	}
	if current.Len() > 0 {
		args = append(args, current.String())
	}
	return args, nil
}

// getManagedProcessHelp returns help output from a managed executable for use in the parameters modal.
func getManagedProcessHelp(processName string) (string, error) {
	executablePath := resolveManagedExecutablePath(processName)
	if _, err := os.Stat(executablePath); err != nil {
		if os.IsNotExist(err) {
			return "", fmt.Errorf("executable not found: %s", executablePath)
		}
		return "", err
	}

	helpText, err := captureHelpOutput(executablePath, "-h")
	if strings.TrimSpace(helpText) != "" {
		return helpText, nil
	}

	helpTextAlt, altErr := captureHelpOutput(executablePath, "--help")
	if strings.TrimSpace(helpTextAlt) != "" {
		return helpTextAlt, nil
	}

	if err != nil {
		return "", err
	}
	if altErr != nil {
		return "", altErr
	}
	return "No help output was returned by this executable.", nil
}

// captureHelpOutput executes a managed executable help flag with timeout and returns combined output.
func captureHelpOutput(executablePath, flagArg string) (string, error) {
	ctx, cancel := context.WithTimeout(context.Background(), 4*time.Second)
	defer cancel()

	cmd := exec.CommandContext(ctx, executablePath, flagArg)
	cmd.Dir = filepath.Dir(executablePath)
	out, err := cmd.CombinedOutput()

	outputText := strings.TrimSpace(string(out))
	if len(outputText) > 18000 {
		outputText = outputText[:18000] + "\n... (output truncated)"
	}

	if ctx.Err() == context.DeadlineExceeded {
		if outputText != "" {
			return outputText, nil
		}
		return "", fmt.Errorf("help command timed out")
	}

	if err != nil && outputText == "" {
		return "", err
	}
	return outputText, nil
}

// stopManagedProcess terminates all running instances for a managed process family.
func stopManagedProcess(processName string) (int, error) {
	procs, err := findManagedProcesses(processName)
	if err != nil {
		return 0, fmt.Errorf("failed to enumerate processes: %w", err)
	}
	if len(procs) == 0 {
		return 0, fmt.Errorf("no running instances found for %s", processName)
	}

	stopped := 0
	for _, proc := range procs {
		if killErr := proc.Kill(); killErr != nil {
			return stopped, fmt.Errorf("failed to stop %s (pid %d): %w", processName, proc.Pid, killErr)
		}
		stopped++
	}
	return stopped, nil
}

// getSections reflects over current and default configuration to produce structured schemas.
func getSections(current Config, def Config) []Section {
	valCurrent := reflect.ValueOf(current)
	valDefault := reflect.ValueOf(def)
	typ := reflect.TypeFor[Config]()

	sections := make([]Section, 0, typ.NumField())
	for i := 0; i < typ.NumField(); i++ {
		field := typ.Field(i)
		sectionName := field.Tag.Get("toml")
		if sectionName == "" {
			sectionName = strings.ToLower(field.Name)
		}

		secValCurrent := valCurrent.Field(i)
		secValDefault := valDefault.Field(i)
		secTyp := field.Type

		var fields []SectionField
		for j := 0; j < secTyp.NumField(); j++ {
			subField := secTyp.Field(j)
			key := subField.Tag.Get("toml")
			if key == "" {
				key = strings.ToLower(subField.Name)
			}
			desc := subField.Tag.Get("comment")

			currentSubVal := secValCurrent.Field(j).Interface()
			defaultSubVal := secValDefault.Field(j).Interface()

			fieldType := "string"
			switch subField.Type.Kind() {
			case reflect.Bool:
				fieldType = "bool"
			case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
				reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
				fieldType = "int"
			case reflect.Slice:
				fieldType = "slice"
			}

			fields = append(fields, SectionField{
				Key:          key,
				Type:         fieldType,
				Description:  desc,
				DefaultValue: defaultSubVal,
				CurrentValue: currentSubVal,
			})
		}

		sections = append(sections, Section{
			Name:   sectionName,
			Fields: fields,
		})
	}
	return sections
}

// updateField mutates configuration settings based on string form payloads.
func updateField(cfg *Config, sectionName, fieldTomlName, strVal string) {
	val := reflect.ValueOf(cfg).Elem()
	typ := val.Type()
	for i := 0; i < val.NumField(); i++ {
		field := typ.Field(i)
		secName := field.Tag.Get("toml")
		if secName == "" {
			secName = strings.ToLower(field.Name)
		}
		if strings.EqualFold(secName, sectionName) {
			secVal := val.Field(i)
			secTyp := secVal.Type()
			for j := 0; j < secVal.NumField(); j++ {
				subField := secTyp.Field(j)
				subKey := subField.Tag.Get("toml")
				if subKey == "" {
					subKey = strings.ToLower(subField.Name)
				}
				if strings.EqualFold(subKey, fieldTomlName) {
					subVal := secVal.Field(j)
					switch subVal.Kind() {
					case reflect.String:
						subVal.SetString(strVal)
					case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64:
						if intVal, err := strconv.Atoi(strVal); err == nil {
							subVal.SetInt(int64(intVal))
						}
					case reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
						if uintVal, err := strconv.ParseUint(strVal, 10, 64); err == nil {
							subVal.SetUint(uintVal)
						}
					case reflect.Bool:
						subVal.SetBool(strVal == "true" || strVal == "on")
					case reflect.Slice:
						parts := strings.Split(strVal, ",")
						var cleanParts []string
						for _, p := range parts {
							p = strings.TrimSpace(p)
							if p != "" {
								cleanParts = append(cleanParts, p)
							}
						}
						subVal.Set(reflect.ValueOf(cleanParts))
					}
					return
				}
			}
		}
	}
}

// openBrowser initiates a platform-appropriate command to spawn the target URL.
func openBrowser(url string) error {
	var cmd string
	var args []string

	switch runtime.GOOS {
	case "windows":
		cmd = "rundll32"
		args = []string{"url.dll,FileProtocolHandler", url}
	case "darwin":
		cmd = "open"
		args = []string{url}
	default:
		cmd = "xdg-open"
		args = []string{url}
	}
	return exec.Command(cmd, args...).Start()
}

// backupConfigFile creates a .bak copy of the configuration file if it exists.
func backupConfigFile(path string) error {
	if _, err := os.Stat(path); os.IsNotExist(err) {
		return nil
	}
	src, err := os.Open(path)
	if err != nil {
		return err
	}
	defer src.Close()

	backupPath := path + ".bak"
	dst, err := os.Create(backupPath)
	if err != nil {
		return err
	}
	defer dst.Close()

	_, err = io.Copy(dst, src)
	return err
}

// printHelp displays the standard command-line help usage menu.
func printHelp() {
	fmt.Println("\033[1mG3pix ❖ AxonASP Configuration Management Usage:\n\033[0m")
	fmt.Println(`    axonadmin
      Starts the interactive web interface.
    axonadmin -edit  <path>    
      Specify TOML configuration file to edit (mimics loader path resolution 
      if omitted)
    axonadmin -create <path>  
      Generate a new default configuration file at specified path
		axonadmin -create-fpm <path>
			Generate a new default FPM pool configuration file (.conf)
    axonadmin -noui           
			Run in headless mode (must be used with -create or -create-fpm)
    axonadmin -h, --help
      Shows this help message.

 ABOUT:
  G3pix ❖ AxonASP
  is a high-performance, cross-platform Classic ASP engine,
  with support to VBScript and JavaScript for Web, FastCGI, and CLI, 
  bridging legacy compatibility with modern APIs.
  
  Copyright (C) 2026 G3pix Ltda. All rights reserved.
  Website: https://g3pix.com.br/axonasp
  
  License: MPL 2.0
  `)
	fmt.Println("\033[0m")
}

// sendJSONError writes a HTTP JSON error payload.
func sendJSONError(w http.ResponseWriter, msg string) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(http.StatusBadRequest)
	_ = json.NewEncoder(w).Encode(map[string]string{
		"status":  "error",
		"message": msg,
	})
}

// sendJSONOK writes a HTTP JSON success payload.
func sendJSONOK(w http.ResponseWriter) {
	w.Header().Set("Content-Type", "application/json")
	_ = json.NewEncoder(w).Encode(map[string]string{
		"status": "ok",
	})
}

func main() {
	var editPath string
	var createPath string
	var createFPMPath string
	var noUI bool
	var aboutFlag bool

	pflag.StringVar(&editPath, "edit", "", "Path to an existing AxonASP TOML configuration file to open for editing in the web interface.")
	pflag.StringVar(&createPath, "create", "", "Create a new default AxonASP TOML configuration file at the provided path and exit.")
	pflag.StringVar(&createFPMPath, "create-fpm", "", "Create a new default AxonASP FPM pool .conf file at the provided path and exit.")
	pflag.BoolVar(&noUI, "noui", false, "Run in headless mode for create operations without opening the browser interface.")
	pflag.BoolVarP(&aboutFlag, "about", "a", false, "Print AxonASP product and licensing information, then exit.")

	pflag.Usage = func() {
		fmt.Printf("G3pix ❖ AxonASP Configuration Manager %s\n", Version)
		fmt.Println("Options available: ")
		pflag.PrintDefaults()
		fmt.Print("\nFor more information, visit: https://g3pix.com.br/axonasp/manual/\n")
	}

	pflag.Parse()

	if aboutFlag {
		fmt.Print(axonconfig.AboutG3pixAxonASP())
		os.Exit(0)
	}

	// Handle extra non-flag arguments as unrecognized.
	if pflag.NArg() > 0 {
		pflag.Usage()
		os.Exit(0)
	}

	// Headless validation
	if (noUI) && createPath == "" && createFPMPath == "" && editPath == "" {
		printHelp()
		os.Exit(0)
	}

	if createPath != "" && createFPMPath != "" {
		fmt.Println("Error: use either -create or -create-fpm, not both at the same time")
		os.Exit(1)
	}

	// If create path is set (with or without headless), generate configuration
	if createPath != "" || (noUI && createPath == "" && createFPMPath == "") {
		target := createPath
		if target == "" {
			target = filepath.Join("config", "axonasp.toml")
		}
		err := createNewConfig(target)
		if err != nil {
			fmt.Printf("Error creating config file: %v\n", err)
			os.Exit(1)
		}
		fmt.Printf("Successfully created configuration file: %s\n", target)
		os.Exit(0)
	}

	if createFPMPath != "" {
		target := createFPMPath
		if abs, err := filepath.Abs(target); err == nil {
			target = abs
		}
		err := createNewFPMConfig(target)
		if err != nil {
			fmt.Printf("Error creating FPM config file: %v\n", err)
			os.Exit(1)
		}
		fmt.Printf("Successfully created FPM configuration file: %s\n", target)
		os.Exit(0)
	}

	// Resolve the target TOML file for editing
	resolvedPath := editPath
	if resolvedPath == "" {
		resolvedPath = resolveConfigPath()
	} else {
		if abs, err := filepath.Abs(resolvedPath); err == nil {
			resolvedPath = abs
		}
	}

	// Ensure the parent directory exists
	if err := os.MkdirAll(filepath.Dir(resolvedPath), 0755); err != nil {
		log.Fatalf("Error ensuring parent directory: %v", err)
	}

	fpmResolvedDir := resolveFPMConfigDir()
	if err := os.MkdirAll(fpmResolvedDir, 0755); err != nil {
		log.Fatalf("Error ensuring FPM config directory: %v", err)
	}

	// Launch local web server
	mux := http.NewServeMux()

	// API Endpoint: Save Section
	mux.HandleFunc("/api/save", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		if err := r.ParseForm(); err != nil {
			sendJSONError(w, "Failed to parse form: "+err.Error())
			return
		}

		var currentCfg Config
		if data, err := os.ReadFile(resolvedPath); err == nil {
			_ = toml.Unmarshal(data, &currentCfg)
		} else {
			currentCfg = NewDefaultConfig()
		}

		for k, vs := range r.Form {
			if strings.HasPrefix(k, "field_") && len(vs) > 0 {
				parts := strings.SplitN(k, "_", 3)
				if len(parts) == 3 {
					section := parts[1]
					key := parts[2]
					strVal := vs[0]
					updateField(&currentCfg, section, key, strVal)
				}
			}
		}

		// Backup existing configuration before saving
		if err := backupConfigFile(resolvedPath); err != nil {
			sendJSONError(w, "Failed to backup config: "+err.Error())
			return
		}

		data, err := toml.Marshal(currentCfg)
		if err != nil {
			sendJSONError(w, "Failed to marshal TOML: "+err.Error())
			return
		}

		if err := os.WriteFile(resolvedPath, data, 0644); err != nil {
			sendJSONError(w, "Failed to write file: "+err.Error())
			return
		}

		sendJSONOK(w)
	})

	// API Endpoint: Recreate configuration file
	mux.HandleFunc("/api/recreate", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}

		// Backup existing configuration before recreating
		if err := backupConfigFile(resolvedPath); err != nil {
			sendJSONError(w, "Failed to backup config: "+err.Error())
			return
		}

		defaultCfg := NewDefaultConfig()
		data, err := toml.Marshal(defaultCfg)
		if err != nil {
			sendJSONError(w, "Failed to marshal TOML: "+err.Error())
			return
		}

		if err := os.WriteFile(resolvedPath, data, 0644); err != nil {
			sendJSONError(w, "Failed to write file: "+err.Error())
			return
		}

		sendJSONOK(w)
	})

	// API Endpoint: Create a new global axonasp.toml file at a requested path.
	mux.HandleFunc("/api/create-global", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		if err := r.ParseForm(); err != nil {
			sendJSONError(w, "Failed to parse form: "+err.Error())
			return
		}

		targetPath := strings.TrimSpace(r.FormValue("target_path"))
		if targetPath == "" {
			targetPath = resolvedPath
		}
		if abs, err := filepath.Abs(targetPath); err == nil {
			targetPath = abs
		}

		if err := backupConfigFile(targetPath); err != nil {
			sendJSONError(w, "Failed to backup target file: "+err.Error())
			return
		}
		if err := createNewConfig(targetPath); err != nil {
			sendJSONError(w, "Failed to create axonasp.toml: "+err.Error())
			return
		}

		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]string{
			"status": "ok",
			"path":   targetPath,
		})
	})

	// API Endpoint: Save one FPM pool configuration file.
	mux.HandleFunc("/api/fpm/save", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		if err := r.ParseForm(); err != nil {
			sendJSONError(w, "Failed to parse form: "+err.Error())
			return
		}

		poolName, err := normalizeFPMFileName(r.FormValue("pool_name"))
		if err != nil {
			sendJSONError(w, "Invalid pool name: "+err.Error())
			return
		}

		poolPath := filepath.Join(fpmResolvedDir, poolName)
		var currentPool FPMPoolConfig
		if data, readErr := os.ReadFile(poolPath); readErr == nil {
			if unmarshalErr := toml.Unmarshal(data, &currentPool); unmarshalErr != nil {
				sendJSONError(w, "Failed to parse existing pool file: "+unmarshalErr.Error())
				return
			}
		} else {
			currentPool = NewDefaultFPMPoolConfig()
		}

		for key, values := range r.Form {
			if !strings.HasPrefix(key, "field_fpm_") || len(values) == 0 {
				continue
			}
			fieldName := strings.TrimPrefix(key, "field_fpm_")
			updateFPMField(&currentPool, fieldName, values[0])
		}

		if err := backupConfigFile(poolPath); err != nil {
			sendJSONError(w, "Failed to backup pool file: "+err.Error())
			return
		}

		data, err := toml.Marshal(currentPool)
		if err != nil {
			sendJSONError(w, "Failed to marshal pool TOML: "+err.Error())
			return
		}
		if err := os.WriteFile(poolPath, data, 0644); err != nil {
			sendJSONError(w, "Failed to write pool file: "+err.Error())
			return
		}

		sendJSONOK(w)
	})

	// API Endpoint: Create a brand-new FPM pool configuration file.
	mux.HandleFunc("/api/fpm/create", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		if err := r.ParseForm(); err != nil {
			sendJSONError(w, "Failed to parse form: "+err.Error())
			return
		}

		poolName, err := normalizeFPMFileName(r.FormValue("filename"))
		if err != nil {
			sendJSONError(w, "Invalid filename: "+err.Error())
			return
		}
		poolPath := filepath.Join(fpmResolvedDir, poolName)

		if _, statErr := os.Stat(poolPath); statErr == nil {
			sendJSONError(w, "Pool file already exists: "+poolName)
			return
		}
		if err := createNewFPMConfig(poolPath); err != nil {
			sendJSONError(w, "Failed to create pool file: "+err.Error())
			return
		}

		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]string{
			"status": "ok",
			"pool":   poolName,
		})
	})

	// API Endpoint: Live host/process telemetry for dashboard refresh.
	mux.HandleFunc("/api/telemetry", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(collectHomeTelemetry(resolvedPath))
	})

	// API Endpoint: Return supported flags/help text for a managed process executable.
	mux.HandleFunc("/api/process/flags", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodGet {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}

		processName := strings.TrimSpace(r.URL.Query().Get("process_name"))
		allowedProcesses := map[string]bool{
			"axonasp-http":    true,
			"axonasp-fastcgi": true,
			"axonasp-fpm":     true,
			"axonasp-service": true,
			"axonasp-cli":     true,
		}
		if !allowedProcesses[processName] {
			sendJSONError(w, "Unsupported process name")
			return
		}

		helpText, err := getManagedProcessHelp(processName)
		if err != nil {
			sendJSONError(w, "Unable to read supported flags: "+err.Error())
			return
		}

		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]string{
			"status":    "ok",
			"process":   processName,
			"help_text": helpText,
		})
	})

	// API Endpoint: Start/stop managed AxonASP processes from the dashboard.
	mux.HandleFunc("/api/process/action", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "Method not allowed", http.StatusMethodNotAllowed)
			return
		}
		if err := r.ParseForm(); err != nil {
			sendJSONError(w, "Failed to parse form: "+err.Error())
			return
		}

		processName := strings.TrimSpace(r.FormValue("process_name"))
		action := strings.TrimSpace(strings.ToLower(r.FormValue("action")))
		rawArgs := strings.TrimSpace(r.FormValue("args"))

		allowedProcesses := map[string]bool{
			"axonasp-http":    true,
			"axonasp-fastcgi": true,
			"axonasp-fpm":     true,
			"axonasp-service": true,
			"axonasp-cli":     true,
		}
		if !allowedProcesses[processName] {
			sendJSONError(w, "Unsupported process name")
			return
		}

		if processName == "axonasp-fpm" && runtime.GOOS == "windows" {
			sendJSONError(w, "axonasp-fpm is not supported on Windows")
			return
		}

		var message string
		switch action {
		case "start":
			if err := startManagedProcess(processName, rawArgs); err != nil {
				sendJSONError(w, err.Error())
				return
			}
			if rawArgs != "" {
				message = processName + " started successfully with parameters"
			} else {
				message = processName + " started successfully"
			}
		case "stop":
			stopped, err := stopManagedProcess(processName)
			if err != nil {
				sendJSONError(w, err.Error())
				return
			}
			message = fmt.Sprintf("%s stopped (%d instance(s))", processName, stopped)
		default:
			sendJSONError(w, "Unsupported action")
			return
		}

		w.Header().Set("Content-Type", "application/json")
		_ = json.NewEncoder(w).Encode(map[string]string{
			"status":  "ok",
			"message": message,
		})
	})

	// Root Handler: serves the static files and templated dashboard UI
	mux.HandleFunc("/", func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path != "/" {
			filePath := strings.TrimPrefix(r.URL.Path, "/")
			filePath = path.Clean(filePath)
			if filePath == "." || strings.HasPrefix(filePath, "../") || strings.Contains(filePath, "/../") {
				http.NotFound(w, r)
				return
			}
			data, err := wwwInterfaceFS.ReadFile(path.Join("www-interface", filePath))
			if err != nil {
				http.NotFound(w, r)
				return
			}
			ext := strings.ToLower(path.Ext(filePath))
			contentType := mime.TypeByExtension(ext)
			if ext == ".svg" {
				// Force a stable MIME type for browsers even when host MIME tables are incomplete.
				contentType = "image/svg+xml"
			}
			if contentType == "" {
				contentType = http.DetectContentType(data)
			}
			w.Header().Set("Content-Type", contentType)
			w.Write(data)
			return
		}

		sectionName := strings.TrimSpace(r.URL.Query().Get("section"))
		activeView := strings.TrimSpace(r.URL.Query().Get("view"))
		activePool := strings.TrimSpace(r.URL.Query().Get("pool"))

		if activeView == "" {
			switch {
			case sectionName != "":
				activeView = "config"
			case activePool != "":
				activeView = "fpm"
			default:
				activeView = "home"
			}
		}

		if sectionName == "" {
			sectionName = "global"
		}

		var currentCfg Config
		if data, err := os.ReadFile(resolvedPath); err == nil {
			_ = toml.Unmarshal(data, &currentCfg)
		} else {
			currentCfg = NewDefaultConfig()
		}

		defaultCfg := NewDefaultConfig()
		sections := getSections(currentCfg, defaultCfg)
		telemetry := collectHomeTelemetry(resolvedPath)

		poolFiles, poolListErr := listFPMConfigFiles(fpmResolvedDir)
		if poolListErr != nil {
			telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "failed to list FPM pools: "+poolListErr.Error())
			poolFiles = []string{}
		}

		if activePool == "" && len(poolFiles) > 0 {
			activePool = poolFiles[0]
		}

		defaultPool := NewDefaultFPMPoolConfig()
		currentPool := defaultPool
		if activePool != "" {
			normalizedPoolName, nameErr := normalizeFPMFileName(activePool)
			if nameErr == nil {
				activePool = normalizedPoolName
				poolPath := filepath.Join(fpmResolvedDir, activePool)
				if data, readErr := os.ReadFile(poolPath); readErr == nil {
					if err := toml.Unmarshal(data, &currentPool); err != nil {
						telemetry.CollectionWarnings = append(telemetry.CollectionWarnings, "failed to parse active pool file: "+err.Error())
					}
				}
			}
		}
		fpmFields := getFPMFields(currentPool, defaultPool)

		var activeSection Section
		found := false
		for _, sec := range sections {
			if strings.EqualFold(sec.Name, sectionName) {
				activeSection = sec
				found = true
				break
			}
		}
		if !found && len(sections) > 0 {
			activeSection = sections[0]
		}

		tmplData, err := wwwInterfaceFS.ReadFile("www-interface/index.html")
		if err != nil {
			http.Error(w, "Template error: "+err.Error(), http.StatusInternalServerError)
			return
		}
		tmpl, err := template.New("index.html").Funcs(template.FuncMap{
			"FormatVal": func(val any) string {
				if slice, ok := val.([]string); ok {
					return strings.Join(slice, ", ")
				}
				return fmt.Sprintf("%v", val)
			},
			"FormatSlice": func(val any) string {
				if slice, ok := val.([]string); ok {
					return strings.Join(slice, ", ")
				}
				return ""
			},
			"FormatBytes": formatBytes,
			"FormatPercent": func(v float64) string {
				return fmt.Sprintf("%.2f%%", v)
			},
		}).Parse(string(tmplData))
		if err != nil {
			http.Error(w, "Template error: "+err.Error(), http.StatusInternalServerError)
			return
		}

		pageData := PageData{
			Sections:      sections,
			ActiveSection: activeSection,
			ResolvedPath:  resolvedPath,
			ActiveView:    activeView,
			IsWindows:     runtime.GOOS == "windows",
			IsGUIEnv:      isGUIEnvironment(),
			FPMDir:        fpmResolvedDir,
			FPMPools:      poolFiles,
			ActivePool:    activePool,
			FPMFields:     fpmFields,
			HomeTelemetry: telemetry,
		}

		if err := tmpl.Execute(w, pageData); err != nil {
			log.Printf("Template execution error: %v", err)
		}
	})

	// Format console startup title and port output to match server/main.go
	fmt.Printf("\033[H\033[2J\033[1mG3pix ❖ AxonASP Server %s \033[0m\n", Version)
	fmt.Printf("Configuration manager started on: %s\n", "8088")
	fmt.Print("\033]0;G3pix ❖ AxonASP Configuration Manager\007\033]11;#0d7423\007\033[1;37m")

	// Trigger async cross-platform browser launch
	go func() {
		time.Sleep(500 * time.Millisecond)
		_ = openBrowser("http://localhost:8088")
	}()

	listener, err := net.Listen("tcp", ":8088")
	if err != nil {
		log.Fatalf("Error listening on port 8088: %v", err)
	}

	if err := http.Serve(listener, mux); err != nil && err != http.ErrServerClosed {
		log.Fatalf("Error serving AxonASP Configuration Manager server: %v", err)
	}
}
