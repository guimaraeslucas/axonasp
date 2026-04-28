# Welcome to AxonASP 2.0: The Ultimate Classic ASP Engine for the Modern Web

Supercharge your legacy code. Build blazing-fast modern APIs. Experience Classic ASP like never before.

Welcome to **AxonASP 2.0**, the definitive, high-performance runtime for executing Classic ASP with VBScript and JScript in pure Go. This is the engine your applications deserve—a complete reinvention of what Classic ASP can be in the modern era.

---

## What Is AxonASP?

AxonASP is a high-performance execution engine that processes Classic ASP pages, compiles VBScript into optimized bytecode, and executes code through an advanced stack-based virtual machine. The engine also supports JScript execution for ASP pages using the JScript language directive. The runtime provides compatibility with classic ASP behavior while introducing powerful modern innovations that make ASP development faster, safer, and more scalable than ever before.

Unlike traditional ASP interpreters, AxonASP runs on **any operating system**—Windows, Linux, macOS, and beyond—using the same codebase. Deploy anywhere Go is supported. No more constraint to Windows and IIS.

---

## Why Choose AxonASP? The Performance Revolution

We didn't just update Classic ASP. We completely reimagined it for the modern web:

### Lightning-Fast Compilation and Execution
- **Zero AST, Pure Bytecode:** Single-pass compilation emits bytecode directly to a highly optimized stack-based Virtual Machine. By eliminating the Abstract Syntax Tree (AST), AxonASP executes scripts with virtually zero-allocation overhead. It's insanely fast and memory efficient.
- **Advanced Caching Architecture:** IIS-style VM pooling combined with aggressive script caching and dynamic execution caching. Your compiled code is cached at every level—script files, Eval expressions, Execute statements, and beyond.
- **Minimal Memory Footprint:** Code is engineered to minimize garbage collection triggers. Scripts complete faster and consume less memory than traditional interpreters.

### Total Compatibility Meets Innovation
- **Classic ASP Standards Compliance:** Your legacy code drops right in. Full support for VBScript semantics, ADOB, MSXML, FileSystemObject, Dictionary, WScript.Shell, and all intrinsic ASP objects.
- **60+ Custom Axon Functions:** High-performance native functions for arrays, strings, system operations, and advanced utilities—written in Go, executing at near-native speed.
- **Enterprise Native Libraries:** G3JSON, G3DB, G3HTTP, G3Mail, G3Image, G3Crypto, G3ZIP, G3TAR, G3ZSTD, G3PDF, G3Template, and many more—zero overhead, maximum power.

### Multiple Execution Modes
- **Web Server Mode:** Run the built-in HTTP server (`axonasp-http`) for development and production.
- **FastCGI Mode:** Integrate with Nginx, Apache, or any FastCGI-compatible web server.
- **Command-Line Mode:** Execute ASP scripts directly from your terminal (`axonasp-cli`). Schedule tasks, automation, system administration—all in ASP.
- **Interactive CLI with TUI:** Explore and test code in real-time with a text-based user interface.

### AI-Ready Architecture
- **Built-in MCP Server:** AxonASP includes a Model Context Protocol (MCP) server. AI agents can connect directly, understand your environment, and autonomously author complete ASP pages using all available native functions.

### Testing and Quality Assurance
- **Automated Test Suite:** The `axonasp-testsuite` executable enables test-driven development natively in ASP. Write assertions, run test discovery automatically, and report CI-friendly results.
- **No More Regex Hacks:** Integrated testing framework—no more manual test pages and prayer-based debugging.

---

## What's New in Version 2.0?

- **Unified Configuration:** Centralized `axonasp.toml` configuration with `.env` support via Viper. Single source of truth for all settings.
- **Modern Architecture Examples:** Complete, production-ready examples for REST, RESTful, MVC, and MVVM—all written in pure ASP.
- **Comprehensive Local Documentation:** The complete manual is built right into the repository. No need to hunt through old forums or outdated blogs. Everything you need is in `./www/manual/md/`.
- **Intelligent Port Defaults:** Updated default proxy port (8801) to avoid firewall conflicts and system port contention out of the box.
- **Native Docker Support:** Containerization with included `Dockerfile` and `docker-compose.yml`. Deploy in seconds.
- **Database Migration Tool:** Built-in converter to migrate legacy Access databases to modern formats (SQLite, MySQL, PostgreSQL, MSSQL).

---

## Understanding This Manual

This manual is organized to help you understand AxonASP from the ground up:

### Section: Runtime
Covers deployment modes, internal architecture, script caching, locale support, global.asa lifecycle, web.config directives, FastCGI setup, reverse proxy configuration, running as a Linux service, installing from pre-built Linux packages (DEB, RPM, APK), the service wrapper, system and error pages, and MCP Server integration with VS Code.

### Section: Configuration
Documents the `axonasp.toml` configuration file: all available keys, default values, and explanations for each setting that controls the engine, server, sessions, caching, and logging.

### Section: ASP Foundations
A complete reference for Classic ASP fundamentals with VBScript and JScript guidance: the `#include` directive, variables, procedures, conditionals, looping, syntax overview, quick reference, forms, cookies, intrinsic ASP objects (Request, Response, Server, Session, Application), ASPError, Scripting.Dictionary, JScript page directives, and global `console` diagnostics.

### Section: Libraries
Full API reference for every built-in native library and compatibility object:

**Native AxonASP Libraries**
G3AXON Functions, G3CRYPTO, G3JSON, G3DB, G3HTTP, G3MAIL, G3IMAGE, G3FILES, G3TESTSUITE, G3TEMPLATE, G3ZIP, G3ZLIB, G3TAR, G3ZSTD, G3FC, G3MD, G3PDF, G3FILEUPLOADER.

**Compatibility Objects**
WScript.Shell, ADOX.Catalog, MSWC Components (AdRotator, NextLink, ContentRotator, Counters, Tools, PageCounter, BrowserType, PermissionChecker, MyInfo), ADODB Family (Connection, Recordset, Command, Fields, Parameters, Errors), ADODB.Stream, Scripting.Dictionary, Scripting.FileSystemObject, VBScript.RegExp, MSXML2 Family (ServerXMLHTTP, DOMDocument, XMLNodeList, XMLElement).

### Section: Examples
Working code demonstrating architectural patterns, library usage, and best practices for Classic ASP applications built on AxonASP.

### Section: Tools
Documentation for the built-in database migration tool, which converts legacy Access databases to modern formats including SQLite, MySQL, PostgreSQL, and MSSQL.

### Section: Authoring
Guidelines for writing manual pages and for programming Classic ASP with LLMs. Read these before contributing documentation.

---

## Key Capabilities at a Glance

**Server-Side Power:**
- Classic ASP page execution with high compatibility for legacy behavior
- VBScript with standard functions and operators
- JScript support for ASP pages using `<%@ Language=JScript %>`
- Full intrinsic object support (Request, Response, Server, Session, Application)
- Advanced error handling with line-level debugging

**Native Libraries (Zero-Overhead):**
G3JSON, G3DB, G3HTTP, G3Mail, G3Image, G3Crypto, G3Zip, G3TAR, G3ZSTD, G3ZLIB, G3Template, G3PDF, G3MD, G3FC, G3FileUploader, and G3TestSuite.

**Compatibility Objects:**
ADODB (with real database backends), MSXML, Scripting.FileSystemObject, Scripting.Dictionary, WScript.Shell, VBScript.RegExp, and MSWC components.

**Production Ready:**
Reverse proxy mode, FastCGI integration, command-line execution, automated testing, MCP for AI agents, comprehensive logging, and performance monitoring.

---

## Getting Started

To begin working with AxonASP, start by reading about the runtime architecture and deployment options. Then select the deployment model that fits your infrastructure, configure your environment, and deploy your applications with confidence.

For detailed API reference on any library or object, consult the menu on the left or use `./www/manual/menu.md` for a complete navigational structure.

---

## A Word About Version 2.0

Version 1.0 of AxonASP is completely deprecated and incompatible with Version 2.0. We made deliberate architectural decisions to deliver superior performance, maintainability, and feature parity with modern requirements. The investment in rewriting the engine from the ground up pays dividends in speed, reliability, and joy of development.

If you're upgrading from Version 1.0, your AxonASP code and configuration *will require migration*. It's worth the effort—the improvements are transformational.

---

## Questions? Ready to Deploy?

This manual contains everything you need to understand, configure, and deploy AxonASP. Each section builds on prior knowledge, so start at the beginning if you're new, or jump directly to specific topics you need to solve.

Your Classic ASP applications—and your team—are about to experience what the platform can truly do.

