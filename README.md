# ❖ AxonASP 2.3: The Ultimate Classic ASP Engine for the Modern Web
<p align="center">
  <img src="www/axonasp.svg" alt="G3Pix AxonASP Logo" width="400"/>
</p>

<p align="center">
  <b>Run Classic ASP on Linux, Windows, and macOS natively with extreme performance.</b>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/version-2.3-blue.svg" alt="Version 2.3"/>
  <img src="https://img.shields.io/badge/Go-1.26+-00ADD8.svg" alt="Go Version"/>
  <img src="https://img.shields.io/badge/platform-Windows%20%7C%20Linux%20%7C%20macOS-lightgrey.svg" alt="Platform"/>
  <img src="https://img.shields.io/badge/license-MPL-green.svg" alt="License"/>
</p>

Welcome to **AxonASP 2.3**, the definitive high-performance runtime for executing Microsoft Classic ASP, VBScript, and JavaScript in GoLang. If you are looking for how to run Classic ASP on Linux or modernize legacy applications within a Docker container, AxonASP provides a robust, zero-allocation VM architecture tailored for modern infrastructure.

## 🚀 Why AxonASP?

AxonASP is built from the ground up for speed, low memory usage, and containerized deployments.

* **Run Classic ASP on Linux & Cross-Platform:** Say goodbye to IIS dependencies. AxonASP runs natively on Windows, Linux, and macOS, either as a standalone web server, via FastCGI, or directly from the CLI.
* **Modern JavaScript (ES6+) Implementation:** We support fully compliant ECMAScript 6+ alongside VBScript. Unlike Node.js, AxonASP executes JavaScript directly through our AST-based engine without forcing a Promise-based paradigm for everything.
  * **Memory Efficiency:** A pure JavaScript API running on AxonASP uses only **18MB in idle mode** (compared to 30MB for Node.js). During execution, parsing and serving JSON takes just **30MB** versus 50MB in Node.js.
  * **Partial Node.js Module Compatibility:** Easily organize and import JavaScript code using CommonJS/ES6 module syntax directly in your ASP environment.
* **High Performance & Active Caching:** AxonASP is heavily optimized to run ASP pages and components swiftly with minimal CPU and memory footprints. With our active caching architecture, **2000 requests per minute** on simple pages consume only about **100MB of memory**.
* **Zero AST for VBScript:** Our VBScript compiler is single-pass, emitting bytecode directly to a stack-based VM to achieve extreme processing speeds.
* **AI-Ready:** Includes a built-in Model Context Protocol (MCP) server, allowing AI agents to connect directly to the runtime, analyze your environment, and autonomously author ASP pages.

## ⚡ Built-in G3 Libraries

AxonASP extends standard ASP with native Go libraries, giving you enterprise power with zero overhead:
- **G3JSON, G3DB, G3HTTP:** Effortlessly handle JSON, connect to databases (SQLite, MySQL, PostgreSQL, MS SQL, Oracle), and fetch external APIs.
- **G3CRYPTO, G3MAIL, G3IMAGE:** Generate secure hashes, send SMTP emails, and process images on the fly.
- **G3ZIP, G3TAR, G3ZLIB, G3ZSTD:** Full compression suite built right in.
- **And much more...** Over 60 custom Axon functions are available to extend ASP's capabilities.

For a full list of libraries and documentation on how to use them, please read our extensive manual at [`./www/manual/`](www/manual/) and explore our website at [https://g3pix.com.br/axonasp](https://g3pix.com.br/axonasp), where you will find pages running in AxonASP 2.3 and also our WASM Playground.

## 🎨 AxonLive with Visual Application Builder

**AxonLive** is a high-performance Reactive Component Framework built directly into the AxonASP Virtual Machine. It empowers developers to create dynamic, stateful, and highly responsive web applications using Classic ASP (VBScript or Server-Side JavaScript) without requiring full page reloads.

### Advantages of G3AxonLive

* **Zero Page Reloads:** All UI interactions (button clicks, form submissions, timers) are sent to the server asynchronously. The server responds with targeted JSON patches, swapping only the modified DOM elements.
* **Strict Backend Control:** All business logic, validation, and state mutation happen exclusively on the server. The client browser merely acts as a dumb terminal rendering the HTML patches, significantly reducing the attack surface.
* **Authenticated Session Binding:** The `/g3al` endpoint binds every async event to the authenticated `ASPSESSIONID` cookie. Client-provided session identifiers are not used as an authority for page routing.
* **Zero Additional Wrappers:** AxonLive is implemented directly inside the `axonvm` engine as a native procedural controller (`G3AXONLIVE`), providing bare-metal performance and zero garbage collection overhead.
* **Granular DOM Manipulation:** Push targeted instructions to modify styles, attributes, classes, or trigger external redirects natively from ASP.
* **WASM + AxonLive:** Run performance-critical VBScript directly in the browser using AxonASP WebAssembly, opening incredible possibilities for offline-capable web applications.

You can use the **AxonLive Builder** available in the `www/axonlive/builder/` directory to seamlessly create and manage reactive components.

## 🚀 Quick Deployment & Execution

AxonASP is ready for modern CI/CD pipelines and containerization.

### Prerequisites
* GoLang 1.26+ (if building from source)

### Building the Engine
Use the provided build scripts to compile AxonASP for your target architecture:

**Linux / macOS:**
```bash
./build.sh --platform "linux" --architecture "amd64"
```

**Windows:**
```powershell
.\build.ps1 -Platform "windows" -Architecture "amd64"
```

You can optionally disable specific libraries (e.g., `lib_g3crypto_disabled`) to create leaner binaries. See the manual for details.

### Deployment Architecture
Deploy AxonASP via its built-in HTTP server (Reverse Proxy Mode) or via FastCGI (`axonasp-fastcgi`). It integrates flawlessly with Nginx and Apache. 

See our full documentation in `www/manual/md/` for examples and complete API details.

## License 📝

This project is licensed under the MPL 2.0 License - see the [LICENSE.txt](LICENSE.txt) file for details.

## Contributing to AxonASP 🤝

While not mandatory, we warmly encourage you to contribute to the AxonASP project. If you develop new components or libraries, please consider sharing your implementation with the community; your contributions help maintain, improve, and expand our ecosystem.

If you find bugs, security vulnerabilities, or have suggestions for new features, **please report them** via GitHub Issues. We welcome your feedback and will do our best to address your concerns promptly.

As an open-source project, AxonASP relies on community support to remain active. Maintenance and infrastructure costs are currently funded personally by the lead developer. Your support—whether through code contributions, bug reports, security patches, or financial donations—is vital to our continued growth. 

* Please see [CONTRIBUTING.md](CONTRIBUTING.md) for details on how to contribute.
* Please see [SECURITY.md](SECURITY.md) for information on reporting security vulnerabilities.

Thanks goes to these wonderful people:

<!-- ALL-CONTRIBUTORS-LIST:START - Do not remove or modify this section -->
<!-- prettier-ignore-start -->
<!-- markdownlint-disable -->
<table>
  <tbody>
    <tr>
      <td align="center" valign="top" width="14.28%"><a href="https://github.com/ikadmm"><img src="https://avatars.githubusercontent.com/u/29986197?v=4?s=100" width="100px;" alt="ikadmm"/><br /><sub><b>ikadmm</b></sub></a><br /><a href="https://github.com/guimaraeslucas/axonasp/issues?q=author%3Aikadmm" title="Bug reports">🐛</a></td>
      <td align="center" valign="top" width="14.28%"><a href="https://github.com/dmitriy-bty"><img src="https://avatars.githubusercontent.com/u/285232216?v=4?s=100" width="100px;" alt="dmitriy-bty"/><br /><sub><b>dmitriy-bty</b></sub></a><br /><a href="https://github.com/guimaraeslucas/axonasp/commits?author=dmitriy-bty" title="Code">💻</a></td>
    </tr>
  </tbody>
</table>

<!-- markdownlint-restore -->
<!-- prettier-ignore-end -->

<!-- ALL-CONTRIBUTORS-LIST:END -->

---

<p align="center">
  <strong>Built with ❤️ by G3Pix</strong>
  <br>
  Making Classic ASP modern, fast, and cross-platform
</p>

---
#### Legal Disclaimer ⚖️

AxonASP is an independent software project developed by G3pix and is **not affiliated with, endorsed by, or sponsored by Microsoft Corporation**. Trademarks such as "Microsoft," "Active Server Pages," "ASP," "VBScript," "JScript" are registered trademarks of Microsoft Corporation. The names and aliases of other third-party libraries, software, products, and services mentioned in this project are the property of their respective owners. The use of these names is solely for descriptive purposes/compatibility and does not imply any endorsement or affiliation with the trademark holders or use of their original source code.

The names and aliases of other third-party libraries, software, products, and services mentioned in this project are the property of their respective owners. The use of these names is solely for descriptive purposes/compatibility and does not imply any endorsement or affiliation with the trademark holders or use of their original source code.