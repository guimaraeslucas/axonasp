# AxonASP WebAssembly (`wasm`)

The `wasm` package provides the WebAssembly (WASM) build target and browser runtime bindings for AxonASP.

## Features

- **In-Browser ASP Engine**: Compiles AxonASP to WASM (`js/wasm` architecture), enabling execution of Classic ASP scripts directly inside modern web browsers.
- **JavaScript DOM Interop**: Exposes JavaScript bridge APIs for passing ASP scripts, receiving output buffers, and interacting with browser-based sandboxes and playgrounds.
- **Zero Backend Required**: Runs client-side for educational playgrounds, documentation viewers, and offline execution scenarios.
