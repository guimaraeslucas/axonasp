# AxonVM (`axonvm`)

`axonvm` is the core Virtual Machine and single-pass compiler engine for Classic ASP in Go.

## Architecture & Capabilities

- **Single-Pass Bytecode Compiler**: Directly translates tokens into compact bytecode without intermediate AST overhead for high-speed compilation.
- **Stack-Based Execution Engine**: High-performance, zero-allocation oriented execution loop with fixed-size call frames and operand stacks.
- **ASP Intrinsic Objects**: Full implementation of Classic ASP intrinsics (`Response`, `Request`, `Server`, `Session`, `Application`, and `ASPError`) under `axonvm/asp/`.
- **Native Component Libraries**: Built-in implementations for `Server.CreateObject(...)` components (`lib_*.go`) such as JSON, Cryptography, File Management, Databases, ZIP/Tar compression, and Image manipulation without relying on Go reflection.
- **Deterministic Built-ins**: High-speed lookup and execution of native VBScript and Axon helper functions.
