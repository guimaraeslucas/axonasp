# Script Caching

## Overview

AxonASP uses a multi-tier bytecode caching system to eliminate redundant compilation of ASP scripts and dynamic code expressions. Once a script is compiled the first time, subsequent requests reuse the cached bytecode directly, bypassing the lexer and compiler entirely. Caching operates across four independent subsystems depending on the source of the compiled code.

## Cache Tiers

### 1. Script Cache (axonvm/script_cache.go)

The primary cache stores compiled bytecode for ASP page files on disk. The file cache operates in two tiers:

- **Memory tier (Tier 1):** An LRU in-memory cache keyed by the source file path. Fastest lookup; no disk I/O.
- **Disk tier (Tier 2):** A binary cache file stored under `temp/cache`. File content is hashed with xxhash (64-bit). On cache miss the server compiles the script and saves the compiled payload to disk.

The disk format includes a `G3AXON` magic header, a binary version number, the source file modification timestamp, and the full compiled program payload. On load, the cached entry's modification time is verified against the source file. Stale entries are rejected and the script is recompiled.

A file watcher (`fsnotify`) monitors the web root for changes and automatically invalidates in-memory cache entries when a source file is modified. This means page edits are picked up immediately without restarting the server.


> **Note**: AxonASP's caching mechanism uses a file watcher to detect file modifications. This can occasionally result in a stale cache or slow updates, causing issues in development environments. To prevent this, you might want to disable the cache during ASP development by setting bytecode_caching_enabled=disabled.

**Cache modes:**

| Mode | Memory tier | Disk tier |
|------|:-----------:|:---------:|
| enabled | Yes | Yes |
| memory-only | Yes | No |
| disk-only | No | Yes |
| disabled | No | No |

Configure the mode in `config/axonasp.toml`:

```toml
[global]
bytecode_caching_enabled = "enabled"
cache_max_size_mb = 128
```

### 2. Eval Cache (axonvm/eval_cache.go)

The Eval cache stores compiled bytecode for expressions passed to the VBScript `Eval()` built-in function. Expressions are keyed by a 64-bit FNV-1a hash derived from:

- The expression text
- The current `Option Compare` mode
- A hash of the active global symbol names
- A hash of the local procedure scope symbol names

Cache capacity is fixed at 512 entries (LRU eviction). An expression that produces the same hash as a cached entry is reused directly without recompilation. The cache is process-wide and shared across all concurrent requests.

### 3. Dynamic Execute Cache (axonvm/dynamic_exec_cache.go)

The dynamic execute cache stores compiled bytecode for code fragments passed to the VBScript `Execute()` and `ExecuteGlobal()` built-in functions. The cache key is a 64-bit FNV-1a hash derived from:

- The code fragment text
- The execution kind (Execute vs ExecuteGlobal)
- The current `Option Compare` and `Option Explicit` modes
- A hash of the active global symbol names and local scope (for Execute)
- A snapshot of active class version and class name (for Execute)

Cache capacity is fixed at 512 entries (LRU eviction). On a cache hit, the compiler symbol snapshot is also restored to keep class and global state consistent.

### 4. Execute File Cache (axonvm/execute_file_cache.go)

A dedicated in-memory-only cache with a capacity of 64 entries handles child pages invoked via `Server.Execute`. This cache uses the same `ScriptCache` subsystem as the primary script cache but operates in `memory-only` mode without disk persistence.

## Configuration Reference

| Setting | Section | Type | Default | Description |
|---------|---------|------|---------|-------------|
| bytecode_caching_enabled | [global] | string | "enabled" | Cache mode: enabled / memory-only / disk-only / disabled |
| cache_max_size_mb | [global] | integer | 128 | Maximum size of the in-memory script cache in megabytes |
| clean_cache_on_startup | [global] | boolean | true | Clear compiled bytecode cache files at server startup |

## Remarks

- Caching is fully transparent. No code changes are needed in ASP scripts to benefit from it.
- The disk cache is stored in `temp/cache`. This directory is created automatically if it does not exist.
- On clean server startup with `clean_cache_on_startup = true`, all existing disk cache files are deleted and a fresh compilation pass occurs on the first request to each script.
- The Eval and dynamic execute caches are always active when the VM is running, regardless of the `bytecode_caching_enabled` setting. Their size is fixed in code and is not controlled by `cache_max_size_mb`.
- The file watcher invalidates only memory-tier entries. Disk cache files are invalidated by the modification time check on load.
- Include file dependencies are tracked in the disk cache. If an included `.asp` or `.inc` file changes, the cache entry for the parent page is also invalidated.
- The binary cache format includes a version number. When the engine is upgraded and the binary version changes, all existing disk cache files are automatically rejected and recompiled.
