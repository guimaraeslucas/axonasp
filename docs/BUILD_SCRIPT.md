# Build Script Documentation

## Overview

The `build.ps1` PowerShell script provides an automated way to build both **AxonASP** executables:
- **axonasp.exe** - Standalone HTTP server (main application)
- **axonaspcgi.exe** - FastCGI server for integration with nginx/Apache/IIS

The script supports cross-platform compilation for Windows, Linux, and macOS with multiple architectures, optional cleaning, and testing capabilities.

## Prerequisites

- **Go 1.25+** installed and available in PATH
- **PowerShell 5.1+** (Windows) or **PowerShell Core 7+** (cross-platform)
- Write permissions in the project directory

## Basic Usage

### Build for Current Platform

```powershell
.\build.ps1
```

This builds both executables for your current operating system and architecture in the `build/` directory.

### Build for Specific Platform

```powershell
# Windows 64-bit
.\build.ps1 -Platform windows -Architecture amd64

# Linux 64-bit
.\build.ps1 -Platform linux -Architecture amd64

# macOS ARM64 (Apple Silicon)
.\build.ps1 -Platform darwin -Architecture arm64
```

### Build for All Platforms

```powershell
.\build.ps1 -Platform all
```

Builds executables for all supported platform/architecture combinations:
- Windows: amd64, 386, arm64
- Linux: amd64, 386, arm64
- macOS: amd64, arm64

## Parameters

### -Platform

Specifies the target operating system.

**Values:**
- `windows` (default) - Build for Windows
- `linux` - Build for Linux
- `darwin` - Build for macOS
- `all` - Build for all platforms

**Example:**
```powershell
.\build.ps1 -Platform linux
```

### -Architecture

Specifies the target CPU architecture.

**Values:**
- `amd64` (default) - 64-bit x86 architecture
- `386` - 32-bit x86 architecture
- `arm64` - 64-bit ARM architecture (Apple Silicon, ARM servers)

**Example:**
```powershell
.\build.ps1 -Architecture arm64
```

### -Clean

Removes the `build/` directory before building.

**Example:**
```powershell
.\build.ps1 -Clean
```

**Use cases:**
- Clean slate before release builds
- Remove old builds from different platforms
- Troubleshoot build issues

### -Test

Runs Go tests before building.

**Example:**
```powershell
.\build.ps1 -Test
```

**What it tests:**
- Unit tests in `asp/` package
- Unit tests in `server/` package
- Unit tests in `vbscript/` package

If any tests fail, the build process stops.

## Common Scenarios

### Development Build

Quick build for testing on your current machine:

```powershell
.\build.ps1
```

### Production Build

Clean build with tests for production deployment:

```powershell
.\build.ps1 -Clean -Test
```

### Deploy to Linux Server

Build Linux binary from Windows development machine:

```powershell
.\build.ps1 -Platform linux -Architecture amd64
```

Then transfer `build/linux-amd64/axonasp` to your Linux server.

### Deploy to macOS (Apple Silicon)

Build ARM64 binary for modern Macs:

```powershell
.\build.ps1 -Platform darwin -Architecture arm64
```

### Build Release Packages

Create builds for all platforms (e.g., for GitHub releases):

```powershell
.\build.ps1 -Platform all -Clean -Test
```

This creates organized subdirectories:
```
build/
├── windows-amd64/
│   ├── axonasp.exe
│   └── axonaspcgi.exe
├── windows-386/
├── windows-arm64/
├── linux-amd64/
│   ├── axonasp
│   └── axonaspcgi
├── linux-386/
├── linux-arm64/
├── darwin-amd64/
└── darwin-arm64/
```

### 32-bit Windows Build

For older Windows systems:

```powershell
.\build.ps1 -Platform windows -Architecture 386
```

### ARM Server Deployment

Build for ARM-based servers (e.g., Raspberry Pi, AWS Graviton):

```powershell
.\build.ps1 -Platform linux -Architecture arm64
```

## Output

### Build Directory Structure

The script creates a `build/` directory with the following structure:

```
build/
└── <platform>-<arch>/
    ├── axonasp[.exe]
    └── axonaspcgi[.exe]
```

**Examples:**
- `build/windows-amd64/axonasp.exe`
- `build/linux-amd64/axonasp` (no .exe extension)
- `build/darwin-arm64/axonaspcgi`

### Console Output

The script provides color-coded feedback:

```
==================================
AxonASP Build Script
==================================
Target Platform: windows
Architecture: amd64
Clean Build: No
Run Tests: No
----------------------------------

Building axonasp for windows/amd64...
SUCCESS: axonasp.exe (24.97 MB)

Building axonaspcgi for windows/amd64...
SUCCESS: axonaspcgi.exe (23.80 MB)

==================================
BUILD SUCCESSFUL!
==================================
Build artifacts: build\windows-amd64
```

## Troubleshooting

### "Go is not recognized"

**Error:** `go : The term 'go' is not recognized...`

**Solution:** Install Go from https://go.dev/dl/ and ensure it's in your PATH.

### "Permission Denied"

**Error:** Cannot create/write to build directory.

**Solution:** Run PowerShell as Administrator or check file permissions.

### Build Fails with Module Errors

**Error:** `cannot find module...`

**Solution:** Ensure you're in the project root and `go.mod` is present:

```powershell
cd e:\lucas\Desktop\Sites\LGGM-TCP\modules\image\ASP\axonasp
.\build.ps1
```

### Tests Fail

**Error:** Build stops during test phase.

**Solution:** Fix test failures first, or build without tests:

```powershell
.\build.ps1 -Test:$false
```

### Old Builds Not Cleaned

**Problem:** Previous builds remain in directory.

**Solution:** Use `-Clean` flag:

```powershell
.\build.ps1 -Clean
```

### Wrong Architecture

**Problem:** Binary doesn't run on target machine.

**Solution:** Verify target architecture:

- Most modern systems: `amd64`
- Older 32-bit Windows: `386`
- Apple Silicon Macs: `arm64` with `darwin`
- Raspberry Pi 4+: `arm64` with `linux`

## Advanced Usage

### Combine Multiple Parameters

```powershell
# Clean build with tests for Linux ARM64
.\build.ps1 -Platform linux -Architecture arm64 -Clean -Test

# Build all platforms with clean slate
.\build.ps1 -Platform all -Clean
```

### Check Build Script Help

```powershell
Get-Help .\build.ps1
Get-Help .\build.ps1 -Detailed
Get-Help .\build.ps1 -Examples
```

### Manual Build (Without Script)

If you prefer manual control:

```powershell
# Main executable
go build -o axonasp.exe .

# FastCGI executable
go build -o axonaspcgi.exe ./axonaspcgi

# Cross-compile for Linux
$env:GOOS="linux"; $env:GOARCH="amd64"; go build -o axonasp ./
```

## Integration with CI/CD

### GitHub Actions Example

```yaml
- name: Build AxonASP
  run: |
    pwsh -File build.ps1 -Platform all -Clean -Test
    
- name: Upload Artifacts
  uses: actions/upload-artifact@v3
  with:
    name: axonasp-builds
    path: build/
```

### GitLab CI Example

```yaml
build:
  script:
    - pwsh -File build.ps1 -Platform all -Clean -Test
  artifacts:
    paths:
      - build/
```

## Binary Sizes

Typical build sizes (may vary by version):

| Platform | Architecture | axonasp | axonaspcgi |
|----------|-------------|---------|------------|
| Windows | amd64 | ~25 MB | ~24 MB |
| Linux | amd64 | ~23 MB | ~22 MB |
| macOS | arm64 | ~22 MB | ~21 MB |

## License

This build script is part of the AxonASP project and follows the same license terms.

## See Also

- [FastCGI Mode Documentation](FASTCGI_MODE.md)
- [FastCGI Quick Start](FASTCGI_QUICKSTART.md)
- [README.md](../README.md)
